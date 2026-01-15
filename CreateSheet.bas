' Menghilangkan karakter terlarang untuk nama Sheet
Function BersihkanNamaSheet(ByVal nama As String) As String
    Dim invalidChars As Variant, i As Integer
    invalidChars = Array("/", "\", "*", "?", ":", "[", "]")
    nama = Left(nama, 30)
    For i = LBound(invalidChars) To UBound(invalidChars)
        nama = Replace(nama, invalidChars(i), "")
    Next i
    BersihkanNamaSheet = nama
End Function

' Mengonversi detik ke format teks (0j 0m 0s)
Function FormatDurasi(ByVal totalDetik As Long) As String
    Dim h As Long, m As Long, s As Long
    h = totalDetik \ 3600
    m = (totalDetik Mod 3600) \ 60
    s = totalDetik Mod 60
    
    If h > 0 Then
        FormatDurasi = h & "j " & m & "m " & s & "s"
    ElseIf m > 0 Then
        FormatDurasi = m & "m " & s & "s"
    Else
        FormatDurasi = s & "s"
    End If
End Function

Function HitungKeteranganWaktu(ByVal jamLog As Date, ByVal tgl As Date, ByVal status As String) As String
    Dim jamAcuan As Date
    Dim hariID As Integer
    Dim selisih As Long
    Const BATAS_ERROR As Long = 10800 ' 3 Jam
    
    hariID = Weekday(tgl, vbMonday)
    status = UCase(status)
    
    ' --- LOGIKA MASUK (IN) ---
    If status Like "*IN*" Then
        If hariID = 5 Then
            jamAcuan = TimeValue("07:00:00")
        Else
            jamAcuan = TimeValue("07:15:00")
        End If
        
        If jamLog > jamAcuan Then
            selisih = DateDiff("s", jamAcuan, jamLog)
            If selisih > BATAS_ERROR Then
                HitungKeteranganWaktu = "Error: >3j"
            Else
                HitungKeteranganWaktu = "Telat: " & FormatDurasi(selisih)
            End If
        Else
            HitungKeteranganWaktu = "ok"
        End If
        
    ' --- LOGIKA PULANG (OUT) ---
    ElseIf status Like "*OUT*" Then
        ' Menentukan Jam Acuan Berdasarkan Hari
        Select Case hariID
            Case 5 ' Jumat
                jamAcuan = TimeValue("11:30:00")
            Case 6 ' Sabtu
                jamAcuan = TimeValue("13:15:00")
            Case Else ' Senin-Kamis & Minggu
                jamAcuan = TimeValue("14:00:00")
        End Select
        
        If jamLog < jamAcuan Then
            selisih = DateDiff("s", jamLog, jamAcuan)
            If selisih > BATAS_ERROR Then
                HitungKeteranganWaktu = "Error: >3j"
            Else
                HitungKeteranganWaktu = "Pulang Awal: " & FormatDurasi(selisih)
            End If
        Else
            HitungKeteranganWaktu = "ok"
        End If
    End If
End Function

Sub PisahAbsensi()
    Dim wsMaster As Worksheet, wsNew As Worksheet
    Dim dict As Object
    Dim lastRow As Long, i As Long, r As Long
    Dim currentKey As Variant
    
    Set wsMaster = ThisWorkbook.ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, 2).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    ' 1. Mapping Karyawan
    For i = 2 To lastRow
        Dim idBersih As String
        idBersih = BersihkanNamaSheet(wsMaster.Cells(i, 2).Value)
        If Not dict.Exists(idBersih) Then dict.Add idBersih, wsMaster.Cells(i, 2).Value
    Next i
    
    ' 2. Eksekusi Per Karyawan
    For Each currentKey In dict.Keys
        ' Kelola Sheet
        On Error Resume Next
        Application.DisplayAlerts = False: Sheets(CStr(currentKey)).Delete: Application.DisplayAlerts = True
        On Error GoTo 0
        
        Set wsNew = Sheets.Add(After:=Sheets(Sheets.Count))
        wsNew.Name = CStr(currentKey)
        
        ' Filter & Copy
        wsMaster.UsedRange.AutoFilter Field:=2, Criteria1:=dict(currentKey)
        wsMaster.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=wsNew.Range("A1")
        wsNew.Cells(1, 6).Value = "Keterangan Waktu"
        
        ' Perhitungan Baris
        Dim rowCount As Long: rowCount = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
        For r = 2 To rowCount
            If IsDate(wsNew.Cells(r, 1)) And IsNumeric(wsNew.Cells(r, 3)) Then
                wsNew.Cells(r, 6).Value = HitungKeteranganWaktu(wsNew.Cells(r, 3).Value, _
                                                               wsNew.Cells(r, 1).Value, _
                                                               wsNew.Cells(r, 4).Value)
            End If
        Next r
        
        ' Formatting
        wsNew.UsedRange.Columns.AutoFit
        wsNew.UsedRange.Borders.LineStyle = xlContinuous
    Next currentKey
    
    wsMaster.AutoFilterMode = False
    Application.ScreenUpdating = True
    MsgBox "Selesai!", vbInformation
End Sub
