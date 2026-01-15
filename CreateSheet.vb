Sub PisahAbsensi()
    Dim wsMaster As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, i As Long, r As Long
    Dim dict As Object, key As Variant, fullID As String
    Dim tgl As Date, jamLog As Date, jamAcuan As Date
    Dim selisih As Long, h As Long, m As Long, s As Long
    Dim teksDurasi As String, hariID As Integer

    Set wsMaster = ThisWorkbook.ActiveSheet
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, 2).End(xlUp).Row
    Set dict = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    
    ' 1. Ambil daftar unik ID-Nama
    For i = 2 To lastRow
        fullID = Left(wsMaster.Cells(i, 3).Value, 31)
        fullID = Replace(Replace(Replace(Replace(fullID, "/", ""), "*", ""), "?", ""), ":", "")
        If Not dict.Exists(fullID) Then dict.Add fullID, wsMaster.Cells(i, 2).Value
    Next i
    
    ' 2. Proses tiap karyawan
    For Each key In dict.Keys
        On Error Resume Next
        Application.DisplayAlerts = False: Sheets(CStr(key)).Delete: Application.DisplayAlerts = True
        On Error GoTo 0
        
        Set wsNew = Sheets.Add(After:=Sheets(Sheets.Count))
        wsNew.Name = CStr(key)
        
        wsMaster.UsedRange.AutoFilter Field:=2, Criteria1:=dict(key)
        wsMaster.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=wsNew.Range("A1")
        wsNew.Columns("B").Delete ' Hapus ID Number
        
        wsNew.Cells(1, 6).Value = "Keterangan Waktu"
        lastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
        
        ' 3. Loop perhitungan per baris
        For r = 2 To lastRow
            tgl = wsNew.Cells(r, 1).Value
            jamLog = wsNew.Cells(r, 3).Value
            hariID = Weekday(tgl, vbMonday) ' 1=Senin, 5=Jumat, 6=Sabtu
            
            ' --- LOGIKA MASUK (IN) ---
            If UCase(wsNew.Cells(r, 4).Value) Like "*IN*" Then
                If hariID = 5 Then jamAcuan = TimeValue("07:00:00") Else jamAcuan = TimeValue("07:15:00")
                
                If jamLog > jamAcuan Then
                    selisih = DateDiff("s", jamAcuan, jamLog)
                    wsNew.Cells(r, 6).Value = "Telat: " & FormatDurasi(selisih)
                Else
                    wsNew.Cells(r, 6).Value = "ok"
                End If
                
            ' --- LOGIKA PULANG (OUT) ---
            ElseIf UCase(wsNew.Cells(r, 4).Value) Like "*OUT*" Then
                If hariID = 5 Then ' Jumat
                    jamAcuan = TimeValue("11:30:00")
                ElseIf hariID = 6 Then ' Sabtu
                    jamAcuan = TimeValue("13:15:00")
                Else ' Senin - Kamis & Minggu
                    jamAcuan = TimeValue("14:00:00")
                End If
                
                If jamLog < jamAcuan Then
                    selisih = DateDiff("s", jamLog, jamAcuan)
                    wsNew.Cells(r, 6).Value = "Pulang Awal: " & FormatDurasi(selisih)
                Else
                    wsNew.Cells(r, 6).Value = "ok"
                End If
            End If
        Next r
        
        ' Rapikan Format
        With wsNew.UsedRange
            .WrapText = False
            .Columns.AutoFit
            .Borders.LineStyle = xlContinuous 
        End With
    Next key
    
    wsMaster.AutoFilterMode = False
    wsMaster.Activate
    Application.ScreenUpdating = True
    MsgBox "Proses Selesai!", vbInformation
End Sub

' Fungsi pendukung asli
Function FormatDurasi(totalDetik As Long) As String
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
