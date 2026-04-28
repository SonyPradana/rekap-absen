' ============================================================================
' TEST SUITE untuk CreateSheet.bas
' Jalankan tests di Immediate Window: Ctrl+G
' ============================================================================

' Helper function untuk test (VBA tidak punya IsNumeric built-in)
Function IsNumeric(ByVal val As Variant) As Boolean
    On Error Resume Next
    IsNumeric = IsNumeric(CDbl(val))
    On Error GoTo 0
End Function

' ============================================================================
' TEST 1: BersihkanNamaSheet
' ============================================================================
Sub TEST_BersihkanNamaSheet()
    Debug.Print "==== TEST: BersihkanNamaSheet ===="
    
    Dim testCases As Variant
    testCases = Array( _
        Array("Nama/Invalid*Sheet?", "NamaInvalidSheet"), _
        Array("Test[123]:Nama", "Test123Nama"), _
        Array("Normal_Name", "Normal_Name"), _
        Array("VeryLongNameThatExceeds30CharactersAndShouldBeTruncated", Left("VeryLongNameThatExceeds30CharactersAndShouldBeTruncated", 30)), _
        Array("*/*/?/*:*[*]*", ""), _
        Array("John\Do*e?File", "JohnDoeFile"), _
        Array("", "") _
    )
    
    Dim i As Long, result As String, expected As String
    Dim passed As Long, failed As Long
    passed = 0: failed = 0
    
    For i = LBound(testCases) To UBound(testCases)
        result = BersihkanNamaSheet(testCases(i)(0))
        expected = testCases(i)(1)
        
        If result = expected Then
            Debug.Print "✓ PASS: '" & testCases(i)(0) & "' -> '" & result & "'"
            passed = passed + 1
        Else
            Debug.Print "✗ FAIL: '" & testCases(i)(0) & "' Expected: '" & expected & "' Got: '" & result & "'"
            failed = failed + 1
        End If
    Next i
    
    Debug.Print "Passed: " & passed & ", Failed: " & failed
    Debug.Print vbNewLine
End Sub

' ============================================================================
' TEST 2: FormatDurasi
' ============================================================================
Sub TEST_FormatDurasi()
    Debug.Print "==== TEST: FormatDurasi ===="
    
    Dim testCases As Variant
    testCases = Array( _
        Array(45, "45s"), _
        Array(60, "1m 0s"), _
        Array(65, "1m 5s"), _
        Array(3600, "1j 0m 0s"), _
        Array(3661, "1j 1m 1s"), _
        Array(7200, "2j 0m 0s"), _
        Array(7325, "2j 2m 5s"), _
        Array(0, "0s"), _
        Array(3599, "59m 59s"), _
        Array(86400, "24j 0m 0s") _
    )
    
    Dim i As Long, result As String, expected As String
    Dim passed As Long, failed As Long
    passed = 0: failed = 0
    
    For i = LBound(testCases) To UBound(testCases)
        result = FormatDurasi(testCases(i)(0))
        expected = testCases(i)(1)
        
        If result = expected Then
            Debug.Print "✓ PASS: " & testCases(i)(0) & "s -> '" & result & "'"
            passed = passed + 1
        Else
            Debug.Print "✗ FAIL: " & testCases(i)(0) & "s Expected: '" & expected & "' Got: '" & result & "'"
            failed = failed + 1
        End If
    Next i
    
    Debug.Print "Passed: " & passed & ", Failed: " & failed
    Debug.Print vbNewLine
End Sub

' ============================================================================
' TEST 3: HitungKeteranganWaktu - Status IN
' ============================================================================
Sub TEST_HitungKeteranganWaktu_IN()
    Debug.Print "==== TEST: HitungKeteranganWaktu (Status IN) ===="
    
    Dim jamLog As Date, tgl As Date, result As String
    Dim passed As Long, failed As Long
    passed = 0: failed = 0
    
    ' Test Case 1: Senin (hariID=1), ON TIME
    jamLog = TimeValue("07:10:00")
    tgl = DateSerial(2026, 4, 27) ' Senin
    result = HitungKeteranganWaktu(jamLog, tgl, "IN")
    If result = "ok" Then
        Debug.Print "✓ PASS: Senin, 07:10 -> 'ok'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Senin, 07:10 Expected: 'ok' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 2: Senin, TERLAMBAT 5 menit
    jamLog = TimeValue("07:20:00")
    tgl = DateSerial(2026, 4, 27)
    result = HitungKeteranganWaktu(jamLog, tgl, "IN")
    If result = "Telat: 5m 0s" Then
        Debug.Print "✓ PASS: Senin, 07:20 -> 'Telat: 5m 0s'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Senin, 07:20 Expected: 'Telat: 5m 0s' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 3: Jumat (hariID=5), ON TIME (07:00)
    jamLog = TimeValue("07:00:00")
    tgl = DateSerial(2026, 5, 1) ' Jumat
    result = HitungKeteranganWaktu(jamLog, tgl, "IN")
    If result = "ok" Then
        Debug.Print "✓ PASS: Jumat, 07:00 -> 'ok'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Jumat, 07:00 Expected: 'ok' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 4: Jumat, TERLAMBAT 10 menit
    jamLog = TimeValue("07:10:00")
    tgl = DateSerial(2026, 5, 1)
    result = HitungKeteranganWaktu(jamLog, tgl, "IN")
    If result = "Telat: 10m 0s" Then
        Debug.Print "✓ PASS: Jumat, 07:10 -> 'Telat: 10m 0s'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Jumat, 07:10 Expected: 'Telat: 10m 0s' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 5: ERROR (>3 jam terlambat)
    jamLog = TimeValue("10:30:00")
    tgl = DateSerial(2026, 4, 27)
    result = HitungKeteranganWaktu(jamLog, tgl, "IN")
    If result = "Error: >3j" Then
        Debug.Print "✓ PASS: Senin, 10:30 -> 'Error: >3j'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Senin, 10:30 Expected: 'Error: >3j' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 6: Status dengan lowercase (case insensitive)
    jamLog = TimeValue("07:05:00")
    tgl = DateSerial(2026, 4, 27)
    result = HitungKeteranganWaktu(jamLog, tgl, "in")
    If result = "Telat: 5m 0s" Then
        Debug.Print "✓ PASS: Status 'in' (lowercase) -> 'Telat: 5m 0s'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Status 'in' Expected: 'Telat: 5m 0s' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    Debug.Print "Passed: " & passed & ", Failed: " & failed
    Debug.Print vbNewLine
End Sub

' ============================================================================
' TEST 4: HitungKeteranganWaktu - Status OUT
' ============================================================================
Sub TEST_HitungKeteranganWaktu_OUT()
    Debug.Print "==== TEST: HitungKeteranganWaktu (Status OUT) ===="
    
    Dim jamLog As Date, tgl As Date, result As String
    Dim passed As Long, failed As Long
    passed = 0: failed = 0
    
    ' Test Case 1: Senin (hariID=1), ON TIME (14:00)
    jamLog = TimeValue("14:30:00")
    tgl = DateSerial(2026, 4, 27) ' Senin
    result = HitungKeteranganWaktu(jamLog, tgl, "OUT")
    If result = "ok" Then
        Debug.Print "✓ PASS: Senin, 14:30 (pulang tepat) -> 'ok'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Senin, 14:30 Expected: 'ok' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 2: Senin, PULANG AWAL 30 menit
    jamLog = TimeValue("13:30:00")
    tgl = DateSerial(2026, 4, 27)
    result = HitungKeteranganWaktu(jamLog, tgl, "OUT")
    If result = "Pulang Awal: 30m 0s" Then
        Debug.Print "✓ PASS: Senin, 13:30 (pulang awal) -> 'Pulang Awal: 30m 0s'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Senin, 13:30 Expected: 'Pulang Awal: 30m 0s' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 3: Jumat (hariID=5), ON TIME (11:30)
    jamLog = TimeValue("11:45:00")
    tgl = DateSerial(2026, 5, 1) ' Jumat
    result = HitungKeteranganWaktu(jamLog, tgl, "OUT")
    If result = "ok" Then
        Debug.Print "✓ PASS: Jumat, 11:45 (pulang tepat) -> 'ok'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Jumat, 11:45 Expected: 'ok' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 4: Jumat, PULANG AWAL
    jamLog = TimeValue("11:00:00")
    tgl = DateSerial(2026, 5, 1)
    result = HitungKeteranganWaktu(jamLog, tgl, "OUT")
    If result = "Pulang Awal: 30m 0s" Then
        Debug.Print "✓ PASS: Jumat, 11:00 -> 'Pulang Awal: 30m 0s'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Jumat, 11:00 Expected: 'Pulang Awal: 30m 0s' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 5: Sabtu (hariID=6), ON TIME (13:15)
    jamLog = TimeValue("13:30:00")
    tgl = DateSerial(2026, 4, 25) ' Sabtu
    result = HitungKeteranganWaktu(jamLog, tgl, "OUT")
    If result = "ok" Then
        Debug.Print "✓ PASS: Sabtu, 13:30 -> 'ok'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Sabtu, 13:30 Expected: 'ok' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    ' Test Case 6: ERROR (>3 jam pulang awal)
    jamLog = TimeValue("09:50:00")
    tgl = DateSerial(2026, 4, 27)
    result = HitungKeteranganWaktu(jamLog, tgl, "OUT")
    If result = "Error: >3j" Then
        Debug.Print "✓ PASS: Senin, 09:50 (>3j awal) -> 'Error: >3j'"
        passed = passed + 1
    Else
        Debug.Print "✗ FAIL: Senin, 09:50 Expected: 'Error: >3j' Got: '" & result & "'"
        failed = failed + 1
    End If
    
    Debug.Print "Passed: " & passed & ", Failed: " & failed
    Debug.Print vbNewLine
End Sub

' ============================================================================
' RUN ALL TESTS
' ============================================================================
Sub RUN_ALL_TESTS()
    Debug.Print "########################################"
    Debug.Print "#  TEST SUITE - CreateSheet.bas        #"
    Debug.Print "#  Tanggal Test: " & Format(Now, "dd/mm/yyyy HH:mm:ss") & "     #"
    Debug.Print "########################################"
    Debug.Print vbNewLine
    
    TEST_BersihkanNamaSheet
    TEST_FormatDurasi
    TEST_HitungKeteranganWaktu_IN
    TEST_HitungKeteranganWaktu_OUT
    
    Debug.Print "########################################"
    Debug.Print "#  SEMUA TEST SELESAI                  #"
    Debug.Print "########################################"
    MsgBox "Test selesai! Lihat Immediate Window (Ctrl+G) untuk hasil lengkap", vbInformation
End Sub
