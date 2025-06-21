Sub AmbilDataBiayaMultiSelect()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim cell As Range
    Dim selectedRange As Range
    Dim filePath As String
    Dim hyperlinkExists As Boolean
    
    Set ws = ActiveSheet
    Set selectedRange = Selection

    If selectedRange Is Nothing Then
        MsgBox "Tidak ada sel yang dipilih.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Tampilkan UserForm Loading
    Load ufmLoading
    ufmLoading.lblStatus.Caption = "? Memulai proses..."
    ufmLoading.Show vbModeless
    DoEvents

    For Each cell In selectedRange
        If cell.Column = 2 Then ' hanya proses kolom B
            ufmLoading.lblStatus.Caption = "? Memproses baris " & cell.row & ": " & cell.Text
            DoEvents
            
            hyperlinkExists = (cell.Hyperlinks.Count > 0)
            
            ' Cek apakah cell berisi hyperlink atau path file
            If hyperlinkExists Then
                filePath = cell.Hyperlinks(1).Address
                filePath = Replace(filePath, "/", "\")
                If Dir(filePath) = "" Then
                    filePath = cell.Hyperlinks(1).TextToDisplay
                    filePath = Replace(filePath, "/", "\")
                End If
                If Dir(filePath) = "" Then
                    cell.Hyperlinks.Delete
                    hyperlinkExists = False
                End If
            End If
            
            If Not hyperlinkExists And cell.value <> "" Then
                filePath = Replace(cell.value, "/", "\")
                If Dir(filePath) <> "" Then
                    ws.Hyperlinks.Add Anchor:=cell, Address:=filePath, TextToDisplay:=cell.value
                    hyperlinkExists = True
                End If
            End If
            
            If Not hyperlinkExists Then
                cell.Offset(0, 1).value = "No valid file reference"
                GoTo NextCell
            End If
            
            If Dir(filePath) = "" Then
                cell.Offset(0, 1).value = "File tidak ditemukan"
                GoTo NextCell
            End If
            
            Call AmbilDataBiayaPerSel(ws, cell)
        End If
NextCell:
    Next cell

    Unload ufmLoading
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "? Semua data selesai diproses.", vbInformation
    Exit Sub

ErrorHandler:
    Unload ufmLoading
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "? Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Terjadi pada prosedur AmbilDataBiayaMultiSelect", vbCritical
End Sub


Sub AmbilDataBiayaPerSel(ws As Worksheet, selectedCell As Range)
    On Error GoTo ErrorHandler
    Dim targetFilePath As String
    Dim targetFileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim revSheetName As String
    Dim highestRev As Integer
    Dim foundRev As Boolean
    Dim searchTerms As Variant
    Dim targetColumns As Variant
    Dim offsetRows As Variant
    Dim offsetCols As Variant
    Dim searchCell As Range
    Dim i As Integer
    Dim nilai As Variant
    Dim fileIsOpen As Boolean
    Dim r As Long
    Dim sheetNames As Variant
    Dim sheetIndex As Integer
    Dim dataFound As Boolean

    ' Simpan nomor baris yang sedang diproses
    r = selectedCell.row
    
    ' Ambil path dan nama file
    targetFilePath = selectedCell.Hyperlinks(1).Address
    targetFileName = selectedCell.Hyperlinks(1).TextToDisplay

    If Left(targetFilePath, 2) = "\\" Then
        targetFilePath = Replace(targetFilePath, "/", "\")
    End If

    If Dir(targetFilePath) = "" Then
        ws.Cells(r, "C").Resize(1, 20).value = "File tidak ditemukan"
        Exit Sub
    End If

    On Error Resume Next
    Set wbSource = Workbooks.Open(fileName:=targetFilePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
    If Err.Number <> 0 Then
        ws.Cells(r, "C").Resize(1, 20).value = "Gagal buka file"
        On Error GoTo ErrorHandler
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    fileIsOpen = True

    ' Array untuk menentukan urutan pencarian sheet
    sheetNames = Array("REV", "COST", "DATA", "INPUT", "CALCULATION", "SUMMARY")
    
    ' Pertama cari sheet dengan REV
    highestRev = -1
    foundRev = False
    
    For Each wsSource In wbSource.Worksheets
        Dim sheetName As String
        Dim revPattern As String
        Dim revPos As Integer
        Dim revNumber As Integer
        Dim revText As String
        
        sheetName = UCase(wsSource.Name)
        
        ' Cari posisi "REV" dengan berbagai variasi
        revPos = 0
        If InStr(1, sheetName, "REV#") > 0 Then
            revPos = InStr(1, sheetName, "REV#")
            revPattern = "REV#"
        ElseIf InStr(1, sheetName, "REV-#") > 0 Then
            revPos = InStr(1, sheetName, "REV-#")
            revPattern = "REV-#"
        ElseIf InStr(1, sheetName, "REV.#") > 0 Then
            revPos = InStr(1, sheetName, "REV.#")
            revPattern = "REV.#"
        ElseIf InStr(1, sheetName, "REV #") > 0 Then
            revPos = InStr(1, sheetName, "REV #")
            revPattern = "REV #"
        ElseIf InStr(1, sheetName, "REV-") > 0 Then
            revPos = InStr(1, sheetName, "REV-")
            revPattern = "REV-"
        ElseIf InStr(1, sheetName, "REV.") > 0 Then
            revPos = InStr(1, sheetName, "REV.")
            revPattern = "REV."
        ElseIf InStr(1, sheetName, "REV") > 0 Then
            revPos = InStr(1, sheetName, "REV")
            revPattern = "REV"
        End If
        
        ' Jika ditemukan pattern REV
        If revPos > 0 Then
            revText = Mid(wsSource.Name, revPos + Len(revPattern))
            revNumber = val(revText)
            
            ' Cari angka revisi tertinggi
            If revNumber > highestRev Then
                highestRev = revNumber
                revSheetName = wsSource.Name
                foundRev = True
            End If
        End If
    Next wsSource

    ' Jika REV ditemukan, cari data di sheet REV terlebih dahulu
    If foundRev Then
        Set wsSource = wbSource.Sheets(revSheetName)
        dataFound = AmbilDataDariSheet(ws, r, wsSource)
        
        ' Jika data ditemukan di sheet REV, lanjut ke file berikutnya
        If dataFound Then GoTo CleanUp
    End If
    
    ' Jika REV tidak ditemukan atau data tidak lengkap di sheet REV,
    ' cari di sheet lainnya sesuai urutan prioritas
    For sheetIndex = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set wsSource = Nothing
        Set wsSource = wbSource.Sheets(sheetNames(sheetIndex))
        On Error GoTo ErrorHandler
        
        If Not wsSource Is Nothing Then
            dataFound = AmbilDataDariSheet(ws, r, wsSource)
            
            ' Jika data ditemukan, keluar dari loop
            If dataFound Then Exit For
        End If
    Next sheetIndex
    
    ' Jika masih belum ditemukan, cari di semua sheet
    If Not dataFound Then
        For Each wsSource In wbSource.Worksheets
            ' Skip sheet yang sudah dicari sebelumnya
            If wsSource.Name <> revSheetName And _
               Not IsInArray(wsSource.Name, sheetNames) Then
                dataFound = AmbilDataDariSheet(ws, r, wsSource)
                
                ' Jika data ditemukan, keluar dari loop
                If dataFound Then Exit For
            End If
        Next wsSource
    End If
    
    ' Jika data tetap tidak ditemukan, beri tanda N/A
    If Not dataFound Then
        ws.Cells(r, "C").Resize(1, 20).value = "N/A"
    End If

CleanUp:
    If fileIsOpen Then
        On Error Resume Next
        wbSource.Close SaveChanges:=False
        On Error GoTo ErrorHandler
    End If
    Exit Sub

ErrorHandler:
    ' Jika terjadi error, isi semua kolom dengan "Error"
    ws.Cells(r, "C").Resize(1, 20).value = "Error"
    Resume CleanUp
End Sub

Function AmbilDataDariSheet(ws As Worksheet, r As Long, wsSource As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    Dim searchTerms As Variant
    Dim targetColumns As Variant
    Dim offsetRows As Variant
    Dim offsetCols As Variant
    Dim searchCell As Range
    Dim i As Integer
    Dim nilai As Variant
    Dim dataFound As Boolean
    Dim foundCount As Integer
    
    searchTerms = Array("Insert", "Oil Cost", "Jig, Tool & Maint", "Man Power Direct", "Electricity", "Depresiasi Mesin", _
                        "TOTAL COST MC", "Jumlah mesin", "Casting Weight", "Man Power", _
                        "Need Line", "MCT TOTAL", "PRODUKTIFITAS LINE", "TOTAL PROFIT AMOUNT")
    targetColumns = Array("G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "T", "Q", "R", "S")
    offsetRows = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1)
    offsetCols = Array(1, 1, 1, 1, 1, 1, 1, 0, 1, 1, 3, 2, 0, 0)
    
    foundCount = 0
    
    ' Cari data utama
    For i = LBound(searchTerms) To UBound(searchTerms)
        Dim specialRange As Range
        Select Case searchTerms(i)
            Case "Casting Weight", "Man Power"
                Set specialRange = wsSource.Range("S:S")
                Set searchCell = specialRange.Find(What:=searchTerms(i), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
            Case "Need Line"
                Set specialRange = wsSource.Range("I:I")
                Set searchCell = specialRange.Find(What:="Need Line", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
            Case Else
                Set searchCell = wsSource.Cells.Find(What:=searchTerms(i), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        End Select

        If Not searchCell Is Nothing Then
            If searchTerms(i) = "Jumlah mesin" Then
                Dim totalMesin As Double
                Dim rowCounter As Long
                totalMesin = 0
                rowCounter = 1
                Do While searchCell.Offset(rowCounter, 0).value <> ""
                    If IsNumeric(searchCell.Offset(rowCounter, 0).value) Then
                        totalMesin = totalMesin + searchCell.Offset(rowCounter, 0).value
                    End If
                    rowCounter = rowCounter + 1
                Loop
                ws.Cells(r, Range(targetColumns(i) & "1").Column).value = totalMesin
                foundCount = foundCount + 1
            Else
                On Error Resume Next
                nilai = searchCell.Offset(offsetRows(i), offsetCols(i)).value
                On Error GoTo ErrorHandler
                If Err.Number = 0 Then
                    ws.Cells(r, Range(targetColumns(i) & "1").Column).value = IIf(IsNumeric(nilai), CDbl(nilai), nilai)
                    foundCount = foundCount + 1
                Else
                    ws.Cells(r, Range(targetColumns(i) & "1").Column).value = "Not found"
                End If
            End If
        Else
            ws.Cells(r, Range(targetColumns(i) & "1").Column).value = "Not found"
        End If
    Next i

    ' Cari meta info: Item, Part No, Cust, Date
    Dim metaLabels As Variant, metaTargets As Variant
    metaLabels = Array("Item", "Part No", "Cust", "Date")
    metaTargets = Array("C", "D", "E", "F")

    Dim j As Integer
    For j = LBound(metaLabels) To UBound(metaLabels)
        Set searchCell = wsSource.Cells.Find(What:=metaLabels(j), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not searchCell Is Nothing Then
            On Error Resume Next
            nilai = searchCell.Offset(0, 1).value
            On Error GoTo ErrorHandler
            If Err.Number = 0 Then
                If Not IsEmpty(nilai) Then
                    If Left(nilai, 2) = ": " Then nilai = Mid(nilai, 3)
                    ws.Cells(r, Range(metaTargets(j) & "1").Column).value = nilai
                    foundCount = foundCount + 1
                Else
                    ws.Cells(r, Range(metaTargets(j) & "1").Column).value = "Null"
                End If
            Else
                ws.Cells(r, Range(metaTargets(j) & "1").Column).value = "Not found"
            End If
        Else
            ws.Cells(r, Range(metaTargets(j) & "1").Column).value = "Not found"
        End If
    Next j
    
    ' Jika ditemukan minimal 50% data, anggap berhasil
    If foundCount >= (UBound(searchTerms) + UBound(metaLabels) + 2) / 2 Then
        AmbilDataDariSheet = True
    Else
        AmbilDataDariSheet = False
    End If
    
    Exit Function
    
ErrorHandler:
    AmbilDataDariSheet = False
End Function

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If UCase(arr(i)) = UCase(val) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function


