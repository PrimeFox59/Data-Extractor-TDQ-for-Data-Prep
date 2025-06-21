Sub ListFileDenganTable()
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim folderPath As String
    Dim wsData As Worksheet, wsSetting As Worksheet
    Dim tbl As ListObject
    Dim existingFiles As Object
    Dim newFileCount As Long
    Dim excludeFolders As Collection
    Dim i As Long
    Dim excludePath As String
    Dim row As ListRow
    Dim existingPath As String

    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsData = ThisWorkbook.Sheets("Data Extract")
    Set wsSetting = ThisWorkbook.Sheets("Setting")

    ' Cek dan ambil path utama
    folderPath = wsSetting.Range("B2").value
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder tidak ditemukan: " & folderPath, vbCritical
        Exit Sub
    End If

    ' Cek dan ambil Table
    On Error Resume Next
    Set tbl = wsData.ListObjects("tblFileList")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Table 'tblFileList' tidak ditemukan di Sheet 'Data Extract'", vbExclamation
        Exit Sub
    End If

    ' Ambil daftar folder yang dikecualikan
    Set excludeFolders = New Collection
    i = 5
    Do While wsSetting.Cells(i, 2).value <> ""
        excludePath = wsSetting.Cells(i, 2).value
        If Right(excludePath, 1) <> "\" Then excludePath = excludePath & "\"
        excludeFolders.Add LCase(excludePath)
        i = i + 1
    Loop

    ' Buat dictionary dari file yang sudah ada di table
    Set existingFiles = CreateObject("Scripting.Dictionary")
    If Not tbl.DataBodyRange Is Nothing Then
        For Each row In tbl.ListRows
            On Error Resume Next
            existingPath = row.Range(1, 2).Hyperlinks(1).Address
            existingPath = Replace(existingPath, "file://", "")
            existingPath = Replace(existingPath, "/", "\")
            If Len(Dir(existingPath)) > 0 Then
                existingFiles.Add existingPath, True
            End If
            On Error GoTo 0
        Next row
    End If

    ' Tambahkan file baru ke tabel
    newFileCount = 0
    Call ScanFolderToTable(fso.GetFolder(folderPath), tbl, existingFiles, newFileCount, excludeFolders)

    MsgBox "Selesai update list!" & vbCrLf & _
           "File baru ditambahkan: " & newFileCount, vbInformation

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
Sub ScanFolderToTable(ByVal folder As Object, ByRef tbl As ListObject, _
                      ByRef existingFiles As Object, _
                      ByRef newFileCount As Long, _
                      ByRef excludeFolders As Collection)

    Dim fileItem As Object, subFolder As Object
    Dim fullPath As String, fileNameOnly As String
    Dim folderPath As String
    Dim skipFolder As Boolean
    Dim excludePath As Variant
    Dim newRow As ListRow

    folderPath = LCase(folder.Path) & "\"

    ' Skip folder yang di-exclude
    skipFolder = False
    For Each excludePath In excludeFolders
        If folderPath = excludePath Then
            skipFolder = True
            Exit For
        End If
    Next excludePath
    If skipFolder Then Exit Sub

    ' Tambahkan file
    For Each fileItem In folder.Files
        fullPath = fileItem.Path
        fileNameOnly = fileItem.Name

        ' Hanya file .xls/.xlsx dan tidak diawali ~$
        If Not existingFiles.Exists(fullPath) And _
           Left(fileNameOnly, 2) <> "~$" And _
           (LCase(Right(fileNameOnly, 4)) = ".xls" Or LCase(Right(fileNameOnly, 5)) = ".xlsx") Then

            existingFiles.Add fullPath, True
            Set newRow = tbl.ListRows.Add
            newRow.Range(1, 1).value = tbl.ListRows.Count ' Nomor
            tbl.Parent.Hyperlinks.Add _
                Anchor:=newRow.Range(1, 2), _
                Address:="file://" & fullPath, _
                TextToDisplay:=fileNameOnly
            newFileCount = newFileCount + 1
        End If
    Next fileItem

    ' Lanjut ke subfolder
    For Each subFolder In folder.SubFolders
        Call ScanFolderToTable(subFolder, tbl, existingFiles, newFileCount, excludeFolders)
    Next subFolder
End Sub

