Attribute VB_Name = "Módulo1"
Public Sub pegaTodosArquivos()

    Dim caminho As String
    Dim result As Variant
    
    caminho = "C:\Users\Gabriel\Documents"
    
    '/C -> termina; /K ->Permace aberto
    result = Split(CreateObject("wscript.shell").Exec("cmd /c dir """ & caminho & """ /b/s").StdOut.ReadAll, vbCrLf)
    ThisWorkbook.Sheets(1).Cells.ClearContents
    
    Call tlhAsciiToUtf8("C:\Users\Gabriel\Documents\Scanned Documents\Digitaliza‡Æo de Boas-vindas.jpg")
    
    ThisWorkbook.Sheets(1).Range("A2").Resize(UBound(result)).Value = Application.WorksheetFunction.Transpose(result)
    ThisWorkbook.Sheets(1).Range("A1") = "Caminho"
End Sub



Sub pegaInfoArquivos()
    Dim fso As FileSystemObject
    Dim temp As File
    Dim nomePasta As Folder
    Dim caminho As String
    Dim comprimento As Long
    
     Set fso = New FileSystemObject
    
    comprimento = ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).Rows.Count, "A").End(xlUp).Row
     
    For i = 2 To comprimento
        caminho = ThisWorkbook.Sheets(1).Cells(i, 1)
        If InStr(caminho, ".") Then
            Set temp = fso.GetFile(caminho)
        ThisWorkbook.Sheets(1).Cells(i, 2) = (temp.Size / 1024) / 1024
        ThisWorkbook.Sheets(1).Cells(i, 3) = temp.Type
        ThisWorkbook.Sheets(1).Cells(i, 4) = temp.DateCreated
        ThisWorkbook.Sheets(1).Cells(i, 5) = temp.DateLastAccessed
        ThisWorkbook.Sheets(1).Cells(i, 6) = temp.DateLastModified
        End If
        
    Next i

    
End Sub


Sub FolderSize()
    Dim fso As Object, fsoFolder As Object
    Dim tamanho As Long
    Const strFolderName As String = "C:\Users\Gabriel\Documents"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFolder = fso.GetFolder(strFolderName)

    MsgBox fsoFolder.Size & " bytes"

    Set fsoFolder = Nothing
    Set fso = Nothing

End Sub

Sub GetFolderSizeDetails()
' Author: Dreams24
' Written for VBA Tricks and tips blog
' http://vbatricksntips.com

    Dim fso As FileSystemObject
    Dim FolderName As Folder

 'Assign path of folder for which you want it's sub folder list and size
    Rootpath = "C:\Users\Gabriel\Documents"
    ThisWorkbook.Sheets(1).Range("A2:D100").ClearContents

    'Initialize Variables And Objects
    i = 1
    Set fso = New FileSystemObject

    'Loop to get each Subfolder in the Root Path
    For Each FolderName In fso.GetFolder(Rootpath).SubFolders
        i = i + 1
        ThisWorkbook.Sheets(1).Cells(i, 1) = FolderName.Name
        On Error Resume Next
        'Folder-Name.Size returns value in Bytes. Thus divided by 1021 to convert in MB
        ThisWorkbook.Sheets(1).Cells(i, 2) = (FolderName.Size / 1024) / 1024
    Next

    MsgBox "All folders listed with their size"
End Sub


