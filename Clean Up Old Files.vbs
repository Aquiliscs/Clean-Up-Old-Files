On Error Resume Next

TempoDeCriacao = 3600 'Em Segundos
TempoEntreVerifiacoes = 3600000 'Em milissegundos
ListaDiretorios = "C:\Temp\ListaDigitalizados.txt"

Do
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set listFile = fso.OpenTextFile(ListaDiretorios)
    While Not listFile.AtEndOfStream
        folderPath = listFile.ReadLine()
        Set folder = fso.GetFolder(folderPath)
        For Each subFolder in folder.SubFolders
            If DateDiff("s", subFolder.DateCreated, Now) >= TempoDeCriacao Then
                subFolder.Delete True
            End If
        Next
        For Each file in folder.Files
            If DateDiff("s", file.DateCreated, Now) >= TempoDeCriacao Then
                file.Delete True
            End If
        Next
    Wend
    listFile.Close
    Set fso = Nothing
    WScript.Sleep(TempoEntreVerifiacoes)
Loop
