Sub CorrigirOrtografiaEmTodosDocumentos()
    Dim doc As Document
    Dim folderPath As String
    Dim fileName As String
    
    ' Defina o caminho da pasta onde estão os documentos a serem corrigidos
    folderPath = "C:\Caminho\Para\Os\Documentos\"
    
    ' Verifique se a pasta existe
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "A pasta especificada não foi encontrada.", vbExclamation
        Exit Sub
    End If
    
    ' Desabilita a exibição de alertas durante a verificação ortográfica
    Application.DisplayAlerts = wdAlertsNone
    
    ' Loop através de todos os arquivos na pasta
    fileName = Dir(folderPath & "*.doc*")
    Do While fileName <> ""
        ' Abre cada documento
        Set doc = Documents.Open(folderPath & fileName)
        
        ' Realiza a verificação ortográfica
        doc.CheckSpelling
        
        ' Salva e fecha o documento
        doc.Save
        doc.Close
        
        ' Vai para o próximo arquivo na pasta
        fileName = Dir
    Loop
    
    ' Habilita a exibição de alertas novamente
    Application.DisplayAlerts = wdAlertsAll
    
    MsgBox "Correção ortográfica concluída em todos os documentos.", vbInformation
End Sub
