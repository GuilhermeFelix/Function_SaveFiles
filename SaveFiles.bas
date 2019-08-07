Function SaveFiles()

    On Error GoTo ErrChk

    Application.DisplayAlerts = False
    Dim olApp As Outlook.Application
    Dim objNS As Outlook.Namespace
    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    Dim olFolder As Outlook.MAPIFolder
    strEmailPath = mdlMain.rfEmailPath
    Set olFolder = GetFolderPath(strEmailPath)
    If (olFolder Is Nothing) Then
        Set olFolder = GetFolderPath(Replace(strEmailPath, " [XIDPATHMAILFOLDERX]", ""))
    End If
    Dim Item As Object
    strPath = wsConfig.Range("prfPDFPath")
    If (Right(strPath, 1) <> "\") Then strPath = strPath & "\"
    If (Mid(strPath, 2, 2) <> ":\") Then strPath = Environ("userprofile") & "\" & strPath
    
    For Each Item In olFolder.Items 'Lista Itens
        If TypeOf Item Is Outlook.MailItem Then
            Dim oMail As Outlook.MailItem: Set oMail = Item
            If (oMail.Categories = "") Then
                cCopied = False
                strDateKey = Replace(Replace(Replace(oMail.CreationTime, ":", ""), "/", ""), " ", "") 'Troca caracteres
                For Each objAtt In oMail.Attachments
                    strFileName = ""
                    If ((LCase(Right(objAtt.DisplayName, 3)) = "pdf") Or (LCase(Right(objAtt.DisplayName, 3)) = "xls") Or (LCase(Right(objAtt.DisplayName, 4)) = "xlsx") Or (LCase(Right(objAtt.DisplayName, 3)) = "csv") Or (LCase(Right(objAtt.DisplayName, 3)) = "htm")) Then
                        
                        'Caso existam emails com arquivos não necessarios. Eliminar os arquivos desnecessários com o trecho abaixo:
                        If (InStr(1, LCase(objAtt.DisplayName), "informthedisplayname") = 0 And Not (mdlMain.rfOrigin = "informorigin" And (LCase(Right(objAtt.DisplayName, 3)) = "pdf" Or LCase(Right(objAtt.DisplayName, 3)) = "htm" Or LCase(Right(objAtt.DisplayName, 4)) = "html"))) Then
                            strFileName = strFileName & strDateKey & " - "
                            If (oMail.FlagStatus = olFlagMarked) Then strFileName = strFileName & "[HIGH] - "
                            iExtLen = InStr(1, Right(objAtt.DisplayName, 5), ".")
                            If (iExtLen = 0) Then Stop 'O Anexo não tem extension?
                            strFileName = strFileName & Left(oMail.Subject, 40) & " - " & Replace(Left(objAtt.DisplayName, 40), Right(objAtt.DisplayName, 6 - iLenExt), "") & Right(objAtt.DisplayName, 6 - iLenExt) 'Salva Anexo
                            strFileName = SymbolTxtConverter(strFileName, 1)
                            'strFileName = Replace(Replace(Replace(Replace(Replace(Replace(strFileName, "*", ""), ":", ""), "/", ""), "\", ""), ">", ""), "<", "") 'Troca caracteres
                            objAtt.SaveAsFile strPath & "\" & strFileName
                            Set objAtt = Nothing
                            cCopied = True
                        End If
                    End If
                Next

                If (cCopied) Then
                    oMail.Categories = "Processando" 'Altera categoria
                    oMail.Save
                Else
                    oMail.Categories = "Verificar Anexo" 'Altera categoria
                    oMail.Save
                End If
            End If
        End If
    Next
    
    Application.DisplayAlerts = True
    
    Exit Function
ErrChk:
    MsgBox Err.Description
    Stop
    Resume
    
End Function
