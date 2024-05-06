On Error Resume Next

Dim strUser, strUsuario
Dim xmlhttp, ostream, FS
Dim url, local, Tamanho, largura, altura
Dim Mensagem, Texto
Dim blnWeOpenedWord
Dim pastaOutlook

' Defina a pasta de assinaturas do Outlook
pastaOutlook = CreateObject("WScript.Shell").SpecialFolders("AppData") & "\Microsoft\Signatures"

' Inicie os Objetos
Set Wshell = CreateObject("Wscript.Shell")
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Set ostream = CreateObject("ADODB.Stream")
Set FS = CreateObject("Scripting.FileSystemObject")

' Objeto utilizado para criar a assinatura de e-mail no Outlook
Set objWord = CreateObject("Word.Application")
blnWeOpenedWord = True

Texto = "Essa mensagem e reservada e sua divulgacao, distribuicao, reproducao ou qualquer forma de uso e proibida e depende de previa autorizacao desta instituicao. O remetente utiliza o correio eletronico no exercicio do seu trabalho ou em razao dele, eximindo esta instituicao de qualquer responsabilidade por utilizacao indevida. Se voce recebeu esta mensagem por engano, favor elimina-la imediatamente."
Texto = Texto & vbCrLf & " This message is reserved and its disclosure, distribution, reproduction or any other form of use is prohibited and shall depend upon previous proper authorization. The sender uses the electronic mail in the exercise of his/her work or by virtue thereof, and the institution accepts no liability for its undue use. If you have received this e-mail by mistake, please delete it immediately."

' Função para definir o nome do arquivo da assinatura de e-mail
Function setaLocalAssinatura
    If strUsuario <> "diesgor" Then
        local = pastaOutlook & "\sign__" & LCase(strUsuario) & ".png"
        setaLocalAssinatura = True
    Else
        setaLocalAssinatura = False
    End If
End Function

' Função para obter o nome de usuário logado
Function getUserName
    strUsuario = Wshell.ExpandEnvironmentStrings("%username%")
    If Not (strUsuario) Then
        getUserName = False
    Else
        getUserName = True
    End If
End Function

' Função para definir as dimensões da Assinatura de E-mail
Function setDimensoesAssinatura
    largura = "543"
    altura = "134"
End Function

' Função para realizar uma consulta na API
Function makeHttpRequest(url)
    xmlhttp.open "GET", url, False
    xmlhttp.send
    makeHttpRequest = xmlhttp.responseText
End Function

' Função para buscar imagem via API e salvar no diretório definido
Function getImagemAssinatura
    url = "https://assinaturas.noroeste-am.com.br/api/sign/" & LCase(strUsuario)
    makeHttpRequest url

    If xmlhttp.Status = 200 Then
        If Not (FS.FolderExists(pastaOutlook)) Then
            FS.CreateFolder (pastaOutlook)
        End If

        With ostream
            .Type = 1 ' binary
            .Mode = 3 ' read-write
            .Open
            .Write xmlhttp.responseBody
            .SaveToFile local, 2 ' save-create-overwrite
            .Close
        End With

        If Err.Number <> 0 Then
            If Err.Number <> 13 Then
                Mensagem = Err.Number & " - Sem permissão para salvar imagem na pasta: " & pastaOutlook
                MsgBox Mensagem
            End If
        End If

        getImagemAssinatura = True
    End If

    MsgBox "Assinatura copiada com sucesso!"
    getImagemAssinatura = True
End Function

' Chamadas para funções principais
If (getUserName) Then
    setDimensoesAssinatura
    If (setaLocalAssinatura) Then
        If (getImagemAssinatura) Then
            Set objDoc = objWord.Documents.Add()
            Set objSelection = objWord.Selection
            Set objEmailOptions = objWord.EmailOptions
            Set objSignatureObjects = objWord.EmailOptions.EmailSignature
            Set objSignatureEntries = objSignatureObjects.EmailSignatureEntries

            Set objShape = objDoc.Shapes
            Set objRange = objDoc.Range()

            Set objEmailOptions = objWord.EmailOptions
            Set objSignatureObject = objEmailOptions.EmailSignature
            Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

            objSignatureObjects.NewMessageSignature = strUsuario & "(" & strUsuario & "@grupocopar.com.br)"
            objSignatureObjects.ReplyMessageSignature = strUsuario & "(" & strUsuario & "@grupocopar.com.br)"

            objDoc.Tables.Add objRange, 2, 1
            Set objTable = objDoc.Tables(1)

            Set assinatura = objTable.Cell(1, 1).Range.InlineShapes.AddPicture(local)
            With assinatura
                .Height = altura
                .Width = largura
            End With

            objTable.Cell(2, 1).Range.Font.Name = "Calibri"
            objTable.Cell(2, 1).Range.Font.Size = 8
            objTable.Cell(2, 1).Range.Font.Color = "green"
            objTable.Range.ParagraphFormat.SpaceAfter = 1
            objTable.Cell(2, 1).Range.Text = Texto

            objTable.Columns(1).PreferredWidth = largura
            objTable.Columns(2).PreferredWidth = largura

            Set objSelection = objDoc.Range()
            objSignatureEntries.Add strUsuario & "(" & strUsuario & "@grupocopar.com.br)", objSelection
            objSignatureObjects.NewMessageSignature = strUsuario & "(" & strUsuario & "@grupocopar.com.br)"
            objSignatureObjects.ReplyMessageSignature = strUsuario & "(" & strUsuario & "@grupocopar.com.br)"

            Set objSelection = objDoc.Range()
            objSignatureEntries.Add strUsuario & "(" & strUsuario & "@grupocopar.com.br)", objSelection
            objSignatureObjects.NewMessageSignature = strUsuario & "(" & strUsuario & "@grupocopar.com.br)"
            objSignatureObjects.ReplyMessageSignature = strUsuario & "(" & strUsuario & "@grupocopar.com.br)"

            Mensagem = "Assinatura gerada com sucesso!!"

            On Error Resume Next
            objDoc.Close 0
            If blnWeOpenedWord Then
                objWord.Quit
            End If
        Else
            MsgBox "Erro ao salvar imagem da assinatura de e-mail"
        End If
    Else
        MsgBox "Erro ao gerar o nome e o local onde o arquivo da assinatura de e-mail será salvo"
    End If
Else
    MsgBox "Erro ao buscar nome de usuário logado no computador"
End If
