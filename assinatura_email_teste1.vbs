On Error Resume Next

Dim strUser, strUsuario
Dim xmlhttp, ostream, FS
Dim url, local, Tamanho, largura, altura
Dim Mensagem, Texto
Dim blnWeOpenedWord
Dim pastaOutlook, pastaAssinaturas

' Defina a pasta de assinaturas do Outlook e a pasta de assinaturas do usuário
pastaOutlook = CreateObject("WScript.Shell").SpecialFolders("AppData") & "\Microsoft\Signatures"
pastaAssinaturas = "C:\Users\" & strUsuario & "\AppData\Roaming\Microsoft\Signatures\"

' Inicie os Objetos
Set Wshell = CreateObject("Wscript.Shell")
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Set ostream = CreateObject("ADODB.Stream")
Set FS = CreateObject("Scripting.FileSystemObject")

' Objeto utilizado para criar a assinatura de e-mail no Outlook
Set objWord = CreateObject("Word.Application")
blnWeOpenedWord = True

Texto = "Esssa mensagem e reservada e sua divulgacao, distribuicao, reproducao ou qualquer forma de uso e proibida e depende de previa autorizacao desta instituicao. O remetente utiliza o correio eletronico no exercicio do seu trabalho ou em razao dele, eximindo esta instituicao de qualquer responsabilidade por utilizacao indevida. Se voce recebeu esta mensagem por engano, favor elimina-la imediatamente."
Texto = Texto & vbCrLf & " This message is reserved and its disclosure, distribution, reproduction or any other form of use is prohibited and shall depend upon previous proper authorization. The sender uses the electronic mail in the exercise of his/her work or by virtue thereof, and the institution accepts no liability for its undue use. If you have received this e-mail by mistake, please delete it immediately."

' Função para definir o nome do arquivo da assinatura de e-mail
Function setaLocalAssinatura
    If strUsuario <> "diesgor" Then
        local = pastaAssinaturas & "sign_" & LCase(strUsuario) & ".htm" ' Alteração do formato de arquivo para .htm
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
        If Not (FS.FolderExists(pastaAssinaturas)) Then ' Alteração para verificar a pasta de assinaturas do usuário
            FS.CreateFolder (pastaAssinaturas)
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

    MsgBox "Assinatura gerada com sucesso!"
    getImagemAssinatura = True
End Function

' Função para definir a assinatura padrão no Outlook
Function definirAssinaturaPadrao()
    Set objEmailOptions = objWord.EmailOptions
    Set objSignatureObjects = objWord.EmailOptions.EmailSignature
    Set objSignatureEntries = objSignatureObjects.EmailSignatureEntries
    
    ' Define a nova assinatura como a assinatura padrão para novas mensagens e respostas/encaminhamentos
    objSignatureObjects.NewMessageSignature = "sign_" & strUsuario
    objSignatureObjects.ReplyMessageSignature = "sign_" & strUsuario
End Function

' Chamadas para funções principais
If (getUserName) Then
    setDimensoesAssinatura
    If (setaLocalAssinatura) Then
        If (getImagemAssinatura) Then
            ' Criação da assinatura no diretório de assinaturas do usuário
            Set objDoc = objWord.Documents.Add()
            Set objSelection = objWord.Selection
            objDoc.SaveAs local, 10 ' Salva o documento como formato HTML
            
            ' Define a assinatura padrão no Outlook
            definirAssinaturaPadrao
            
            ' Fechar o Word
            objWord.Quit
            
            Mensagem = "Assinatura gerada com sucesso!"
        Else
            MsgBox "Erro ao salvar imagem da assinatura de e-mail"
        End If
    Else
        MsgBox "Erro ao gerar o nome e o local onde o arquivo da assinatura de e-mail será salvo"
    End If
Else
    MsgBox "Erro ao buscar nome de usuário logado no computador"
End If
