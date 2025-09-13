Imports System.Net.Mail


Public Class Email


    ' </summary>

    ' <param name="from">Endereco do Remetente</param>

    ' <param name="recepient">Destinatario</param>

    ' <param name="bcc">recipiente Bcc</param>

    ' <param name="cc">recipiente Cc</param>

    ' <param name="subject">Assunto do email</param>

    '<param name="body">Corpo da mensagem de email</param>


    Public Shared Sub enviaMensagemEmail(ByVal from As String, ByVal recepient As String, ByVal bcc As String, ByVal cc As String, ByVal subject As String, ByVal body As String, ByVal servidorSMTP As String, ByVal nfe_chave As String)


        ' cria uma instância do objeto MailMessage

        Dim mMailMessage As New MailMessage()
        Dim retval As String
        Dim arquivo As String
        Dim arquivo99 As String
        'Dim recepient_cp As String
        'Dim aux As String
        'Dim email_to As String
        'Dim x As Integer
        'Dim y As Integer
        ' Define o endereço do remetente

        mMailMessage.From = New MailAddress(from)

        'recepient_cp = recepient
        'y = Len(LTrim(RTrim(recepient_cp)))

        '' Define o destinario da mensagem
        'x = InStr(recepient_cp, ";")

        'While x > 0
        '    email_to = recepient_cp.Substring(0, x - 1)
        '    mMailMessage.To.Add(New MailAddress(email_to))
        '    aux = recepient_cp.Substring(x)

        '    recepient_cp = aux

        '    x = InStr(recepient_cp, ";")
        '    y = Len(recepient_cp)
        'End While

        'mMailMessage.To.Add(New MailAddress(recepient_cp))


        mMailMessage.To.Add(New MailAddress(recepient))

        If Not bcc Is Nothing And bcc <> String.Empty Then
            mMailMessage.Bcc.Add(New MailAddress(bcc))
        End If

        If Not cc Is Nothing And cc <> String.Empty Then
            mMailMessage.CC.Add(New MailAddress(cc))
        End If


        mMailMessage.Subject = subject
        mMailMessage.Body = body
        mMailMessage.IsBodyHtml = True
        mMailMessage.Priority = MailPriority.Normal

        arquivo = "\\pelicano\desuninfe$\uninfe4\retorno\" & nfe_chave & "-nfe.err"
        retval = Dir(arquivo)
        arquivo99 = nfe_chave & "-nfe.err"
        If retval = arquivo99 Then
            mMailMessage.Attachments.Add(New Attachment(arquivo))
        End If

        arquivo = "\\pelicano\desuninfe$\uninfe4\SemProtocolo\" & nfe_chave & "-pro-rec.xml"
        retval = Dir(arquivo)
        arquivo99 = nfe_chave & "-pro-rec.xml"
        If retval = arquivo99 Then
            mMailMessage.Attachments.Add(New Attachment(arquivo))
        End If

        Dim mSmtpClient As New SmtpClient(servidorSMTP)

        mSmtpClient.Send(mMailMessage)


    End Sub


End Class


