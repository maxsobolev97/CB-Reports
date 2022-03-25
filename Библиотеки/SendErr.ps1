Class SendErrors{

    $serverSmtp = "example.example.ru" 
    $port = 25
    $From = "example@example.ru" 
    $To = "example@example.ru" 
    $subject = "Ошибка обработки!"
    $user = "example"
    $pass = "example"

    SendMail($fileLog, $fileOtch, $form, $subject){

        $attLog = New-object Net.Mail.Attachment($fileLog)
        $attOtch = New-object Net.Mail.Attachment($fileOtch)
        $mes = New-Object System.Net.Mail.MailMessage
        $mes.From = $this.from
        $mes.To.Add($this.to) 
        $mes.Subject = $subject 
        $mes.IsBodyHTML = $true 
        $mes.Body = "Ошибка обработки!"
        $mes.Attachments.Add($attLog)
        $mes.Attachments.Add($attOtch) 
        $smtp = New-Object Net.Mail.SmtpClient($this.serverSmtp, $this.port)
        $smtp.Credentials = New-Object System.Net.NetworkCredential($this.user, $this.pass);
        $smtp.Send($mes) 
        $attLog.Dispose()
        $attOtch.Dispose()
    }

    SendMSG($fileOtch, $subject, $body){

        $attOtch = New-object Net.Mail.Attachment($fileOtch)
        $mes = New-Object System.Net.Mail.MailMessage
        $mes.From = $this.from
        $mes.To.Add($this.to) 
        $mes.Subject = $subject 
        $mes.IsBodyHTML = $true 
        $mes.Body = "$body"
        $mes.Attachments.Add($attOtch) 
        $smtp = New-Object Net.Mail.SmtpClient($this.serverSmtp, $this.port)
        $smtp.Credentials = New-Object System.Net.NetworkCredential($this.user, $this.pass);
        $smtp.Send($mes) 
        $attOtch.Dispose()
    }

    SendMail($form, $subject, $body){

        $mes = New-Object System.Net.Mail.MailMessage
        $mes.From = $this.from
        $mes.To.Add($this.to) 
        $mes.Subject = $subject 
        $mes.IsBodyHTML = $true 
        $mes.Body = "$body"
        $smtp = New-Object Net.Mail.SmtpClient($this.serverSmtp, $this.port)
        $smtp.Credentials = New-Object System.Net.NetworkCredential($this.user, $this.pass);
        $smtp.Send($mes) 

    }

    SendConfirm($to, $subject, $body){

        $mes = New-Object System.Net.Mail.MailMessage
        $mes.From = $this.from
        $mes.To.Add($to) 
        $mes.Subject = $subject 
        $mes.IsBodyHTML = $true 
        $mes.Body = "$body"
        $smtp = New-Object Net.Mail.SmtpClient($this.serverSmtp, $this.port)
        $smtp.Credentials = New-Object System.Net.NetworkCredential($this.user, $this.pass);
        $smtp.Send($mes) 

    }

}