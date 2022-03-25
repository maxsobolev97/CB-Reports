Class HTML_Templates{
    
    [CSS_Template]$htmlCSS = [CSS_Template]::new()
    [HTML_Funcs]$htmlFuncs = [HTML_Funcs]::new()

    $htmlHead = "<html><head><meta charset='utf-8'>" + $this.htmlCSS.cssStart + $this.htmlCSS.list + $this.htmlCSS.cssEnd + `
                $this.htmlCSS.cssStart + $this.htmlCSS.list + $this.htmlCSS.cssEnd + $this.htmlCSS.cssStart + $this.htmlCSS.table + `
                $this.htmlCSS.cssEnd + $this.htmlCSS.cssStart + $this.htmlCSS.accordion + $this.htmlCSS.cssEnd + $this.htmlScript.accardion + "</head>"

    $htmlBodyStart = "<body>"
    $htmlBodyEnd = "</body></html>"
    $htmlTableStart = "<table><tr><td>"
    $htmlTableStart319 = "<table style='width: 319px; margin-top: 2%;margin-right: 50px;'><tr><td>"
    $htmlTableEnd = "</td></tr></table>"

    [string]FirstPage([string]$htmlString){

        $logs = $this.htmlFuncs.listTrasportLog()
        $instr = $this.htmlFuncs.listInstructions()

        $html = $this.htmlHead + $this.htmlBodyStart + `
                "<table style='margin-left: 12%;'><tr><td style='display: block;'>" + `
                $this.htmlTableStart319 + "<div class='block2 style='width: 319px;'><ol class='rounded'>" + $logs + "</div>" + "<div class='block2 style='width: 319px;'><ol class='rounded'>" + $instr + $this.htmlTableEnd + `
                "</td><td>" +`
                $this.htmlTableStart +  "<h1>Доступные формы отчетности:</h1><div class='block1'><ol class='rounded'>" + `
                $htmlString + $this.htmlTableEnd + `
                "</div></td></tr></table>" + `
                $this.htmlBodyEnd

        return $html

    }

    [string]SecondPage([string]$url){
        
        $form = $url.Replace("/" , "")
        $logINPath = "W:\example\$form\IN\LOG"
        $logOUTPath = "W:\example\$form\OUT\LOG"
        $errINPath = "W:\example\$form\IN\ERR"
        $errOUTPath = "W:\example\$form\OUT\ERR"

        $logINFilesHTML = "<h2>Папки логов о приеме отчетности:</h2><ol class='rounded'>"
        $logOUTFilesHTML = "<h2>Папки логов об отправке отчетности:</h2><ol class='rounded'>"
        $errINFilesHTML = "<h2>Файлы с ошибками при приеме:</h2><ol class='roundeder'>"
        $errOUTFilesHTML = "<h2>Файлы с ошибками при отправке:</h2><ol class='roundeder'>"

        $logINString = $this.htmlFuncs.listDirs($form, $logINPath)
        $logOUTString = $this.htmlFuncs.listDirs($form, $logOUTPath)
        $errINString = $this.htmlFuncs.listDirs($form, $errINPath)
        $errOUTString = $this.htmlFuncs.listDirs($form, $errOUTPath)
        
        $html = ""

        $htmlListEnd = "</ol><br>"

        if($errINString){
        
            $html = $this.htmlHead + $this.htmlBodyStart + $this.htmlTableStart + "<h1>Доступные данные для формы $form</h1><br><div class='block1'>" + `
                        $errINFilesHTML + $errINString + $htmlListEnd + `
                        $logINFilesHTML + $logINString + $htmlListEnd + `
                        $logOUTFilesHTML + $logOUTString + $htmlListEnd + `
                        $this.htmlTableEnd + "</div>" + $this.htmlBodyEnd
        
        } elseif($errOUTString){
        
            $html = $this.htmlHead + $this.htmlBodyStart + $this.htmlTableStart + "<h1>Доступные данные для формы $form</h1><br><div class='block1'>" + `
                        $errOUTFilesHTML + $errOUTString + $htmlListEnd + `
                        $logINFilesHTML + $logINString + $htmlListEnd + `
                        $logOUTFilesHTML + $logOUTString + $htmlListEnd + `
                        $this.htmlTableEnd + "</div>" + $this.htmlBodyEnd
        
        } elseif($errINString -and $errOUTString){
        
            $html = $this.htmlHead + $this.htmlBodyStart + $this.htmlTableStart + "<h1>Доступные данные для формы $form</h1><br><div class='block1'>" + `
                    $errINFilesHTML + $errINString + $htmlListEnd + `
                    $errOUTFilesHTML + $errOUTString + $htmlListEnd + `
                    $logINFilesHTML + $logINString + $htmlListEnd + `
                    $logOUTFilesHTML + $logOUTString + $htmlListEnd + `
                    $this.htmlTableEnd + "</div>" + $this.htmlBodyEnd
        
        } else {
        
            $html = $this.htmlHead + $this.htmlBodyStart + $this.htmlTableStart + "<h1>Доступные данные для формы $form</h1><br><div class='block1'>" + `
                    $logINFilesHTML + $logINString + $htmlListEnd + `
                    $logOUTFilesHTML + $logOUTString + $htmlListEnd + `
                    $this.htmlTableEnd + "</div>" + $this.htmlBodyEnd
        
        }

        return $html

    }

    [string]ThirdPage([string]$url){
    
        $list = $this.htmlFuncs.listDirs($url)
        $form = $list[0]
        $htmlString = $list[1]
        $html = $this.htmlHead + $this.htmlBodyStart + $this.htmlTableStart + "<h1>Доступные данные для $form</h1><br><div class='block1'>" + `
                $htmlString + $this.htmlTableEnd + "</div>" + $this.htmlBodyEnd

        return $html

    }

}

Class CSS_Template{

    $cssStart = "<style type='text/css'>"
    $cssEnd = "</style>"

    $list = ".rounded {
            counter-reset: li; 
            list-style: none; 
            font: 14px 'Trebuchet MS', 'Lucida Sans';
            padding: 0;
            text-shadow: 0 1px 0 rgba(255,255,255,.5);
            }
            .rounded a {
            position: relative;
            display: block;
            padding: .4em .4em .4em 2em;
            margin: .5em 0;
            background: #bde7ab;
            color: #444;
            text-decoration: none;
            border-radius: .3em;
            transition: .3s ease-out;
            }
            .rounded a:hover {background: #abcbe7;}
            .rounded a:hover:before {transform: rotate(360deg);}
            .rounded a:before {
            content: counter(li);
            counter-increment: li;
            position: absolute;
            left: -1.3em;
            top: 50%;
            margin-top: -1.3em;
            background: #8FD4C1;
            height: 2em;
            width: 2em;
            line-height: 2em;
            border: .3em solid white;
            text-align: center;
            font-weight: bold;
            border-radius: 2em;
            transition: all .3s ease-out;
            }
            
            .roundeder {
            counter-reset: li; 
            list-style: none; 
            font: 14px 'Trebuchet MS', 'Lucida Sans';
            padding: 0;
            text-shadow: 0 1px 0 rgba(255,255,255,.5);
            }
            .roundeder a {
            position: relative;
            display: block;
            padding: .4em .4em .4em 2em;
            margin: .5em 0;
            background: #e7abab;
            color: #444;
            text-decoration: none;
            border-radius: .3em;
            transition: .3s ease-out;
            }
            .roundeder a:hover {background: #fbb54d;}
            .roundeder a:hover:before {transform: rotate(360deg);}
            .roundeder a:before {
            content: counter(li);
            counter-increment: li;
            position: absolute;
            left: -1.3em;
            top: 50%;
            margin-top: -1.3em;
            background: #8FD4C1;
            height: 2em;
            width: 2em;
            line-height: 2em;
            border: .3em solid white;
            text-align: center;
            font-weight: bold;
            border-radius: 2em;
            transition: all .3s ease-out;
            }"

    $div = ".block1 {
            width: 50%;
            padding: 2%;
            }"

    $table = "table {
              width: 700px; /* Ширина таблицы */
              margin: auto; /* Выравниваем таблицу по центру окна  */
              font-size: 14px;
             }"

    $accordion = ".spoiler input, .spoiler div  { 
                    display: none; /* Скрываем содержимое */
                   }
                   .spoiler label::before {
                    content: '+';
                    margin-right: 5px; 
                    color: black;
                   }
                   /* Открытый спойлер */
                   .spoiler :checked + label::before { content: '-'; }
                   .spoiler :checked ~ div {
                    display: block;
                   }"

}

Class HTML_Funcs{

    $formsDir = "W:\example\example"
    $transportDir = "W:\example\LOGS"
    $instructionDir = "W:\example\Instructions"

    [string]listDirs(){
    
        $htmlToDayRep = "<h2>Формы обработанные сегодня $(Get-Date -Format "dd MMMM yyyy")</h2><ol class='rounded'>"
        $htmlArcRep = "<h2>Ранее обработанные формы</h2><ol class='rounded'>"
        $htmlString = "<ol class='rounded'>"
        $forms = Get-ChildItem $this.formsDir | sort LastWriteTime -Descending

            foreach($form in $forms){

                $pathInErr = $form.FullName + "\IN\ERR"
                $pathOutErr = $form.FullName + "\OUT\ERR"
                $inErr = Get-ChildItem $pathInErr
                $outErr = Get-ChildItem $pathOutErr
                $time = Get-Date $form.LastWriteTime -Format "HH:mm"
                $date = $form.LastWriteTime.ToShortDateString().ToString()
                $tableStart = "<table style='width: 100%; font: initial;'>"
                $tdLeftStart = "<td style='width: 50%; text-align: left;'>"
                $tdRigntStart = "<td style='width: 50%; text-align: right;'>"
                $tdEnd = "</td>"
                $tableEnd = "</table>"

                if($form.LastWriteTime.Date -eq $(Get-Date).Date){
                    
                    if($inErr -or $outErr){
                    
                        $htmlToDay = "<li><a style='background: #e7abab;' href='/$form'>$tableStart $tdLeftStart $form - БЫЛИ ОШИБКИ ОБРАБОТКИ! $tdEnd $tdRigntStart $time $tdEnd $tableEnd</a></li>"
                        $htmlToDayRep = $htmlToDayRep + $htmlToDay
                    
                    } else {
                    
                        $htmlToDay = "<li><a href='/$form'>$tableStart $tdLeftStart $form $tdEnd $tdRigntStart $time $tdEnd $tableEnd</a></li>"
                        $htmlToDayRep = $htmlToDayRep + $htmlToDay
                    
                    }
                
                } else {
                    
                    if($inErr -or $outErr){
                        
                        $html = "<li><a style='background: #e7abab;' href='/$form'>$tableStart $tdLeftStart $form - БЫЛИ ОШИБКИ ОБРАБОТКИ! $tdEnd $tdRigntStart $date $tdEnd $tableEnd</a></li>"
                        $htmlArcRep = $htmlArcRep + $html

                    } else {
                    
                        $html = "<li><a href='/$form'>$tableStart $tdLeftStart $form $tdEnd $tdRigntStart $date $tdEnd $tableEnd</li>"
                        $htmlArcRep = $htmlArcRep + $html
                    
                    }

                }
    
        }

        if($htmlToDayRep -ne "<h2>Формы обработанные сегодня $(Get-Date -Format "dd MMMM yyyy")</h2><ol class='rounded'>"){
        
            $htmlString = $htmlToDayRep + "</ol>" + $htmlArcRep + "</ol>"

            return $htmlString

        } else {
        
            $htmlString = "<h2>Cегодня $(Get-Date -Format "dd MMMM yyyy") обработка форм не производилась.</h2>" + $htmlArcRep + "</ol>"

            return $htmlString
        
        }
        

    }

    [bool]formURL([string]$url){
    
        $reqForm = $url.Replace("/" , "")
        $existForm = Get-ChildItem $this.formsDir | Where-Object {$_.Name -eq $reqForm}

        if($existForm){
        
            return $true

        } else {

            return $false

        }
    
    }

    [string]listDirs([string]$form, [string]$path){
    
        $htmlString = ""
        $dirObj = Get-Item $path
        $midUrl = $dirObj.Name + "/" + $dirObj.Parent.Name

        $tableStart = "<table style='width: 100%; font: initial;'>"
        $tdLeftStart = "<td style='width: 50%; text-align: left;'>"
        $tdRigntStart = "<td style='width: 50%; text-align: right;'>"
        $tdEnd = "</td>"
        $tableEnd = "</table>"

        $startAccardion = "<div class='spoiler'><input type='checkbox' id='$($dirObj.Parent.Name)'><label for='$($dirObj.Parent.Name)' style='background-color: bisque; padding: 3px; width: 688px; display: block; text-align: center; border-radius: 3px;'>Архив логов за прошлые отправки</label><div>"
        $endAccardion = "</div>"

        $itsFirst = $true

        $files = Get-ChildItem $path | sort -Descending
        
        if($path -match "ERR"){

            foreach($file in $files){

                $time = Get-Date $file.LastWriteTime -Format "HH:mm"
                $date = $file.LastWriteTime.ToShortDateString().ToString()

                $fileName = $file.Name
                $htmlString = $htmlString + "<li><a href='/$form/$midUrl/$fileName'>$tableStart $tdLeftStart $fileName $tdEnd $tdRigntStart $time $date $tdEnd $tableEnd</a></li>"
    
            }

        } else {
            

            foreach($file in $files){

                if($itsFirst){

                    $fileName = $file.Name
                    $htmlString = $htmlString + "<li><a href='/$form/$midUrl/$fileName'>$fileName</a></li>" + $startAccardion
                    $itsFirst = $false

                } else {
                
                    $fileName = $file.Name
                    $htmlString = $htmlString + "<li><a href='/$form/$midUrl/$fileName'>$fileName</a></li>"
                
                }
    
            }

            $htmlString = $htmlString + $endAccardion
        
        }

        return $htmlString

    
    }

    [array]listDirs([string]$url){
    
        $htmlString = ""
        $htmlMainLog = ""
        $hmtlErrLog = ""
        $mUrl = $url.Split("/")
        $midAccardion = "</label><div>"
        $endAccardion = "</div>"

        $tableStart = "<table style='width: 100%; font: initial;'>"
        $tdLeftStart = "<td style='width: 50%; text-align: left;'>"
        $tdRigntStart = "<td style='width: 50%; text-align: right;'>"
        $tdEnd = "</td>"
        $tableEnd = "</table>"


        $path = $this.formsDir + "\" + $mUrl[1] + "\" + $mUrl[3] + "\" + $mUrl[2] + "\" + $mUrl[4]

        if($mUrl[3] -eq "IN"){
        
            $htmlString = "<h2>Логи приема отчетности за $($mUrl[4])</h2>"
        
        } elseif($mUrl[3] -eq "OUT"){
        
            $htmlString = "<h2>Логи отправки отчетности за $($mUrl[4])</h2><ol class='rounded'>"
        
        }

        $objMainLog = Get-Item $($path + "\" + "!MainLog.txt")

        if($objMainLog){
            $time = Get-Date $objMainLog.LastWriteTime -Format "HH:mm"
            $htmlString = $htmlString + "<br><h3>Основной лог-файл обработки</h3><ol class='rounded'>" + "<li><a href='$url/$($objMainLog.Name)'>$tableStart $tdLeftStart $($objMainLog.Name) $tdEnd $tdRigntStart $time $tdEnd $tableEnd</a></li>" + "</ol>"
        
        }

        $errFiles = Get-ChildItem $($path + "\" + "ErrLog")

        if($errFiles){
            
            $startAccardion = "<div class='spoiler'><input type='checkbox' id='ErrLog'><label for='ErrLog' style='background-color: bisque; padding: 3px; width: 688px; display: block; text-align: center; border-radius: 3px;'>"
            $htmlString = $htmlString + $startAccardion + "Лог-файлы ошибок КА/ЗК" + $midAccardion + "<ol class='roundeder'>"

            foreach($errFile in $errFiles){

                $time = Get-Date $errFile.LastWriteTime -Format "HH:mm"
                $errFileName = $errFile.Name
                $htmlString = $htmlString + "<li><a href='$url/$errFileName'>$tableStart $tdLeftStart $errFileName $tdEnd $tdRigntStart $time $tdEnd $tableEnd</a></li>"
    
            }

            $htmlString = $htmlString + "</ol>" + $endAccardion + $endAccardion + "<br>"
        
        }

        $files = Get-ChildItem $path -File

        if($files){
            
            $startAccardion = "<div class='spoiler'><input type='checkbox' id='Log'><label for='Log' style='background-color: bisque; padding: 3px; width: 688px; display: block; text-align: center; border-radius: 3px;'>"
            $htmlString = $htmlString + $startAccardion + "Лог-файлы успеха КА/ЗК" + $midAccardion +"<ol class='rounded'>"

            foreach($file in $files){

                $time = Get-Date $file.LastWriteTime -Format "HH:mm"
                $fileName = $file.Name
                if($fileName -ne "!MainLog.txt"){
                
                    $htmlString = $htmlString + "<li><a href='$url/$fileName'>$tableStart $tdLeftStart $fileName $tdEnd $tdRigntStart $time $tdEnd $tableEnd</a></li>"
                
                }
            }

            $htmlString = $htmlString + "</ol>" + $endAccardion + $endAccardion + "<br>"

        }

        return $mUrl[1], $htmlString

    
    }

    [array]downloadFile([string]$url){
    
        $filePath = ""
        $mUrl = $url.Split("/")

        if($mUrl[2] -eq "LOG"){
        
            if($mUrl[-1] -match ".txt"){

                $filePath = $this.formsDir + "\" + $mUrl[1] + "\" + $mUrl[3] + "\" + $mUrl[2] + "\" + $mUrl[4] + "\" + $mUrl[5]

            } elseif($mUrl[-1] -match ".err"){
        
                $filePath = $this.formsDir + "\" + $mUrl[1] + "\" + $mUrl[3] + "\" + $mUrl[2] + "\" + $mUrl[4] + "\ErrLog\" + $mUrl[5]
        
            }

        } elseif($mUrl[2] -eq "ERR"){
        
            $filePath = $this.formsDir + "\" + $mUrl[1] + "\" + $mUrl[3] + "\" + $mUrl[2] + "\" + $mUrl[4]

        } elseif($mUrl[1] -match "Mail_"){
        
            $filePath = $this.transportDir + "\" + $mUrl[1] + "\" + $mUrl[2]

        } elseif($mUrl[1] -match "Instructions"){
        
            $filePath = $this.instructionDir + "\" + $mUrl[2]

        }
                
        $content = Get-Content -Encoding Byte -Path ($filePath)

        return $content

    }

    [string]listTrasportLog(){
    
        $tableStart = "<table style='width: 100%; font: initial;'>"
        $tdLeftStart = "<td style='width: 50%; text-align: left;'>"
        $tdRigntStart = "<td style='width: 50%; text-align: right;'>"
        $tdEnd = "</td>"
        $tableEnd = "</table>"

        $endAccardion = "</div>"

        $itsFirstSVK = $true
        $itsFirstMail = $true

        $svkString = "<h2>Логи работы транспорта SVK:</h2>"
        $mailString = "<h2>Логи работы транспорта MAIL:</h2>"

        $htmlString = ""

        $dirs = Get-ChildItem $this.transportDir

        foreach($dir in $dirs){
        
            $startAccardion = "<div class='spoiler'><input type='checkbox' id='$($dir.Name)'><label for='$($dir.Name)' style='background-color: bisque; padding: 3px; display: block; text-align: center; border-radius: 3px;'>Архив транспортных логов</label><div>"
        
            if($dir.Name -eq "Mail_SVK"){

                $htmlString = $htmlString + $svkString + "<ol class='rounded'>"

            } elseif($dir.Name -eq "Mail_Robot1"){
            
                $htmlString = $htmlString + $mailString + "<ol class='rounded'>"
            
            }

            $logs = Get-ChildItem $dir.FullName | sort -Descending

            foreach($log in $logs){

                $time = Get-Date $log.LastWriteTime -Format "HH:mm"

                if($itsFirstSVK -and $($dir.Name ) -eq "Mail_SVK"){

                    $htmlString = $htmlString + "<li><a href='/$($log.Directory.Name)/$($log.Name)'>$tableStart $tdLeftStart $($log.Name)  $tdEnd $tdRigntStart $time $tdEnd $tableEnd</a></li>" + $startAccardion
                    $itsFirstSVK = $false
            
                } elseif($itsFirstMail -and $($dir.Name ) -eq "Mail_Robot1") {
                
                    $htmlString = $htmlString + "<li><a href='/$($log.Directory.Name)/$($log.Name)'>$tableStart $tdLeftStart $($log.Name)  $tdEnd $tdRigntStart $time $tdEnd $tableEnd</a></li>" + $startAccardion
                    $itsFirstMail = $false
                
                } else{
                
                    $htmlString = $htmlString + "<li><a href='/$($log.Directory.Name)/$($log.Name)'>$tableStart $tdLeftStart $($log.Name)  $tdEnd $tdRigntStart $time $tdEnd $tableEnd</a></li>"
                
                }

            }

            $htmlString = $htmlString + "</ol>" + "<br>"

        }

        return $htmlString
    
    }

    [string]listInstructions(){
    
        $instrString = "<h2>Инструкции:</h2>"

        $htmlString = $instrString + "<ol class='rounded'>"

        $instrs = Get-ChildItem $this.instructionDir

        foreach($instr in $instrs){

            $htmlString = $htmlString + "<li><a href='/$($instr.Directory.Name)/$($instr.Name)'>$($instr.Name)</a></li>"

        }

        $htmlString = $htmlString + "</ol>" + "<br>"

        return $htmlString
    
    }
}

[HTML_Templates]$HTML_Templates = [HTML_Templates]::new()
[CSS_Template]$CSS_Template = [CSS_Template]::new()
[HTML_Funcs]$HTML_Funcs = [HTML_Funcs]::new()

$http = [System.Net.HttpListener]::new()
$http.Prefixes.Add("http://+:8080/")
$http.Start()


while($http.IsListening){

    $context = $http.GetContext()
    $url = $context.Request.RawUrl
    $method = $context.Request.HttpMethod
    
    if($method -eq "GET" -and $url -eq "/"){
    
        [string]$html = $HTML_Templates.FirstPage($HTML_Funcs.listDirs())
        $context.Response.StatusCode = 200
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.OutputStream.Close()
    
    } elseif($method -eq "GET" -and $HTML_Funcs.formURL($url)){
        
        [string]$html = $HTML_Templates.SecondPage($url)
        $context.Response.StatusCode = 200
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.OutputStream.Close()
    
    } elseif($method -eq "GET" -and $url -match "/LOG/" -and $url.Split("/").Count -eq 5){
        
        [string]$html = $HTML_Templates.ThirdPage($url)
        $context.Response.StatusCode = 200
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.OutputStream.Close()
    
    } elseif($method -eq "GET" -and (($url -match "/LOG/" -and $url.Split("/").Count -eq 6) -or ($url -match "/ERR/") -or ($url -match "/Mail_") -or ($url -match "/Instructions"))){
        
        $Content = $HTML_Funcs.downloadFile($url)
        $Context.Response.ContentType="application/octet-stream"
        $Context.Response.ContentEncoding=[System.Text.Encoding]::Default
        $Context.Response.ContentLength64=$Content.Length
        $Context.Response.KeepAlive=$false
        $Context.Response.StatusCode=200
        $Context.Response.StatusDescription="OK"
        $Context.Response.OutputStream.Write($Content, 0, $Content.Length)
        $Context.Response.OutputStream.Close()
        $Context.Response.Close()
    
    } elseif($method -eq "GET" -and $url -eq "/stop"){
    
        [string]$html = "Stopped!"
        $context.Response.StatusCode = 200
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.OutputStream.Close()
        $http.Stop()
    
    }
    
    [System.GC]::Collect()
}