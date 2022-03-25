Import-Module "W:\example\Library_Files.ps1"
Import-Module "W:\example\SendErr.ps1"

function isFilesInDirs{

    Param(

        $sReportsPath,
        $sPS

    )

    $sReportsDirsSVK = Get-ChildItem $sReportsPath -Filter "FNS"
    RunScripts $sReportsDirsSVK

} # Опрос папок и вызов обрабатывающих скриптов

function RunScripts{

    Param($sReportsDirs)

    foreach($sReportDir in $sReportsDirs){

        [string]$FolderName = $sReportDir.Name
        Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Просматриваю папку OUT: $FolderName"
        $sReportDirOut = $sReportDir.FullName + "\OUT"
        $sFilesInDir = Get-ChildItem $sReportDirOut -File

        if($sFilesInDir -ne $null){
            
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Обнаружен файл в папке OUT $FolderName"
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Начинаю обработку поступившего файла"
            Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Запуск скрипта обработки $FolderName"
            $sScriptName = $sReportDir.FullName + "\Scripts\" + $sReportDir.Name + "_OUT.ps1"
           
            Start-Process -FilePath $sPS $sScriptName -Wait -NoNewWindow

        }

    }

} # Запуск обрабатывающих скриптов

[Run_exe]$Run_exe = [Run_exe]::New()
$sPS = $Run_exe.sPS
[string[]]$aVFDImages = $Run_exe.aVFDImages
[Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::New()
$sReportsPath = $Path_to_Folders.sReportsPath
[SendErrors]$SendErr = [SendErrors]::New()

Start-Job -FilePath "W:\example\ErrorsInCrypTool.ps1" -Name "CryptErrors"

$Run_exe.fMountVFDImage($aVFDImages[0])

while($true){

    isFilesInDirs $sReportsPath $sPS
    [System.GC]::Collect()
    if($Error){
        $err = $Error.Item(0).ToString()
        $SendErr.SendMail("FNS_OUT", "Ошибка опроса папок FNS_OUT!", "$err")
        $Error.Clear()
    }
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")    Пауза 10 секунд"
    Start-Sleep -Seconds 10

}