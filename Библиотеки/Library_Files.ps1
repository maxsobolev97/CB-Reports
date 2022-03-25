<#
    Библиотека для работы с файлами отчетности (зашифровать / расшифровать,
    поставить / удалить подпись, набор файловых констант, работа с файлами, библиотека для запуска
    exe файлов.
#>

Class Crypt_funcs{

    [string]$sCrypTool = "C:\example\CrypTool.exe"
    [string]$sNBKISkript = "C:\example\Send_NBKI.cmd"
    [string]$sCryptTCP = "C:\example\cryptcp.x64.exe"

    [string]$sBRlocalPSE = "pse://signed/C:\example\MDPREI\scs\example\local.pse"
    [string]$sBRlocalGDBM = "file://C:\example\MDPREI\scs\example\local.gdbm"
    [string]$sFOIVlocalPSE = "pse://signed/C:\example\scs\example\local.pse"
    [string]$sFOIVlocalGDBM = "file://C:\Uexample\scs\example\local.gdbm"
    [string]$sDriveName = "b:\"
    
    [string]$sCERTfns311 = "CN=example,ST=example"
    [string]$sCERTfss = "CN=example,ST=example"
    [string]$sCERTdit = "CN=example,ST=example"
    [string]$sCERTdfmvk = "CN=example,ST=example"
    [string]$sCERTfns = "CN=example,O=example,ST=example"
    [string]$sCERTfsfm = "CN=example,ST=example"
    [string]$sCERTvk = "CN=example example,ST=example"
    [string]$sCERTbr = "CN=example,O=example.example"
    [Run_exe]$fRunExe = [Run_exe]::new()
    [File_operations]$fDelFile = [File_operations]::new()
    
    

    [string]fDecryptCIK($sFile) {

	    [int]$iSilentLevel = 1
	    [int]$iBase64Type = 0
	    [int]$iDERType = 1
	    [int]$iRegistryStore = 0
	    [int]$iDTPlanData = 0
	    [int]$iDTEnvelopedData = 3
	    [int]$iDTSignedData = 2
	    [int]$iDecryptWizardType = 1024
	    [int]$iAllOK = 0
        [string]$sEncrytedDataFile = $sFile
        [string]$sFilePath = Split-Path $sFile -Parent
        [string]$sFileNameSource = Split-Path $sFile -Leaf
        [string]$sFileName, [string]$sFileExt1, [string]$sFileExt2, [string]$sFileExt3 = $sFileNameSource -split "\."
        [string]$sSignedDataFile = $sFilePath + "\" + $sFileName + "." + $sFileExt1 + "." + $sFileExt2
        [string]$sPlainDataFile = $sFilePath + "\" + $sFileName + "." + $sFileExt1
        [string]$sCheckResult = ""

        $oProfileStore = New-Object -com DigtCrypto.ProfileStore
        $oProfileStore.Open($iRegistryStore)
        $oProfiles = $oProfileStore.Store
        $oProfile = $oProfiles.DefaultProfile
        $oProfile.SilentLevel = $iSilentLevel
        $oProfile.CollectData($iDecryptWizardType) | Out-Null
        $iCheckResult = $oProfile.CheckData($iDecryptWizardType)
        if($iCheckResult -eq $iAllOK) {
            $oPKCS7Message = New-Object -com DigtCrypto.PKCS7Message
            $oPKCS7Message.Profile = $oProfile
            #$oPKCS7Message.Load($iDTEnvelopedData, $sEncrytedDataFile, "") | Out-Null
            #$oPKCS7Message.Decrypt() | Out-Null
            $oPKCS7Message.Save($iDTPlanData, $iBase64Type, $sSignedDataFile) | Out-Null
            $oPKCS7Message.Load($iDTSignedData, $sSignedDataFile, "") | Out-Null
            $oPKCS7Message.Save($iDTPlanData, $iBase64Type, $sPlainDataFile) | Out-Null
            $oPKCS7Message.Load($iDTPlanData, $sPlainDataFile, "") | Out-Null
            $oPKCS7Message.Save($iDTPlanData, $iBase64Type, $sPlainDataFile) | Out-Null
            $FileOper = $this.fDelFile
            $FileOper.fDelFile($sEncrytedDataFile)
            $FileOper.fDelFile($sSignedDataFile)
        }

        return $sPlainDataFile
    } #Расшифровка входящего сообщения ЦИК

    fEncryptReportGTUsig($sReportPath) {
        
        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sBRlocalPSE + " -w" + $this.sBRlocalGDBM + " -g" + $this.sDriveName + " -s -f" + $sReportPath) | Out-Null
        $run.fRunExe($this.sCrypTool," -v" + $this.sBRlocalPSE + " -w" + $this.sBRlocalGDBM + " -g" + $this.sDriveName + " -e -a" + $this.sCERTbr + " -f" + $sReportPath) | Out-Null

    } #Зашифровать отчетность на ГУ ЦБ

    fEncryptReportGTUsig($sReportPath, $sLog) {
        
        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sBRlocalPSE) -w$($this.sBRlocalGDBM) -g$($this.sDriveName) -s -f$sReportPath" -Wait -NoNewWindow
        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sBRlocalPSE) -w$($this.sBRlocalGDBM) -g$($this.sDriveName) -e -a$($this.sCERTbr) -f$sReportPath" -Wait -NoNewWindow

    } #Зашифровать отчетность на ГУ ЦБ

    fDecryptReplyGTUsig($sReportPath) {
        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sBRlocalPSE + " -w" + $this.sBRlocalGDBM + " -g" + $this.sDriveName + " -r -f" + $sReportPath) | Out-Null

    } #Расшифровать и снять КА отчетность ГУ ЦБ

    fDecryptReplyGTUsig($sReportPath, $sLog) {
       
        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sBRlocalPSE) -w$($this.sBRlocalGDBM) -g$($this.sDriveName) -r -f$sReportPath" -Wait -NoNewWindow

    } #Расшифровать и снять КА отчетность ГУ ЦБ

    fDecryptReportFSsig($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sFOIVlocalPSE + " -w" + $this.sFOIVlocalGDBM + " -g" + $this.sDriveName + " -d -f" + $sReportPath) | Out-Null

    } #Расшифровать отчетность Федеральных Служб 

    fDecryptReportFSsig($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -d -f$sReportPath" -Wait -NoNewWindow

    } #Расшифровать 440

    fRemoveKAReportFSsig($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sFOIVlocalPSE + " -w" + $this.sFOIVlocalGDBM + " -g" + $this.sDriveName + " -r -f" + $sReportPath) | Out-Null

    } #Снять КА с отчетности Федеральных Служб

    fRemoveKAReportFSsig($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -r -f$sReportPath" -Wait -NoNewWindow

    } #Снять КА с 440

    fSetKAReportFSsig($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sFOIVlocalPSE + " -w" + $this.sFOIVlocalGDBM + " -g" + $this.sDriveName + " -s -f" + $sReportPath) | Out-Null

    } #Установить КА на отчетность Федеральных Служб

    fSetKAReportFSsig($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -s -f$sReportPath" -Wait -NoNewWindow

    } #Установить КА на 440

    fCompressReportFSGzip($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sFOIVlocalPSE + " -w" + $this.sFOIVlocalGDBM + " -g" + $this.sDriveName + " -y -f" + $sReportPath) | Out-Null

    } #Сжать отчетность с помощью CrypTool методом GZip на Федеральные службы

    fCompressReportFSGzip($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -y -f$sReportPath" -Wait -NoNewWindow

    } #Сжать отчетность с помощью CrypTool методом GZip 440

    fDecompressReportFSGzip($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -x -f$sReportPath" -Wait -NoNewWindow

    } #Распаковать отчетность с помощью CrypTool методом GZip 440

    fEncryptFNSsig($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sFOIVlocalPSE + " -w" + $this.sFOIVlocalGDBM + " -g" + $this.sDriveName + " -e -a" + $this.sCERTfns + " -f" + $sReportPath) | Out-Null

    } #Зашифровать на Федеральные службы

    fEncryptFNSsig($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -e -a$($this.sCERTfns) -f$sReportPath" -Wait -NoNewWindow

    } #Зашифровать на Федеральные службы - ФНС 440

    fEncryptKFMsig($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sFOIVlocalPSE + " -w" + $this.sFOIVlocalGDBM + " -g" + $this.sDriveName + " -s -f" + $sReportPath) | Out-Null

    } #Зашифровать на Федеральные Службы - отчетность КФМ (финмон)

    fEncryptKFMsig($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -s -f$sReportPath" -Wait -NoNewWindow

    } #Зашифровать на Федеральные Службы - отчетность КФМ (финмон)

    fEncrypt550sig($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sFOIVlocalPSE + " -w" + $this.sFOIVlocalGDBM + " -g" + $this.sDriveName + " -e -a" + $this.sCERTdfmvk + " -f" + $sReportPath) | Out-Null

    } #Зашифровать 550-П

    fEncrypt550sig($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -e -a$($this.sCERTdfmvk) -f$sReportPath" -Wait -NoNewWindow

    } #Зашифровать 550-П

    fEncryptMIFNSsig($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sFOIVlocalPSE + " -w" + $this.sFOIVlocalGDBM + " -g" + $this.sDriveName + " -e -a" + $this.sCERTfns + ";" + $this.sCERTfss + " -f" + $sReportPath) | Out-Null
    
    } #Зашифровать MIFNS

    fEncryptMIFNSsig($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -e -a$($this.sCERTfns + ";" + $this.sCERTfss) -f$sReportPath" -Wait -NoNewWindow
    
    } #Зашифровать MIFNS

    fEncryptFSFMsig($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sFOIVlocalPSE + " -w" + $this.sFOIVlocalGDBM + " -g" + $this.sDriveName + " -e -a" + $this.sCERTfsfm + " -f" + $sReportPath) | Out-Null
    
    } #Зашифровать на Федеральные службы - ФинМон

    fEncryptFSFMsig($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -e -a$($this.sCERTfsfm) -f$sReportPath" -Wait -NoNewWindow
    
    } #Зашифровать на Федеральные службы - ФинМон

    fEncryptNBKI() {

       Start-Process $this.sNBKISkript -ArgumentList "1" -Wait

    } #Зашифровать НБКИ

    fDEncryptNBKI() {

       Start-Process $this.sNBKISkript -ArgumentList "2" -Wait

    } #Расшифровать ответ НБКИ

    fRemoveKAReportBRsig($sReportPath) {

        $run = $this.fRunExe
        $run.fRunExe($this.sCrypTool," -v" + $this.sBRlocalPSE + " -w" + $this.sBRlocalGDBM + " -g" + $this.sDriveName + " -r -f" + $sReportPath) | Out-Null
    
    } #Убрать КА с отчетности ЦБ

    fRemoveKAReportBRsig($sReportPath, $sLog) {

        $run = $this.fRunExe
        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sBRlocalPSE) -w$($this.sBRlocalGDBM) -g$($this.sDriveName) -r -f$sReportPath" -Wait -NoNewWindow
    
    } #Убрать КА с отчетности ЦБ

    fCrypt4512u($sReportPath, $sDir){
    
        $run = $this.fRunExe
        $run.fRunExe($this.sCryptTCP," -dir $sDir -signf -thumbprint 0baead054976246c6a289319c99d85615d6442ad -nochain -norev -cert -addchain -fext .sig " + $sReportPath) | Out-Null
    
    }

    fEncryptVKsig($sReportPath, $sLog) {

        Start-Process $this.sCrypTool -ArgumentList "-b0 -o$sLog -v$($this.sFOIVlocalPSE) -w$($this.sFOIVlocalGDBM) -g$($this.sDriveName) -e -a$($this.sCERTvk) -f$sReportPath" -Wait -NoNewWindow

    }

} #Общие методы и свойства для шифрования/расшифровки отчетности ЦБ и Федеральных Служб

Class Path_to_Folders{

    [string]$sReportsPath = "W:\example"
    [string]$sPath_to_local_process = "C:\example"
    [string]$sReportsPathShare = "Y:"
    [string]$sNBKIfolderIN = "\\example\IN"
    [string]$sNBKIfolderOUT = "\\example\OUT"
    [string]$sFNSFolderIn = "\\example\in"
    [string]$sFNSFolderOut = "\\example\out"
    [string]$sFSFM550FolderIn = "\\example\in"
    [string]$sFSFM550FolderOut = "\\example\out"
    [string]$sCIKFolderIn = "\\example\in"
    [string]$sCIKFolderOut = "\\example\out"
    [string]$sNBKIinWorkFolder = "example\_in"
    [string]$sNBKIoutWorkFolder = "example\_out"
    [string]$sNBKIinFolder = "C:\example\IN"
    [string]$sNBKIoutFolder = "C:\example\OUT"
    [string]$sPath_mifns_serv = "\\example"
    [string]$sPath_mifns_work = "W:\example"
    [string]$sMailLogs = "W:\example\LOGS"

} #Пути до папок с отчетностью

Class File_operations{
    
    [string]$sARJ = "C:\example\arj32.exe"
    [string]$sRAR = "C:\example\WinRAR.exe"
    [string]$s7zip = "C:\example\7z.exe"
    [Run_exe]$fRunExe = [Run_exe]::new()


    [string]fArchReport($sReportPath) {

        [string]$sFileName = Split-Path $sReportPath -Leaf
        [string]$sPath = Split-Path $sReportPath -Parent
        [string[]]$aFileNameParts = $sFileName.Split(".")
        [string]$sArchReportPath = $sPath+ "\" + $($aFileNameParts[0]) + ".arj"

        $run = $this.fRunExe
        $run.fRunExe($this.sARJ," a -e " + $sArchReportPath + " " + $sReportPath) | Out-Null

        Return $sArchReportPath

    } #Добавление файлов в архив .arj

    fDelFile($sFilePath) {
    
        Remove-Item $sFilePath

    } #Удалить файл

    fDelFiles($sFilesPath) {
        
        Remove-Item ($sFilesPath + "\*.*")

    }  #Удалить файлы

    fMoveFileToARC($sFilePath) {  

        $oFile = Get-ChildItem $sFilePath
        
        [Path_to_Folders]$Path_to_Folders = [Path_to_Folders]::new()
        $FormName = Split-Path $oFile.Directory -Parent
        if($FormName -match "OUT"){
        
            $FormName = Split-Path $FormName -Parent

        }
        $FormName = Split-Path $FormName -Leaf
        if($oFile.Directory.Name -notmatch "PROCESS"){
            $sRootDir = $Path_to_Folders.sReportsPath + "\" + $FormName + "\" + $oFile.Directory.Name
        } else{
        
            $sRootDir = $Path_to_Folders.sReportsPath + "\" + $FormName + "\OUT"
        
        }
        [string]$sDate = Get-Date -format yyyyMMddHHmm
        [string]$sARCFilePath = $sRootDir + "\ARC\" + $oFile.BaseName + "_" + $sDate + $oFile.Extension

        Move-Item $sFilePath $sARCFilePath
    
    }  #Переместить файл в архивную папку ARC

    fMoveNBKIFileToARC($sFilePath) {
        
        $oFiles = Get-ChildItem $sFilePath -File

        foreach($oFile in $oFiles) {

            $sFilePath = $oFile.FullName

            [string]$sDate = Get-Date -format yyyyMMddHHmm
            [string]$sARCFilePath = $oFile.DirectoryName + "\ARC\" + $oFile.BaseName + "_" + $sDate + $oFile.Extension

            move $sFilePath $sARCFilePath
    
        }
    
    } #Переместить файл НБКИ в архивную папку ARC

    fCopyFileToFolder($sFilePath, $sFolderPath){
        
        Copy-Item $sFilePath -Destination $sFolderPath
    
    }  #Скопировать файл в директорию

    fRenameFolder($sFolderSource, $sFolderTarget) {

        Rename-Item $sFolderSource $sFolderTarget
    
    } #Переименовать директорию

    fCopyFilesToFolder($sFromFolderPath, $sToFolderPath) {
        
        [string]$sFromFolderPath = $sFromFolderPath + "\*.*"
        Copy-Item $sFromFolderPath -Destination $sToFolderPath
    
    } #Множественное копирование файлов в директории

    [string]fTkvitBody($sReplyPath) {
    
        [xml]$xmlfile = Get-Content $sReplyPath
        [string]$innerText = $xmlfile.InnerXML
    
        if ($innerText){
      
            return $innerText
      
        }else{
        
            return $null

        }
    
    }  #Разбор квитков из ЦБ (из XML в текст)

    fMoveFile($sFilePath, $sTargetFolder) {
    
        Move-Item $sFilePath $sTargetFolder
    
    } #Переместить файл

} #Операции над файлами (копировать/удалить/переименовать)

Class Run_exe{

    [string]$sPS = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

    fRunExe($sProgram, $sArgs) {
       
        $oProcess = New-Object System.Diagnostics.Process
        $oProcess.StartInfo.FileName = $sProgram
        $oProcess.StartInfo.Arguments = $SArgs
        $oProcess.StartInfo.RedirectStandardOutput = $true
        $oProcess.StartInfo.UseShellExecute = $false
        $oProcess.Start()
        $oProcess.WaitForExit()

        [string]$sProcessOut = $oProcess.StandardOutput.ReadToEnd()

    } #Запуск любого EXE файла

    [string]$sVFDImagesPath = "C:\example"
    [string]$sVFD = "imdisk"
    
    $aVFDImages = @("example.img", "example.img", "example.img")
    
    fMountVFDImage($sVFDImage){
        $this.fRunExe($this.sVFD," -D -m B:\")| Out-Null
        $this.fRunExe($this.sVFD," -a -f " + $this.sVFDImagesPath + "\" + $sVFDImage + " -m B:\")| Out-Null
    
    } #Монтирование виртуальной дискеты с ключом

    [string]$sRAR = "example\WinRAR.exe"

    fUnzipReport($sReportPath) {

        [string]$sPath = Split-Path $sReportPath -Parent
        $this.fRunExe($this.sRAR," x " + $sReportPath + " " + $sPath) | Out-Null
    } #Разархивировать файл  с помощью WinRAR

} #Запуск программ 