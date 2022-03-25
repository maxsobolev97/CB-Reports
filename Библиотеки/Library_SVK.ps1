<#
    Библиотека для работы с каналом SVK
#>

Class SVK_operations{

    [bool]fGetAuth() {
 
	    [string[]]$aCommands = @("example","example")
	    [string]$sHost = "example"
	    [string]$sPort = "example"
	    [int]$iWait = 500
        [string]$sResult = ""

	    try{

            $oSocket = New-Object System.Net.Sockets.TcpClient($sHost, $sPort)

	        if($oSocket) {

		        $oStream = $oSocket.GetStream()
		        $oBuffer = New-Object System.Byte[] 1024
		        $oEncoding = New-Object System.Text.AsciiEncoding
	            $oWriter = New-Object System.IO.StreamWriter($oStream)

                ForEach($sCommand in $aCommands) {

                    $oWriter.WriteLine($sCommand)
                    $oWriter.Flush()
                    Start-Sleep -m ($iWait * 4)

                }

	        }
            
            return $true

        } catch {
        
            return $false

        }

    } #Авторизация на канале SVK

    [bool]fNeedAuth() {

	    [string]$sHost = "example"
	    [string]$sPort = "example"
	    [int]$iWait = 500
        [string]$sResult = ""
        [bool]$bReturn = $true

	    $oSocket = New-Object System.Net.Sockets.TcpClient($sHost, $sPort)
	    if($oSocket) {
		    $oStream = $oSocket.GetStream()
		    $oBuffer = New-Object System.Byte[] 1024
		    $oEncoding = New-Object System.Text.AsciiEncoding
	        Start-Sleep -m ($iWait * 4)
            While($oStream.DataAvailable) {
                $oRead = $oStream.Read($oBuffer, 0, 1024) 
                $sResult += ($oEncoding.GetString($oBuffer, 0, $oRead))
            }
	    }
        if($sResult.Contains("Must")) {
            $bReturn = $False
        }
        return $bReturn
    } #Проверка требуется ли авторизация на канале SVK

    fCheckAuth() {

        [bool]$bNeedAuth = $this.fNeedAuth()

        while($bNeedAuth -eq $true) {

            $this.fGetAuth()
            $bNeedAuth = $this.fNeedAuth()

        }

    } #Если не авторизован - авторизоваться

} #Операции с каналом SVK