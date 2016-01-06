####################################################################################################################
#
#  Name: PSSonosControlerPreview.ps1
#  Author: @SimonDettling (www.msitproblog.com)
#  Version: 0.1
#  Disclaimer: This Script is not extensively tested and is a Preview of what's coming. Use at your own risk!
#
###################################################################################################################

# Enter the IP Adress of your Sonos Component, that is connect via Ethernt. (e.g. Playbar)
$sonosIP = "192.168.1.55"

# Port that is used for communication (Default = 1400)
$port = 1400

# Hash table containing SOAP Commands
$soapCommandTable = @{
    "Pause" = @{
        "path" = "/MediaRenderer/AVTransport/Control"
        "soapAction" = "urn:schemas-upnp-org:service:AVTransport:1#Pause"
        "message" =  '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:Pause xmlns:u="urn:schemas-upnp-org:service:AVTransport:1"><InstanceID>0</InstanceID></u:Pause></s:Body></s:Envelope>'
    }
    "Play" = @{
        "path" = "/MediaRenderer/AVTransport/Control"
        "soapAction" = "urn:schemas-upnp-org:service:AVTransport:1#Play"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:Play xmlns:u="urn:schemas-upnp-org:service:AVTransport:1"><InstanceID>0</InstanceID><Speed>1</Speed></u:Play></s:Body></s:Envelope>'
    }
    "Next" = @{
        "path" = "/MediaRenderer/AVTransport/Control"
        "soapAction" = "urn:schemas-upnp-org:service:AVTransport:1#Next"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:Next xmlns:u="urn:schemas-upnp-org:service:AVTransport:1"><InstanceID>0</InstanceID></u:Next></s:Body></s:Envelope>'
    }
    "Previous" = @{
        "path" = "/MediaRenderer/AVTransport/Control"
        "soapAction" = "urn:schemas-upnp-org:service:AVTransport:1#Previous"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:Previous xmlns:u="urn:schemas-upnp-org:service:AVTransport:1"><InstanceID>0</InstanceID></u:Previous></s:Body></s:Envelope>'
    }
    "Rewind" = @{
        "path" = "/MediaRenderer/AVTransport/Control"
        "soapAction" = "urn:schemas-upnp-org:service:AVTransport:1#Seek"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:Seek xmlns:u="urn:schemas-upnp-org:service:AVTransport:1"><InstanceID>0</InstanceID><Unit>REL_TIME</Unit><Target>00:00:00</Target></u:Seek></s:Body></s:Envelope>'
    }
    "RepeatAll" = @{
        "path" = "/MediaRenderer/AVTransport/Control"
        "soapAction" = "urn:schemas-upnp-org:service:AVTransport:1#SetPlayMode"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:SetPlayMode xmlns:u="urn:schemas-upnp-org:service:AVTransport:1"><InstanceID>0</InstanceID><NewPlayMode>REPEAT_ALL</NewPlayMode></u:SetPlayMode></s:Body></s:Envelope>'
    }
    "RepeatOne" = @{
        "path" = "/MediaRenderer/AVTransport/Control"
        "soapAction" = "urn:schemas-upnp-org:service:AVTransport:1#SetPlayMode"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:SetPlayMode xmlns:u="urn:schemas-upnp-org:service:AVTransport:1"><InstanceID>0</InstanceID><NewPlayMode>REPEAT_ONE</NewPlayMode></u:SetPlayMode></s:Body></s:Envelope>'
    }
    "RepeatOff" = @{
        "path" = "/MediaRenderer/AVTransport/Control"
        "soapAction" = "urn:schemas-upnp-org:service:AVTransport:1#SetPlayMode"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:SetPlayMode xmlns:u="urn:schemas-upnp-org:service:AVTransport:1"><InstanceID>0</InstanceID><NewPlayMode>NORMAL</NewPlayMode></u:SetPlayMode></s:Body></s:Envelope>'
    }
    "SetVolume" = @{
        "path" = "/MediaRenderer/RenderingControl/Control"
        "soapAction" = "urn:schemas-upnp-org:service:RenderingControl:1#SetVolume"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:SetVolume xmlns:u="urn:schemas-upnp-org:service:RenderingControl:1"><InstanceID>0</InstanceID><Channel>Master</Channel><DesiredVolume>###DESIRED_VOLUME###</DesiredVolume></u:SetVolume></s:Body></s:Envelope>'
    }
    "Mute" = @{
        "path" = "/MediaRenderer/RenderingControl/Control"
        "soapAction" = "urn:schemas-upnp-org:service:RenderingControl:1#SetMute"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:SetMute xmlns:u="urn:schemas-upnp-org:service:RenderingControl:1"><InstanceID>0</InstanceID><Channel>Master</Channel><DesiredMute>1</DesiredMute></u:SetMute></s:Body></s:Envelope>'
    }
    "Unmute" = @{
        "path" = "/MediaRenderer/RenderingControl/Control"
        "soapAction" = "urn:schemas-upnp-org:service:RenderingControl:1#SetMute"
        "message" = '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/"><s:Body><u:SetMute xmlns:u="urn:schemas-upnp-org:service:RenderingControl:1"><InstanceID>0</InstanceID><Channel>Master</Channel><DesiredMute>0</DesiredMute></u:SetMute></s:Body></s:Envelope>'
    }        
}

Function Set-SonosController {
    # Get Parameters
    Param(
        [Parameter(Mandatory=$true,Position=0,valueFromPipeline=$true)]
        [string]
        $action,

        [Parameter(Mandatory=$false,Position=1,valueFromPipeline=$true)]
        [int]
        $volume
    )

    # Get action from Hash Table, and throw error if it does not exist
    $actionHandler = $soapCommandTable.GetEnumerator() | Where-Object {$_.Key -eq $action}
    If (!$actionHandler) {
        throw "Action '$action' can not be found in Hash Table."
    }
    
    # Assign values from Hash Table
    $uri = "http://${sonosIP}:$port$($actionHandler.Value.path)"
    $soapAction = $actionHandler.Value.soapAction
    $soapMessage = $actionHandler.Value.message

    # Section for special Actions
    Switch ($action) {
        'setVolume' {
            If ($volume -gt 60) {
                $volume = 60
            }
            $soapMessage = $soapMessage.Replace("###DESIRED_VOLUME###", $volume)
        }
    }

    # Create SOAP Request
    $soapRequest = [System.Net.WebRequest]::Create($uri)

    # Set Headers
    $soapRequest.Accept = 'gzip'
    $soapRequest.Method = 'POST'
    $soapRequest.ContentType = 'text/xml; charset="utf-8"'
    $soapRequest.KeepAlive = $false
    $soapRequest.Headers.Add("SOAPACTION", $soapAction)

    # Sending SOAP Request
    $requestStream = $soapRequest.GetRequestStream()
    $soapMessage = [xml] $soapMessage
    $soapMessage.Save($requestStream)
    $requestStream.Close()

    # Sending Complete, Get Response
    $response = $soapRequest.GetResponse()
    $responseStream = $response.GetResponseStream()
    $soapReader = [System.IO.StreamReader]($responseStream)
    $returnXml = [Xml] $soapReader.ReadToEnd() 
    $responseStream.Close()
}

Write-Host "###############################################################"
Write-Host "#                                                             #"
Write-Host "#            SONOS PowerShell Controller (Preview)            #"
Write-Host "#                       msitproblog.com                       #"
Write-Host "#                                                             #"
Write-Host "###############################################################"
Write-Host ""
Write-Host ""
Write-Host "[1] Play"
Write-Host "[2] Pause"
Write-Host "[3] Previous Track"
Write-Host "[4] Next Track"
Write-Host "[5] Rewind Track"
Write-Host "[6] Mute"
Write-Host "[7] Unmute"
Write-Host "[8] Repeat All"
Write-Host "[9] Repeat One"
Write-Host "[10] Repeat Off"
Write-Host "[11] Set Volume"
Write-Host "----------------"
Write-Host "[99] Exit"
Write-Host "----------------"
Write-Host ""

While (1) {
    Switch (Read-Host "Please select an Action") {
     1 {Set-SonosController "Play"}
     2 {Set-SonosController "Pause"}
     3 {Set-SonosController "Previous"}
     4 {Set-SonosController "Next"}
     5 {Set-SonosController "Rewind"}
     6 {Set-SonosController "Mute"}
     7 {Set-SonosController "Unmute"}
     8 {Set-SonosController "RepeatAll"}
     9 {Set-SonosController "RepeatOne"}
     10 {Set-SonosController "RepeatOff"}
     11 {
        $volume = Read-Host "Enter Volume (1-50)"
        Set-SonosController "SetVolume" $volume
    }
     99 {Exit}
    }
}