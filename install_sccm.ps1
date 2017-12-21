$CREDENTIALS = ""
$SPLASH_HTA_PATH = ""

if (Test-Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe") {
    Start-Process "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe" -Credential $CREDENTIALS        
} else {
    if (Test-Path "C:\Windows\ccmsetup") {
        cd 'C:\Windows\ccmsetup'
        ccmsetup.exe /uninstall
        ccmsetup.exe /remove
    }
    Start-Process $SPLASH_HTA_PATH
}
