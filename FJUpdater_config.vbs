'===================================================================================================
' настройки отправки почты через CDO (актуальны с ключем /mail+)
Const csTo              = "admin@domain.local"
Const csFrom            = "fjupdater@domain.local"
Const csSMTPServer      = "mail.domain.local"
Const csSendUserName    = ""
Const csSendPassword    = ""
Const ciSMTPServerPort  = 25
Const cbSMTPUseSSL      = false
Const ciSMTPAuthenticate= 0 ' 0 =anonymous, Do not authenticate
                            ' 1 = basic (clear-text) authentication
                            ' 2 = NTLM
Const ciSendUsing       = 2 ' 1 = cdoSendUsingPickup, send message using the local SMTP service pickup directory.
                            ' 2 = cdoSendUsingPort, send the message using the network (SMTP over the network).
                            ' 3 = cdoSendUsingExchange, send the message using MS Exchange Mail Server.
'===================================================================================================
' пути, куда сохранять установщик (при /WEBMode+ и WEBModeSaveInstall+),
' или откуда устанавливать уже загруженные программы (/WEBMode-)
Const csInstallerPath 		= "\\SERVER\INSTALL_CLIENT$\FJUpdater\"
