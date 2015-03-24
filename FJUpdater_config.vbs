'===================================================================================================
' settings to send mail with CDO (used with key /mail+)
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
' path where to save the installers (used with keys /WEBMode+ & WEBModeSaveInstall+),
' or from install already downloaded programs (with key /WEBMode-)
' for security purposes - RW for server (reference computer), RO - for other computers
' пути, куда сохранять установщик (при /WEBMode+ и WEBModeSaveInstall+),
' или откуда устанавливать уже загруженные программы (/WEBMode-)
' по соображениям безопасности - RW для сервера (reference computer), RO - для клиентов
Const csInstallerPath 		= "\\SERVER\INSTALL_CLIENT$\FJUpdater\"
