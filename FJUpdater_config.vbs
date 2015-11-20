'===================================================================================================
' settings to send mail with CDO (used with key /mail+)
' ��������� �������� ����� ����� CDO (��������� � ������ /mail+)
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
' ����, ���� ��������� ���������� (��� /WEBMode+ � WEBModeSaveInstall+),
' ��� ������ ������������� ��� ����������� ��������� (/WEBMode-)
' �� ������������ ������������ - RW ��� ������� (reference computer), RO - ��� ��������
Const csInstallerPath 		= "\\SERVER\INSTALL_CLIENT$\FJUpdater\"
' path for log-files, RW for all
' ����� ��� ���������� ���-������, ������ ���� �������� �� ������
Const csLogsPath 		= "\\SERVER\INSTALL_CLIENT$\FJUpdater.LOG\"
