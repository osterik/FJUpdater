'===================================================================================================
' ��������� �������� ����� ����� CDO (��������� � ������ /mail+)
Const csTo              = "xxxxx@xxxxx.xxx"
Const csFrom            = "xxxxx@xxxxx.xxx"
Const csSMTPServer      = "mail.xxxxx.xxx"
Const csSendUserName    = "xxxxx@xxxxx.xxx"
Const csSendPassword    = "xxxxx"
Const ciSMTPServerPort  = 25
Const cbSMTPUseSSL      = false
Const ciSMTPAuthenticate= 1 ' 0 =anonymous, Do not authenticate
                            ' 1 = basic (clear-text) authentication
                            ' 2 = NTLM
Const ciSendUsing       = 2 ' 1 = cdoSendUsingPickup, send message using the local SMTP service pickup directory.
                            ' 2 = cdoSendUsingPort, send the message using the network (SMTP over the network).
                            ' 3 = cdoSendUsingExchange, send the message using MS Exchange Mail Server.
'===================================================================================================
' ����, ���� ��������� ���������� (��� /WEBMode+ � WEBModeSaveInstall+),
' ��� ������ ������������� ��� ����������� ��������� (/WEBMode-)
Const csInstallerPath 	= "\\Network\Update\"
' ���� ���� ���������� �������� ������ ����
Const ReportFilePath = "\\Network\Update\"