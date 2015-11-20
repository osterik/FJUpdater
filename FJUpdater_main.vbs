'-----------------------------------------------------------------------------------------------------------------------
'
' Copyright (c) Ilya Kisleyko. All rights reserved.
' 
' AUTHOR  : Ilya Kisleyko
' E-MAIL  : osterik@gmail.com
' DATE    : 01.08.2015
' NAME    : FJUpdater_main.vbs
' COMMENT : ������ ��� �������� ������������ ������������� �� ���������� ������ Java&Flash
'          � ��������������� ���������� (����� ��������(HTTP) ��� ��������� ����(SMB)) ��� �������������.
'          ������ ����� ��������� ����������� ���������� ������ Java&Flash � ��������� ����� (+ ���� � ������� ������),
'          ��� ��� ����������� ������� �������� ��������������� ������������ ���������� Java&Flash ����� GPO
'          (������� ������ �����������).
'          �� ������� � ����������� ����� ������������ ����� ��������� ����� �� �����.
'                                                                                                                       
'��� ������� ��� ���������� ������ ��������� ���������� ������������� �������� ����� ��������
'
'��������� ��������� ������:
'/mail:1 = ���������� ����� ��� ������� ������ ��� ��� ������� ����������
'/mail:0 = ����� ����� (default)
'
'/LogFile:0 = �� ��������� ��� (default)
'/LogFile:1 = ������� �������, ���� � ��������� csLogsPath, ��� ����� = ��� ����������, ����� ��� ������ �������,
'/LogFile:2 = ������� �������, ���� � ��������� csLogsPath, ��� ����� = ��� ����������, �� ������� ������ ������
'
'/debug:0 = �� �������� ������
'/debug:1 = ���������� ������ ��������� � ����������� ��������� (default) � ������
'/debug:2 = ���������� ��������� ���������
'/debug:3 = ���������� ��������� ��������� � ���������� ��������� � �� ���������
'
'/WebMode:1 = ��������� ���������� �������� �� ��������� (default)
'/WebMode:0 = �� ��������� ����� (csInstallerPath)
'
'/WEBModeSaveInstall:1 = ��������� ���������� � ��������� ����� (csInstallerPath)
'/WEBModeSaveInstall:0 = ��������� ���������� � %TEMP% (default)
'���� /WEBModeSaveInstall ����� ����� ������ ��� /WEBMode:1
'
'/WEBModeSaveInstallForce = ������ �����, ������������� ������� ���������� ������ ������������
'� ��������� � ��������� ����� (csInstallerPath), ���� ������������� ��� ������ �������� �� ���������
'� ��������������� �� ����� Oracle&Adobe
'������������ �� ���������� �������������� ��� ������� ��������� ��������������� ������������ ����������,
'� ����� ��� ������������� ��������, �� ��������� �� ����� ������ ������.
'
'/MailTest = ������ �����, �������� �������� �������� �����, ��������� �������� ���������
'
'/ShowVersion[:comp] = ������ �����, ������� ����� ����� ������������� �� ���������� �������� Java & Flash
' ��� ������� ��� ���������� - �� ��������� ����������, ���� ���� �������� - �� ���������� ��� ���\����� ����������
'
'/IgnoreJava        = �� ��������� Java Native (32 on 32, 64 on 64)
'/IgnoreJavaWoW6432 = �� ��������� Java WoW6432 (32bit on 64bit system)
'/IgnoreFlashA      = �� ��������� FlashA
'/IgnoreFlashP      = �� ��������� FlashP
'
'/UninstallJavaOld[:comp]  = ���������������� ������ ������ JRE & Java Auto Updater, �� ������ JDK
' ��� ������� ��� ���������� - �� ��������� ����������, ���� ���� �������� - �� ���������� ��� ���\����� ����������,
' � ����� ������ ������������ ������ ������ �������� (������ �����)

'���� ��� ������� ���������� ����������� �������� � ����������� "������ �����", �� �� ����������� ����������,
'� ������ ��� ������������ ������������� (mail & debug) � ������������ �����.
'
'/? ��� /help = ������� 
'-----------------------------------------------------------------------------------------------------------------------
'[HELP_END]
' ������ "'[HELP_END]" ������������ ��� ������������ ��� �������� ������ �������

Option Explicit

Const IsDebug = 1 ' ���� �������
If IsDebug <> 1 then On Error Resume Next

' ����� ������ � �������������\������� ���������� ������, ������� ����� � ��������� �����������
Const csJavaInstaller		= "java_installer.exe"
Const csJavaInstallerMSI	= "java_installer.msi"
Const csJavaInstaller64		= "java_installer64.exe"
Const csJavaInstallerVers	= "java_current.txt"
Const csFlashPInstaller		= "flashP_installer.exe"	' Flash Plugin (Firefox, Mozilla, Netscape, Opera)
Const csFlashPInstallerVers	= "flashP_current.txt"
Const csFlashAInstaller		= "flashA_installer.exe"	' Flash ActiveX (IE)
Const csFlashAInstallerVers	= "flashA_current.txt"

'===================================================================================================
' ��� �������� ������ ���������� ������ ((��� /WEBMode:1)
' �������� ����������� ������ Java �������. �� ����.
'Const csJavaVersionCurrentLnk = "http://www.java.com/applet/JreCurrentVersion2.txt"
Const csFlashVersionCurrentLnk = "http://www.adobe.com/software/flash/about/"
' � ��� ������\������ �����������
Const csJavaInstallerLink   = "http://java.com/en/download/windows_manual.jsp"
'Const csFlashPInstallerLink = "http://fpdownload.macromedia.com/pub/flashplayer/latest/help/install_flash_player.exe"
'Const csFlashAInstallerLink = "http://fpdownload.macromedia.com/pub/flashplayer/latest/help/install_flash_player_ax.exe"
'� ������ 17.0.0.188 - �� ��������, ������ ����������� ����� ������� sFlashInstallerLinkGet
Const csFlashInstallerLink = "http://fpdownload.macromedia.com/pub/flashplayer/pdc/"
' ����� ������� ������������
'Const csJavaInstallerParams = "/s /v /qn IEXPLORER=1 MOZILLA=1 REBOOT=ReallySuppress JAVAUPDATE=0 WEBSTARTICON=0"   ' JAVA 1.7
Const csJavaInstallerParams = "/s"     '  JAVA 1.8 code by https://github.com/airosa-id
Const csFlashInstallerParams = "/install"
Const csJavaInstallerParamsMSI = "/quiet"

'===================================================================================================
' ��� ������� � �������, ������� RegRead
Const HKEY_CLASSES_ROOT  	= &H80000000
Const HKEY_CURRENT_USER 	= &H80000001
Const HKEY_LOCAL_MACHINE 	= &H80000002
Const HKEY_USERS 		= &H80000003
Const HKEY_CURRENT_CONFIG  	= &H80000005

' ���������� ���������, ����� ���� �������������� �� ��������� ������
Dim glbMail 					' /mail:1 = ���������� ����� ��� ������� ������ ��� ��� ������� ���������� 
						' /mail:0 = ����� ����� (default)
Dim glbLogFile					'/LogFile:0 = �� ��������� ��� (default)
						'/LogFile:1 = ������� �������, ���� � ��������� csLogsPath, ��� ����� = ��� ����������, ����� ��� ������ �������,
						'/LogFile:2 = ������� �������, ���� � ��������� csLogsPath, ��� ����� = ��� ����������, �� ������� ������ ������

Dim glbDebug					' /debug:0 = �� �������� ������
						' /debug:1 = ���������� ������ ��������� � ����������� ��������� (default)
						' /debug:2 = ���������� ��������� ��������� 
						' /debug:3 = ���������� ��������� ��������� � ���������� ��������� � �� ���������
Dim glbWEBMode 					' /WEBMode:1 = ��������� ���������� �������� �� ��������� (default)
						' /WEBMode:0 = �� ��������� ����� (csInstallerPath)
Dim glbWEBModeSaveInstall			' ����� ����� ������ ��� /WEBMode:1
						' /WEBModeSaveInstall:1 = ��������� ���������� � ��������� ����� (csInstallerPath)
						' /WEBModeSaveInstall:0 =  ��������� ���������� � %TEMP% (default)
'Dim glbComputer				' ���(��� �����) ����������, �� ������� ��������� ������ ��������
Dim glbIgnoreJava				' �� ��������� Java Native (32 on 32, 64 on 64)
Dim glbIgnoreJavaWoW6432			' �� ��������� Java WoW6432 (32bit on 64bit system)
Dim glbIgnoreFlashA				' �� ��������� FlashA
Dim glbIgnoreFlashP				' �� ��������� FlashP
Dim glbUninstallJavaOld				' ���������������� ������ ������ JRE & Java Auto Updater, �� ������ JDK
' ���������� ����������
Dim sInstallerPath				' ���� ���������� ����������� ��� /WEBModeSaveInstall:1
Dim sLog 					' ����� ���, ���������� �� ����� ��� ������� ��� ��������� ����������
Dim bNeedToSendLog				' ���� ������������� �������� ���� �� �����
Dim sJavaNativeUpdateStatus, sJavaWoW6432UpdateStatus, sFlashAUpdateStatus, sFlashPUpdateStatus		' �������� ������ �������� ����������
Dim sJavaVersionCurrent, sFlashPVersionCurrent, sFlashAVersionCurrent ' ����� ������ ��� ��������� ����������, ������������ ��� ������ ����� � ������� ������

Dim objFSO
Dim objShell

Sub Main
	If IsDebug <> 1 then On Error Resume Next
	
	' Abort If the host is not cscript
	If not IsHostCscript() then
		call wscript.echo("������������� ��������� ������ ������ � �������������� CScript.")
		wscript.quit
	End If

	Call WriteLog(FormatDateTime(Now,2) & " JavaFlashUpdaterAdmin v2.9 (c) Ilya Kisleyko, osterik@gmail.com", 2)

	' ���� ��� ������������� ���������� ���� �� �����
	bNeedToSendLog = false

	' ������ ��������
	sJavaNativeUpdateStatus   = "UNKNOWN"
	sJavaWoW6432UpdateStatus  = "UNKNOWN"
	sFlashAUpdateStatus       = "UNKNOWN"
	sFlashPUpdateStatus       = "UNKNOWN"

	' Parse command-line parameters
	If bParseCommandLine = false Then
		' �� ������ ��������� ��������� ��������� ������
		Call WriteLog("Main, ERROR! can't parse command line switches!", 1)
		Exit Sub 
	End If

	'���� ���������� �����������...
	If glbWEBMode then 
		If glbWEBModeSaveInstall Then
			sInstallerPath = csInstallerPath	' ...��� /WEBModeSaveInstall:1
		Else
			sInstallerPath = sEnvGet("%temp%")	' ...��� /WEBModeSaveInstall:0
		End If
	else
		sInstallerPath = csInstallerPath		' ..��� ��������� ������
	End If
	
	' ��� ������������� ��������� ���� � ����� �������� ��������
	If Right(sInstallerPath,1) <> "\" Then 
		sInstallerPath = sInstallerPath + "\"
	End If
	If Right(csLogsPath,1) <> "\" Then
		csLogsPath = csLogsPath + "\"
	End If

if glbIgnoreJava = False Then
	Call WriteLog("---------------------------------------------------------------------------------------------------",2)
	Call WriteLog("JavaNative - START",2)
	' JAVA Native (32 on 32, 64 on 64) - ������� ������ ���������� � �������������
	' ��� ��������� �� ���������� ������ ...
	If glbWEBMode Then ' WWW
		sJavaVersionCurrent = sJavaVersionWEBGet()
	Else               ' ��������� ���������
		sJavaVersionCurrent = sJavaVersionLocalGet()
	End If
	If sJavaVersionCurrent <> "" then 
		Dim sJavaVersionInstalled
		sJavaVersionInstalled = sJavaVersionInstalledGet(".")
		If  sJavaVersionInstalled <> "" then ' �� ���������� ������������ �����-�� ������
			If sJavaVersionInstalled <> sJavaVersionCurrent then
				' �� ���������, ���������
				Call WriteLog("WARNING! Installed and current Native(32 on 32, 64 on 64) Java version is NOT identical",1)
				' ���� ����� ��������� �����
				bNeedToSendLog = True
				If glbWEBMode Then
					' �������� ������ �� ������ � ��������� ����������
					If Is64BitSystem (".") = true then
						Call HttpGetSave(sJavaGetLinkToDownload("64"), sInstallerPath & csJavaInstaller64) 
					else
						Call HttpGetSave(sJavaGetLinkToDownload("32"), sInstallerPath & csJavaInstaller) 
					End If
				End If
				If Is64BitSystem (".") = true then
					Call myRun(sInstallerPath & csJavaInstaller64 & " " & csJavaInstallerParams)
				else
					Call myRun(sInstallerPath & csJavaInstaller & " " & csJavaInstallerParams)
				End If
				Call WriteLog("Install complete, check Native(32 on 32, 64 on 64) Java version again",2)
			
				' �������� ��������� ���������
				If sJavaVersionInstalledGet(".") = sJavaVersionCurrent then
					sJavaNativeUpdateStatus = "SUCCESS"
					If glbWEBModeSaveInstall then
						' ���������� ���� � ������� ���������� ������ �����������
						Call myRun("cmd /c echo " & sJavaVersionCurrent & "> " & sInstallerPath & csJavaInstallerVers)
					End If
				Else
					sJavaNativeUpdateStatus = "FAILED"
				End If
				Call WriteLog("Java Update status = " & sJavaNativeUpdateStatus,1)
			Else
				Call WriteLog("Installed and current Java version is identical",2)
			End If
		End If
	End If
	Call WriteLog("JavaNative - FINISH",2)
End If

if glbIgnoreJavaWoW6432 = False Then
	Call WriteLog("---------------------------------------------------------------------------------------------------",2)
	Call WriteLog("JavaWoW6432 - START",2)
	' JAVA WoW (32bit on 64bit system) - ������� ������ ���������� � �������������
	' ��� ��������� �� ���������� ������ ...
	If glbWEBMode Then ' WWW
		sJavaVersionCurrent = sJavaVersionWEBGet()
	Else               ' ��������� ���������
		sJavaVersionCurrent = sJavaVersionLocalGet()
	End If
	If sJavaVersionCurrent <> "" then 
		Dim sJavaWoW6432VersionInstalled
		sJavaWoW6432VersionInstalled = sJavaWoW6432VersionInstalledGet(".")
		If  sJavaWoW6432VersionInstalled <> "" then ' �� ���������� ������������ �����-�� ������
			If sJavaWoW6432VersionInstalled <> sJavaVersionCurrent then
				' �� ���������, ���������
				Call WriteLog("WARNING! Installed and current WoW6432(32bit Java on 64bit OS) Java version is NOT identical",1)
				' ���� ����� ��������� �����
				bNeedToSendLog = True
				If glbWEBMode Then
					' �������� ������ �� ������ � ��������� ����������
					Call HttpGetSave(sJavaGetLinkToDownload(""), sInstallerPath & csJavaInstaller)
				End If
				' �������������, ���� ��������� ���� � ���������� glbWEBModeSaveInstall �� ����������� ������� �������, ��� ��� ��� ����������� ������ ��-��� ������� ������
				' ������������������ ������������, � ���� ���������� ����������, �� ����������� msi �����. ��� �������� � ���, ��� ��� ���������� ��������� java �� ����� �������,
				' � ��� ������ � ���������� ��� ������� �� ������� �������, ���������� ���� ������������� msi ���� � ���� 
				'C:\Windows\SysWOW64\config\systemprofile\AppData\LocalLow\, � ���� ��� � �������� C:\Windows\system32\config\systemprofile\AppData\LocalLow\

				If glbWEBModeSaveInstall then
					Call WriteLog("Installing EXE package",2)
					Call myRun(sInstallerPath & csJavaInstaller & " " & csJavaInstallerParams)
				else
					Set objFSO = CreateObject("Scripting.FileSystemObject")
					If objFSO.FileExists(sInstallerPath & csJavaInstallerMSI) then
						Call WriteLog("Installing MSI package",2)
						Call myRun("msiexec /i " & sInstallerPath & csJavaInstallerMSI & " " & csJavaInstallerParamsMSI)
					else
						Call WriteLog("The MSI package is missing",2)
					End If
					Set objFSO = Nothing
				End If

				Call WriteLog("Install complete, check WoW6432(32bit Java on 64bit OS) Java version again",2)

				' �������� ��������� ���������
				If sJavaWoW6432VersionInstalledGet(".") = sJavaVersionCurrent then
					sJavaWoW6432UpdateStatus = "SUCCESS"
					If glbWEBModeSaveInstall then
						' ���������� ���� � ������� ���������� ������ �����������
						Call myRun("cmd /c echo " & sJavaVersionCurrent & "> " & sInstallerPath & csJavaInstallerVers)
						' �������� ������������ ������ java � ������� �����������
						Call WriteLog("Copying MSI package to share directory",2)
						Set objShell = CreateObject("WScript.Shell")
						Dim AppData
						AppData = objShell.expandEnvironmentStrings("%APPDATA%")
						Set objFSO = CreateObject("Scripting.FileSystemObject")
						Call WriteLog (AppData & "\..\LocalLow\Sun\Java\jre" & sJavaVersionCurrent & "\jre" & sJavaVersionCurrent & ".msi to " & sInstallerPath & csJavaInstallerMSI,3)
						objFSO.CopyFile AppData & "\..\LocalLow\Sun\Java\jre" & sJavaVersionCurrent & "\jre" & sJavaVersionCurrent & ".msi", sInstallerPath & csJavaInstallerMSI, True
						Set objFSO = Nothing
						Set objShell = Nothing
					End If
				Else
					sJavaWoW6432UpdateStatus = "FAILED"
				End If
				Call WriteLog("JavaWoW6432 Update status = " & sJavaWoW6432UpdateStatus,1)
			Else
				Call WriteLog("Installed and current JavaWoW6432 version is identical",2)
			End If		
		End If
	End If
	Call WriteLog("JavaWoW6432 - FINISH",2)
End If

if glbUninstallJavaOld = true Then
	Call WriteLog("---------------------------------------------------------------------------------------------------",2)
	Call WriteLog("UninstallJavaOld - START",2)
	Call UninstallJavaOld(".")
	Call WriteLog("UninstallJavaOld - FINISH",2)
End If
	
if glbIgnoreFlashA = False Then
	Call WriteLog("---------------------------------------------------------------------------------------------------",2)
	Call WriteLog("FlashActiveX - START",2)
	' FlashActiveX - ������� ������ ���������� � �������������
	' ��� ��������� �� ���������� ������ ...	
	If glbWEBMode Then
		sFlashAVersionCurrent = sFlashVersionWEBGet("A")
	Else
		sFlashAVersionCurrent = sFlashVersionLocalGet("A")
	End If
	
	If sFlashAVersionCurrent <> "" then 
		If sFlashVersionInstalledGet(".", "A") <> "" then ' �� ���������� ������������ �����-�� ������
			If sFlashVersionInstalledGet(".", "A") <> sFlashAVersionCurrent then
				' �� ���������, ���������
				Call WriteLog("WARNING! Installed and current FlashActiveX version is NOT identical",1)
				' ���� ����� ��������� �����
				bNeedToSendLog = True
				If glbWEBMode Then
					' �������� ������ �� ������ � ��������� ����������
					Call HttpGetSave(sFlashInstallerLinkGet("A",sFlashAVersionCurrent), sInstallerPath & csFlashAInstaller)
				End If
				' �������������
				Call myRun(sInstallerPath & csFlashAInstaller & " " & csFlashInstallerParams)
				' �������� ��������� ���������
				If sFlashVersionInstalledGet(".", "A") = sFlashAVersionCurrent then
					sFlashAUpdateStatus = "SUCCESS"
					' ���������� ���� � ������� ���������� ������ �����������
					Call myRun("cmd /c echo " & sFlashAVersionCurrent & "> " & sInstallerPath & csFlashAInstallerVers)
				Else
					sFlashAUpdateStatus = "FAILED"
				End If
				Call WriteLog("FlashActiveX Update status = " & sFlashAUpdateStatus,1)
			Else
				Call WriteLog("Installed and current FlashActiveX version is identical",2)
			End If		
		End If
	End If
	Call WriteLog("FlashActiveX - FINISH",2)
End If

if glbIgnoreFlashP = False Then
	Call WriteLog("---------------------------------------------------------------------------------------------------",2)
	Call WriteLog("FlashPlugin - START",2)
	' FlashPlugin - ������� ������ ���������� � �������������
	' ��� ��������� �� ���������� ������ ...	
	If glbWEBMode Then
		sFlashPVersionCurrent = sFlashVersionWEBGet("P")
	Else
		sFlashPVersionCurrent = sFlashVersionLocalGet("P")
	End If
	If sFlashPVersionCurrent <> "" Then 
		If sFlashVersionInstalledGet(".", "P") <> "" then ' �� ���������� ������������ �����-�� ������
			If sFlashVersionInstalledGet(".", "P") <> sFlashPVersionCurrent then
				' �� ���������, ���������
				Call WriteLog("WARNING! Installed and current FlashPlugin version is NOT identical",1)
				' ���� ����� ��������� �����
				bNeedToSendLog = True
				If glbWEBMode Then
					' �������� ������ �� ������ � ��������� ����������
					Call HttpGetSave(sFlashInstallerLinkGet("P",sFlashPVersionCurrent), sInstallerPath & csFlashPInstaller)
				End If
				' �������������
				Call myRun(sInstallerPath & csFlashPInstaller & " " & csFlashInstallerParams)
				' �������� ��������� ���������
				If sFlashVersionInstalledGet(".", "P") = sFlashPVersionCurrent then
					sFlashPUpdateStatus = "SUCCESS"
					' ���������� ���� � ������� ���������� ������ �����������
					Call myRun("cmd /c echo " & sFlashPVersionCurrent & "> " & sInstallerPath & csFlashPInstallerVers)
				Else
					sFlashPUpdateStatus = "FAILED"
			End If
				Call WriteLog("FlashPlugin Update status = " & sFlashPUpdateStatus,1)
			Else
			Call WriteLog("Installed and current FlashPlugin version is identical",2)
			End If		
		End If		
	End If
	Call WriteLog("FlashPlugin - FINISH",2)
End If

	' ��� ������������� ���������� �����
	Call SendLog()

	' ��� ������������� ���������� ����� � ���-����
	Call WriteLogToFile()

	Call WriteLog("===================================================================================================",2)
End Sub ' Main

Sub UninstallJavaOld(sComputer)
	'uninstall old JRE & Java Auto Updater, skip JDK
	If IsDebug <> 1 then On Error Resume Next
	if sComputer = "" then sComputer = "."

	'parse installed java version
	dim sJavaVersionInstalled
	dim sjvMajor1, sjvMajor2, sjvMiddle, sjvMinor               ' sample:
	sJavaVersionInstalled = sJavaVersionInstalledGet(sComputer) ' = 1.8.0_51
	If sJavaVersionInstalled = "" then
		Call WriteLog("UninstallJavaOld : there is no JAVA on computer " & sComputer,2)
		Exit Sub
	End If
	sjvMajor1 = ExtractWord(0, sJavaVersionInstalled, caDelims) ' = 1
	sjvMajor2 = ExtractWord(1, sJavaVersionInstalled, caDelims) ' = 8
	sjvMiddle = ExtractWord(2, sJavaVersionInstalled, caDelims) ' = 0
	sjvMinor  = ExtractWord(3, sJavaVersionInstalled, caDelims) ' = 51

	'query installed softwate from WMI
	Dim objWMI, colJava, objJava
	Dim s,s1,s2,s3,s4
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
	'Set colJava = objWMI.ExecQuery ("SELECT * FROM Win32_Product WHERE ( name LIKE '%java 7%' AND name LIKE '%update%' )")
	Set colJava = objWMI.ExecQuery ("SELECT * FROM Win32_Product WHERE name LIKE '%java%'")
	For Each objJava in colJava
		s = objJava.Caption

		' uninstall 'Java Auto Updater'
		if s = "Java Auto Updater" then
			Call WriteLog("UninstallJavaOld : UNINSTALL Begin  : " & s,1)
			objJava.Uninstall
			Call WriteLog("UninstallJavaOld : UNINSTALL End    : " & s,1)
		end if

		'
		s1 = ExtractWord(0, s, caDelims)
		If lCase(s1) = "java" then
			s2 = ExtractWord(1, s, caDelims)
			If lCase(s2) <> "se" then ' it seems to be a JAVA Development Kit, skipping
				If lCase(s2) = "tm" then ' it seems to be a JAVA 1.6
					s2 = ExtractWord(2, s, caDelims)
				End If
				if s2 < sjvMajor2 then ' previous major release
					Call WriteLog("UninstallJavaOld : UNINSTALL Begin  : " & s,1)
					objJava.Uninstall
					Call WriteLog("UninstallJavaOld : UNINSTALL End    : " & s,1)
				End If
				if s2 = sjvMajor2 then ' current major release
					s3 = ExtractWord(2, s, caDelims)
					if lCase(s3) = "update" then
						s4 = ExtractWord(3, s, caDelims)
						if s4 < sjvMinor then ' previous minor release
							Call WriteLog("UninstallJavaOld : UNINSTALL Begin  : " & s,1)
							objJava.Uninstall
							Call WriteLog("UninstallJavaOld : UNINSTALL End    : " & s,1)
						Elseif s4 = sjvMinor then ' current release
							Call WriteLog("UninstallJavaOld : current release  : " & s,2)
						Else
							Call WriteLog("UninstallJavaOld : unknown          : " & s,2)
						End If
					Else
						Call WriteLog("UninstallJavaOld : unknown          : " & s,2)
					end if
				End If
			else
				Call WriteLog("UninstallJavaOld : JDK, skipping    : " & s,2)
			End If
		End If
	Next
End Sub 'UninstallJavaOld

Sub SendLog
	dim sComputerName
	sComputerName = sEnvGet("%COMPUTERNAME%")

	If bNeedToSendLog and glbMail then
		Call WriteLog("Sending log to e-mail",2)
		Call mySendMail("Updater Java=[" & sJavaNativeUpdateStatus &_
			"], JavaWoW6432 = [" & sJavaWoW6432UpdateStatus &_
			"], FlashA =[" & sFlashAUpdateStatus & "], FlashP =[" & sFlashPUpdateStatus &_
			"] on " & sComputerName, sLog)
	End If
End Sub

Sub WriteLogToFile
	dim sComputerName
	sComputerName = sEnvGet("%COMPUTERNAME%")

	if glbLogFile > 0 then
		Dim objFSO,objFile,sLogFile

		sLogFile = csLogsPath & sComputerName & ".log"
		Call WriteLog("Writing log to " & sLogFile,2)

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Select Case glbLogFile
		Case 1
			Set objFile = objFSO.OpenTextFile(sLogFile, 2, true) ' ForWriting
		Case 2
			Set objFile = objFSO.OpenTextFile(sLogFile, 8, true) ' ForAppending
		End Select

		objFile.WriteLine(sLog)
		objFile.WriteLine("===================================================================================================")
		objFile.WriteLine("")
		objFile.Close
		Set objFSO = Nothing
	End If
End Sub

Function sJavaGetLinkToDownload (sBits)
	' ������ ��������� csJavaInstallerLink
	' � ���������� ������ �� ������ ���������� ������ Java
	' ���� sBits = 64, �� ���������� ������ �� 64-������ ������, ����� = 32�
	If IsDebug <> 1 then On Error Resume Next
	Dim sHttpText				' ���������� ��������� ��������
	Dim objRegExp				' ��� ��������� ������ �� ����������
	Dim objMatches, objMatch		' ��������������� ������� ��� �������� html
	
	' ��������� �������� ������ ����� ��� ������ �������� Java � sHttpText
	Call WriteLog("sJavaGetLinkToDownload(" & sBits &"), Downloadning HTML",3)
	sHttpText = sHttpGet(csJavaInstallerLink)
	If sHttpText = "" then
		Call WriteLog("sJavaGetLinkToDownload, ERROR! Failed to get the html by the link " & csJavaInstallerLink,1)
		bNeedToSendLog = True
		Exit Function
	End If
	
	' �������� ��������� HTML, ���� ��� <a>
	Call WriteLog("sJavaGetLinkToDownload, Parsing HTML, try to find ""<a title=""Download Java... """,3)
	Set objRegExp = CreateObject("VBScript.RegExp")
	If sBits = "64" then 
		objRegExp.Pattern= "<a\s+title=""Download Java software for Windows \(64-bit(.*)>"
	Else
		objRegExp.Pattern= "<a\s+title=""Download Java software for Windows Offline(.*)>"
	End If
	objRegExp.Global = True
	Set objMatches = objRegExp.Execute(sHttpText)
	If objMatches.Count = 0 then
		Call WriteLog("sJavaGetLinkToDownload, ERROR! Failed to found ""<a title=""Download Java... """,3)
		If Err.Number <> 0 then Call WriteLogError("sJavaGetLinkToDownload", Err.Number, Err.Description, Err.Source)
		bNeedToSendLog = True
		Exit Function
	Else
		Set objMatch = objMatches.Item(0)
		sHttpText = objMatch.Value
		'���� href
		Call WriteLog("sJavaGetLinkToDownload, Parsing HTML, try to find ""href""",3)
		objRegExp.Pattern= "href=""http://(.*?)"""
		Set objMatches = objRegExp.Execute(sHttpText)
		If objMatches.Count = 0 then
			Call WriteLog("sJavaGetLinkToDownload, ERROR! Failed to found ""href""",1)
			If Err.Number <> 0 then Call WriteLogError("sJavaGetLinkToDownload", Err.Number, Err.Description, Err.Source)
			bNeedToSendLog = True
			Exit Function
		Else
			Set objMatch = objMatches.Item(0)
			' �������� ������ �������
			sHttpText=Mid(objMatch.Value,7,len(objMatch.Value)-7)
			Call WriteLog("sJavaGetLinkToDownload, OK! Download link = " & sHttpText,3)
			sJavaGetLinkToDownload = sHttpText
		End If
	End If
End Function ' sJavaGetLinkToDownload

Function sJavaVersionInstalledGet (strComputer)
	' ���������� ������ � ������� ������ Java, ������������� �� ���������� strComputer, ��� ��������� ���������� (���� �����)
	' ���������� ������ ����������� � ������������ ������� (32 ��� 32, 64 ��� 64)
	sJavaVersionInstalledGet = sJavaVersionInstalledGetA (strComputer, "")
End Function ' sJavaVersionInstalledGet

Function sJavaWoW6432VersionInstalledGet (strComputer)
	' ���������� ������ � ������� ������ Java, ������������� �� ���������� strComputer, ��� ��������� ���������� (���� �����)
	' � sWoW6432 = ���������� ������ 32-������ Jav� �� 64-������ �������
	' ������ �� 64������ ��������
	
	If Is64BitSystem(strComputer) then
		sJavaWoW6432VersionInstalledGet = sJavaVersionInstalledGetA (strComputer, "\Wow6432Node")
	Else
		sJavaWoW6432VersionInstalledGet = ""
	End If
End Function ' sJavaWoW6432VersionInstalledGet

Function sJavaVersionInstalledGetA (strComputer, sRegPathModifier)
	'sRegPathModifier -  ����������� � ����, ���� ����������� 32������ ������ �� 64������ �������. 
	' ������ ���� = "\Wow6432Node"

	If IsDebug <> 1 then On Error Resume Next
	
	Dim sJavaMajorVersionInstalled, sJavaVersionInstalled

	If strComputer = "" Then strComputer = "."
	' ������ ��������� ������������� Java
	sJavaMajorVersionInstalled = sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE" & sRegPathModifier & "\JavaSoft\Java Runtime Environment", "CurrentVersion")
	Call WriteLog("sJavaVersionInstalledGetA, sJavaMajorVersionInstalled = " & sJavaMajorVersionInstalled,3)

	If sJavaMajorVersionInstalled = "1.8" then ' JAVA 1.8 code by https://github.com/airosa-id
		Dim sJavaMicroVersion, sJavaBrowserVersionInstalled
		sJavaMicroVersion = sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE" & sRegPathModifier & "\JavaSoft\Java Runtime Environment\"& sJavaMajorVersionInstalled, "MicroVersion")
		Call WriteLog("sJavaVersionInstalledGetA, sJavaMicroVersion = " & sJavaMicroVersion,3)
		sJavaBrowserVersionInstalled  = mid(sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE" & sRegPathModifier & "\JavaSoft\Java Runtime Environment", "BrowserJavaVersion"),4,2)

		Call WriteLog("sJavaVersionInstalledGetA, modificator (sJavaBrowserVersionInstalled) = " & sJavaBrowserVersionInstalled, 3)
		sJavaVersionInstalled = sJavaMajorVersionInstalled & "." & sJavaMicroVersion & "_" &  sJavaBrowserVersionInstalled
	End If

	If sJavaMajorVersionInstalled = "1.7"  then
		sJavaMajorVersionInstalled = Right(sJavaMajorVersionInstalled,1)
		' ������ ������ ����� ������ ������������� Java
		sJavaVersionInstalled  = sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE" & sRegPathModifier & "\JavaSoft\Java Runtime Environment", "Java" & sJavaMajorVersionInstalled & "FamilyVersion")
	End If

	If sJavaVersionInstalled = "" then
		Call WriteLog("sJavaVersionInstalledGetA, INFO: Can't get the installed version of Java from the registry. Probably Java on this computer is not installed",3)
		'bNeedToSendLog = True
		'Exit Function
	End If
	Call WriteLog("sJavaVersionInstalledGetA (" & sRegPathModifier & "), JavaVersionInstalled = " & sJavaVersionInstalled,2)
	sJavaVersionInstalledGetA = sJavaVersionInstalled
End Function ' sJavaVersionInstalledGetA

Function sJavaVersionWEBGet
	' ���������� ������ � ������� ���������� ������ JAVA, ���� � �����
	If IsDebug <> 1 then On Error Resume Next
	Dim  sJavaVersionWEB
'	������������������ ���� ���� - ��� ������� � �������� �������� ����������� ������ ������ ��� �� ������. � ��������� � ������ 1.8.0_51 �� ����� �� �����������...
'	sJavaVersionWEB = sHttpGet (csJavaVersionCurrentLnk)
'	sJavaVersionWEB = Replace(sJavaVersionWEB, vbCrLf, "") ' ������� ������ ������� ������
'	If sJavaVersionWEB = "" then ' �� ������ �������� ���������� ������ Java �� ������
'		Call WriteLog("sJavaVersionWEBGet, ERROR! Can't get current version of Java by the link " & csJavaVersionCurrentLnk,1)
'		bNeedToSendLog = True
'		Exit Function
'	End If

'	��� ��� ���� ��������� ������ Java �� ������ �� ����������� (������ �� ������ 1.8.0_51), ������ ����� ������� � ����� ���������� �������� ����������� ������ ���.
'	�� ��� ����� �� ������� ������ ���������� �������� ��������� ������ Java.. ���.
	Dim  MajorVersion
	Dim  MinorVersion
	Dim  sJavaLocalVersion
	Dim  objShell, objRegExp, objMatches

	Set objShell = WScript.CreateObject("WScript.Shell")
	MajorVersion = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion") 
'	� ���� �� Java?	
	If MajorVersion = "" then
		Call WriteLog("sJavaVersionWEBGet, ERROR! can't determine the local version of Java ",1)
		bNeedToSendLog = True
		Set objShell = nothing
		Set objRegExp = nothing
		Set objMatches = nothing
		Exit Function
	End If	 
	MinorVersion = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\" & MajorVersion & "\MicroVersion")
	sJavaLocalVersion = MajorVersion & "." & MinorVersion

'	����� �������� ������� ������ � ������������ ���. ��� ��� ���� ������� ���������� ������� ������� ���������� Java - ����� ������� ��� ��� �������.
	sJavaVersionWEB = sHttpGet ("https://javadl-esd-secure.oracle.com/update/" & sJavaLocalVersion & "/map-" & sJavaLocalVersion & ".xml")
	If sJavaVersionWEB = "" then ' �� ������ �������� ���������� ������ Java �� ������
		Call WriteLog("sJavaVersionWEBGet, ERROR! Can't get current version of Java by the link " & csJavaVersionCurrentLnk,1)
		bNeedToSendLog = True
		Set objShell = nothing
		Set objRegExp = nothing
		Set objMatches = nothing
		Exit Function
	End If
	Set objRegExp = CreateObject("VBScript.RegExp")
'	������ ��� �� ������� �� ������� ������
	objRegExp.Pattern= "<url>(.*[0-9]+\.[0-9]+\.[0-9]+_[0-9]+)-b"
	objRegExp.Global = True
	Set objMatches = objRegExp.Execute(sJavaVersionWEB)
	If objMatches.Count = 0 then
		Call WriteLog("sJavaVersionWEBGet, ERROR! can't determine the version of Java from parsed data received by the link ",1)
		bNeedToSendLog = True
		Set objShell = nothing
		Set objRegExp = nothing
		Set objMatches = nothing
		Exit Function
	Else
		sJavaVersionWEB = left(mid(objMatches.Item(0).Value,70,objMatches.Item(0).Length),objMatches.Item(0).Length-71)
	End If

	sJavaVersionWEBGet = sJavaVersionWEB
	Call WriteLog("sJavaVersionWEBGet, JavaVersionWEB   = " & sJavaVersionWEB,2)
	Set objShell = nothing
	Set objRegExp = nothing
	Set objMatches = nothing
End Function ' sJavaVersionWEBGet

Function sJavaVersionLocalGet
	' ���������� ������ � ������� ���������� ������ JAVA, ���� � ���������� �����������
	If IsDebug <> 1 then On Error Resume Next
	Dim  sJavaVersionLocal
	
	sJavaVersionLocal = sFileRead (sInstallerPath & csJavaInstallerVers)
	If sJavaVersionLocal = "" then
		Call WriteLog("sJavaVersionLocalGet, ERROR! Can't get current version of Java from file " & sInstallerPath & csJavaInstallerVers,1)
		bNeedToSendLog = True
		Exit Function
	End If
	sJavaVersionLocalGet = sJavaVersionLocal
	Call WriteLog("sJavaVersionLocalGet, JavaVersionLocal   = " & sJavaVersionLocal,2)
End Function ' sJavaVersionLocalGet

Function sFlashVersionInstalledGet (strComputer, sFlashType)
	' ���������� ������ � ������� ������ Flash, ������������� �� ���������� strComputer, ��� ��������� ���������� (���� �����)
	' sFlashType - ��� ����-�������������, A=ActiveX, P=Plugin
	If IsDebug <> 1 then On Error Resume Next
	
	Dim sTemp
	If strComputer = "" Then strComputer = "."
	Select Case sFlashType
		Case "A"
			sTemp = sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE\Macromedia\FlashPlayerActiveX", "Version")
		Case "P"
			sTemp = sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE\Macromedia\FlashPlayerPlugin", "Version")
	End Select
	
	If sTemp = "" then
		Call WriteLog("sFlashVersionInstalledGet, INFO: Can't get the installed version of Flash from the registry. " & _
					"Probably Flash type [" & sFlashType & "on this computer is not installed",3)
		'bNeedToSendLog = True
	Else
		Call WriteLog("sFlashVersionInstalledGet, FlashVersionInstalled type [" & sFlashType & "] = " & sTemp,2)
		sFlashVersionInstalledGet = sTemp
	End If
End Function ' sFlashVersionInstalledGet

Function sFlashVersionWEBGet (sFlashType)
	' ���������� ������ � ������� ���������� ������ Flash, ���� � �����
	' sFlashType - ��� ����-�������������, A=ActiveX, P=Plugin	
	If IsDebug <> 1 then On Error Resume Next
	Dim sTemp
	Dim sHttpText				' ���������� ����������� ��������
	Dim objRegExp				' ��� ��������� ������ ���������� ������
	Dim objMatches			 	' ��������������� ������ ��� �������� html
	
	sHttpText = sHttpGet (csFlashVersionCurrentLnk)
	If sHttpText = "" then
		Call WriteLog("sFlashVersionWEBGet, ERROR! Can't get HTML by the link  " & csFlashVersionCurrentLnk,1)
		bNeedToSendLog = True
		Exit Function
	End If
	' �������� ��������� HTML,���� ������ ������� � ������������������� ����, ����������� �������.
	Set objRegExp = CreateObject("VBScript.RegExp")
	objRegExp.Pattern= "<td>([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)</td>"
	objRegExp.Global = True
	Set objMatches = objRegExp.Execute(sHttpText)
	If objMatches.Count = 0 then
		Call WriteLog("sFlashVersionWEBGet, ERROR! Can't find Flash version number from the link " & csFlashVersionCurrentLnk, 1)
		If Err.Number <> 0 then Call WriteLogError("sFlashVersionWEBGet", Err.Number, Err.Description, Err.Source)
		bNeedToSendLog = True
		Exit Function
	Else
		Select Case sFlashType
			Case "A"
				'Internet Explorer (and other browsers that support Internet Explorer ActiveX controls and plug-ins)
				'Internet Explorer - ActiveX
				sTemp = Mid(objMatches.Item(0).Value ,5,len(objMatches.Item(0).Value )-9) 
			Case "P"
				'Firefox, Mozilla, Netscape, Opera (and other plugin-based browsers)
				'Firefox, Mozilla - NPAPI
				sTemp = Mid(objMatches.Item(2).Value ,5,len(objMatches.Item(2).Value )-9) 	
			End Select
			' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
			' !!! �������� ����� Item(x) ��� ��������� �� ����� Adobe �� ������ csFlashVersionCurrentLnk !!!
			' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	End If
	
	If sTemp = "" then
		Call WriteLog("sFlashVersionWEBGet, ERROR! Can't get current version of Flash type [" & sFlashType & "], from link " & csFlashVersionCurrentLnk,1)
		bNeedToSendLog = True
	Else
		Call WriteLog("sFlashVersionWEBGet, INFO: Flash type [" & sFlashType & "] = " & sTemp,3)
		sFlashVersionWEBGet = sTemp
	End If
End Function ' sFlashVersionWEBGet

Function sFlashVersionLocalGet (sFlashType)
	' ���������� ������ � ������� ���������� ������ Flash, ���� � ���������� �����������
	If IsDebug <> 1 then On Error Resume Next
	Dim  sTemp, sInstallerVersPath

	Select Case sFlashType
		Case "A"
			sInstallerVersPath = sInstallerPath & csFlashAInstallerVers 
		Case "P"
			sInstallerVersPath = sInstallerPath & csFlashPInstallerVers
	End Select

	sTemp =  sFileRead (sInstallerVersPath)
	If sTemp = "" then
		Call WriteLog("sFlashVersionLocalGet, ERROR! Can't get current version of Flash type [" & sFlashType & "], by the link "  & sInstallerVersPath,1)
		bNeedToSendLog = True
		Exit Function
	End If
	sFlashVersionLocalGet = sTemp
	Call WriteLog("sFlashVersionLocalGet, FlashVersionLocal type [" & sFlashType & "] = " & sTemp,3)
End Function ' sFlashVersionLocalGet

Function sFlashInstallerLinkGet (sFlashType, sFlashVersion)
	' ���������� ������-������ ��� ���������� ������� ���������������� ���� � ����� Adobe
	' ��������� � ������ 17.0.0.188
	' fpdownload.macromedia.com/pub/flashplayer/pdc/17.0.0.188/install_flash_player_ppapi.exe = Opera_new&Chromium (PPAPI)
	' fpdownload.macromedia.com/pub/flashplayer/pdc/17.0.0.188/install_flash_player_ax.exe = IE (ActiveX)
	' fpdownload.macromedia.com/pub/flashplayer/pdc/17.0.0.188/install_flash_player.exe = Firefox&Opera_old (NPAPI)
	' Const csFlashInstallerLink = "http://fpdownload.macromedia.com/pub/flashplayer/pdc/"
	If IsDebug <> 1 then On Error Resume Next
	Dim  sTemp

	sTemp = csFlashInstallerLink & sFlashVersion
	Select Case sFlashType
		Case "A"
			sTemp = sTemp & "/install_flash_player_ax.exe"
		Case "P"
			sTemp = sTemp & "/install_flash_player.exe"
	End Select

	sFlashInstallerLinkGet = sTemp
	Call WriteLog("sFlashInstallerLinkGet, FlashInstallerLink, type [" & sFlashType & "] = " & sTemp,3)
End Function ' sFlashVersionLocalGet

' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' ��������� ��������� � ������� - ������ � ������, �������, http, �������� ����� � �.�.

Function bParseCommandLine
	' ��������� ��������� ��������� ������, ������������� �������� �� ��������� ��� ���������� ����������								
	' ���������� false ��� ��������� ������
	If IsDebug <> 1 then On Error Resume Next

	glbMail = False
	glbLogFile = 0
	glbDebug = 1
	glbWEBMode = True
	glbWEBModeSaveInstall = false

	glbIgnoreJava   = False
	glbIgnoreJavaWoW6432 = False
	glbIgnoreFlashA = False
	glbIgnoreFlashP = False
	glbUninstallJavaOld = False
	
	bParseCommandLine = true
	dim objNamed
	Set objNamed = WScript.Arguments.Named

	' must be first
	If objNamed.Exists("debug") Then 
		select case lcase(objNamed.Item("debug"))
			case "0","1","2","3"
				glbDebug = cint(objNamed.Item("debug"))
				Call WriteLog ("bParseCommandLine, glbDebug = " & glbDebug,3)
			case else
				Call WriteLog ("bParseCommandLine, glbDebug ERROR",1)
		end select
	End If

	If objNamed.Exists("mail") Then 
		select case lcase(objNamed.Item("mail"))
			case "+", "1", "true"
				glbMail = true
				Call WriteLog ("bParseCommandLine, glbMail = true",3)
			case "-", "0", "false"
				glbMail = false
				Call WriteLog ("bParseCommandLine, glbMail = false",3)
			case else
				Call WriteLog ("bParseCommandLine, glbMail ERROR",1)
		end select
	End If

	If objNamed.Exists("logfile") Then
		select case lcase(objNamed.Item("logfile"))
			case "+", "1", "true"
				glbLogFile = 1
				Call WriteLog ("bParseCommandLine, glbLogFile = 1",3)
			case "-", "0", "false"
				glbLogFile = 0
				Call WriteLog ("bParseCommandLine, glbLogFile = 0",3)
			case "2"
				glbLogFile = 2
				Call WriteLog ("bParseCommandLine, glbLogFile = 2",3)
			case else
				Call WriteLog ("bParseCommandLine, glbLogFile ERROR",1)
		end select
	End If

	If objNamed.Exists("webmode") Then 
		select case lcase(objNamed.Item("webmode"))
			case "+", "1", "true"
				glbWEBMode = true
				Call WriteLog ("bParseCommandLine, glbWEBMode = true",3)
			case "-", "0", "false"
				glbWEBMode = false
				Call WriteLog ("bParseCommandLine, glbWEBMode = false",3)
			case else
				Call WriteLog ("bParseCommandLine, glbWEBMode ERROR",1)
		end select
	End If

	If objNamed.Exists("webmodesaveinstall") Then 
		select case lcase(objNamed.Item("webmodesaveinstall"))
			case "+", "1", "true"
				glbWEBModeSaveInstall = true
				Call WriteLog ("bParseCommandLine, glbWEBModeSaveInstall = true",3)
			case "-", "0", "false"
				glbWEBModeSaveInstall = false
				Call WriteLog ("bParseCommandLine, glbWEBModeSaveInstall = false",3)
			case else
				Call WriteLog ("bParseCommandLine, glbWEBModeSaveInstall ERROR",1)
		end select
	End If
	
	If objNamed.Exists("webmodesaveinstallforce") Then 
		Call WriteLog ("bParseCommandLine, WEBModeSaveInstallForce = true",3)
		' ������ ���������� ��� ����������
		sInstallerPath = csInstallerPath		' ���� ���������� ����������� ��� /WEBModeSaveInstall+
		' ��� ������������� ��������� ���� � ����� �������� ��������
		If Right(sInstallerPath,1) <> "\" Then 
			sInstallerPath = sInstallerPath + "\"
		End If

		'java
		Call WriteLog("---------------------------------------------------------------------------------------------------",2)
		Call WriteLog("JAVA - Compare downloaded and current Java version",2)
		if sJavaVersionWEBGet() <>  sJavaVersionLocalGet() then
			Call WriteLog("WARNING! Downloaded and current Java version is NOT identical",1)
			bNeedToSendLog = True
			Call WriteLog ("Java 32bit - START ---",2)
			Call HttpGetSave(sJavaGetLinkToDownload(""), sInstallerPath & csJavaInstaller)
			Call WriteLog ("Java 64bit - START ---",2)
			Call HttpGetSave(sJavaGetLinkToDownload("64"), sInstallerPath & csJavaInstaller64)

			sJavaVersionCurrent = sJavaVersionWEBGet
			Call myRun("cmd /c echo " & sJavaVersionCurrent & "> " & sInstallerPath & csJavaInstallerVers)
			sJavaNativeUpdateStatus   = "UPDATED"
			sJavaWoW6432UpdateStatus  = "UPDATED"
			Call WriteLog ("Java - FINISH  ===",2)
		End If
		Call WriteLog("JAVA - OK",2)
		Call WriteLog("===================================================================================================",2)
		
		'flash ActicveX
		Call WriteLog("---------------------------------------------------------------------------------------------------",2)
		Call WriteLog("FlashActiveX - Compare downloaded and current FlashActiveX version",2)
		if sFlashVersionWEBGet("A") <> sFlashVersionLocalGet("A") then
			Call WriteLog("WARNING! Downloaded and current FlashActiveX version is NOT identical",1)
			bNeedToSendLog = True
			Call WriteLog ("FlashA - START ---",2)
			sFlashAVersionCurrent = sFlashVersionWEBGet("A")
			Call HttpGetSave(sFlashInstallerLinkGet("A",sFlashAVersionCurrent), sInstallerPath & csFlashAInstaller)
			sFlashAVersionCurrent = sFlashVersionWEBGet("A")
			Call myRun("cmd /c echo " & sFlashAVersionCurrent & "> " & sInstallerPath & csFlashAInstallerVers)
			sFlashAUpdateStatus       = "UPDATED"
			Call WriteLog ("FlashA - FINISH ===",2)
		End If
		Call WriteLog("FlashActiveX - OK",2)
		Call WriteLog("===================================================================================================",2)
		'flash Plugin
		Call WriteLog("---------------------------------------------------------------------------------------------------",2)
		Call WriteLog("FlashPlugin - Compare downloaded and current FlashPlugin version",2)
		if sFlashVersionWEBGet("P") <> sFlashVersionLocalGet("P") then
			Call WriteLog("WARNING! Downloaded and current FlashPlugin version is NOT identical",1)
			bNeedToSendLog = True
			Call WriteLog("---------------------------------------------------------------------------------------------------",2)
			Call WriteLog ("FlashP - START ---",2)
			sFlashPVersionCurrent = sFlashVersionWEBGet("P")
			Call HttpGetSave(sFlashInstallerLinkGet("P",sFlashPVersionCurrent), sInstallerPath & csFlashPInstaller)
			sFlashPVersionCurrent = sFlashVersionWEBGet("P")
			Call myRun("cmd /c echo " & sFlashPVersionCurrent & "> " & sInstallerPath & csFlashPInstallerVers)
			sFlashPUpdateStatus       = "UPDATED"
			Call WriteLog ("FlashP - FINISH ===",2)
		End If
		Call WriteLog("FlashPlugin - OK",2)
		Call WriteLog("===================================================================================================",2)

		' ��� ������������� ���������� �����
		Call SendLog()

		' ��� ������������� ���������� ����� � ���-����
		Call WriteLogToFile()

		Call WriteLog("===================================================================================================",2)
		WScript.Quit
	End If

	If objNamed.Exists("showversion") Then 
		dim strComputer
		If objNamed.Item("showversion") <> "" then
			strComputer = objNamed.Item("showversion")
		else
			strComputer = "."
		End If
		Call WriteLog ("bParseCommandLine, ShowVersion for comp " & strComputer,3)
		
		If strComputer <> "." Then
			' ������ �������� ��� ���� �������
			If IsPingSucsess (strComputer) Then
				WScript.Echo ("��������� " & strComputer & " ��������!")
			Else
				WScript.Echo ("��������� " & strComputer & " �� ��������!")
				WScript.Quit
			End If
		End If
		' TODO - ���������� ����� ������
		If Is64BitSystem (strComputer) = true then
		   ' ������ ��������
			call WriteLog("JAVA64 = " & sJavaVersionInstalledGet(strComputer) ,1)
			' ��� 64 ������ ������ ������������� ���������� ������ 32� ���
			call WriteLog("JAVA32 = " & sJavaWoW6432VersionInstalledGet(strComputer) ,1)
		else
			call WriteLog("JAVA   = " & sJavaVersionInstalledGet(strComputer) ,1)
		End If

		call WriteLog("FlashA = " & sFlashVersionInstalledGet(strComputer, "A") ,1)
		call WriteLog("FlashP = " & sFlashVersionInstalledGet(strComputer, "P") ,1)
		WScript.Quit
	End If

	If objNamed.Exists("mailtest") Then 	
		Call mySendMail("FJUpdater test mail from " & sEnvGet("%COMPUTERNAME%"), "Test")
		Call WriteLog ("bParseCommandLine, mailtest",3)
		WScript.Quit
	End If

	If objNamed.Exists("IgnoreJava") Then
		glbIgnoreJava = True
		Call WriteLog ("bParseCommandLine, IgnoreJava",3)
	End If

	If objNamed.Exists("IgnoreJavaWoW6432") Then
		glbIgnoreJavaWoW6432 = True
		Call WriteLog ("bParseCommandLine, IgnoreJavaWoW6432",3)
	End If

	If objNamed.Exists("IgnoreFlashA") Then
		glbIgnoreFlashA = True
		Call WriteLog ("bParseCommandLine, IgnoreFlashA",3)
	End If

	If objNamed.Exists("IgnoreFlashP") Then
		glbIgnoreFlashP = True
		Call WriteLog ("bParseCommandLine, IgnoreFlashP",3)
	End If
	If objNamed.Exists("UninstallJavaOld") Then
		If objNamed.Item("UninstallJavaOld") <> "" then
			dim sComputer
			sComputer = objNamed.Item("UninstallJavaOld")
			Call WriteLog ("bParseCommandLine, UninstallJavaOld for comp " & sComputer,3)
			If sComputer <> "." Then
				' ������ �������� ��� ���� �������
				If IsPingSucsess (sComputer) Then
					WScript.Echo ("��������� " & sComputer & " ��������!")
				Else
					WScript.Echo ("��������� " & sComputer & " �� ��������!")
					WScript.Quit
				End If
			End If
			Call UninstallJavaOld(sComputer)
			WScript.Quit
		else
			glbUninstallJavaOld = True
			Call WriteLog ("bParseCommandLine, UninstallJavaOld",3)
		End If
	End If
	If (objNamed.Exists("?")) or (objNamed.Exists("help"))  Then 	
		Call PrintHelp()
		WScript.Quit
	End If

	If Err.Number <> 0 then 
		Call WriteLogError("bParseCommandLine", Err.Number, Err.Description, Err.Source)
		bParseCommandLine = false
		'todo : ����� Help'�
	End If
End Function

sub PrintHelp
	' ������� �������
	If IsDebug <> 1 then On Error Resume Next
	Call WriteLog("PrintHelp", 0)	
	Dim objFSO, fIn, sTemp, sTemp2

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	' TODO = ������� ������� WScript.ScriptName � ������ wsf �� vbs (��� txt)
	' �� ������� ����� ������� �������� ��� �������, �������� ���� ������� �������
	Set fIn = objFSO.OpenTextFile(left (WScript.ScriptFullName, Len(WScript.ScriptFullName)-Len(WScript.ScriptName)) & "FJUpdater_main.vbs", 1)
	
	Do While (Not fIn.AtEndOfStream) 
		sTemp = fIn.ReadLine
		If sTemp = "'[HELP_END]" then wscript.quit
		If sTemp <> "" then Wscript.Echo (Mid(sTemp,2,Len(sTemp)-1))
	Loop
	
end Sub

Sub mySendMail (sSubject, sTextBody)
	' ���������� ����� ����� CDO
	If IsDebug <> 1 then On Error Resume Next
	
	If not IsPingSucsess (csSMTPServer) Then
		Call WriteLog("mySendMail : WARNING! can't ping SMTP server!", 1)
	End If
	
	dim objMessage

	Const cdoURL = "http://schemas.microsoft.com/cdo/configuration/"
	
	Call WriteLog("mySendMail, sSubject=" & sSubject,3)
	Set objMessage = CreateObject("CDO.Message") 
	With objMessage
		.Subject = sSubject 
		.From = csFrom
		.To = csTo
		.TextBody = sTextBody
		'��������
		.fields.Item("urn:schemas:mailheader:importance").Value = "high"	
		'��������� 
		.fields.Item("urn:schemas:mailheader:priority").Value = 1
		.fields.Update()
	End with

	With objMessage.Configuration.Fields
		'Send Using method (1 = local smtp, 2 = remote network smtp, 3 = MS Exchange)
		.Item(cdoURL & "sendusing") = ciSendUsing
		'Name or IP of Remote SMTP Server
		.Item(cdoURL & "smtpserver") = csSMTPServer
		'Type of authentication, NONE, Basic (Base64 encoded), NTLM
		.Item(cdoURL & "smtpauthenticate") = ciSMTPAuthenticate
		'Your UserID on the SMTP server
		.Item(cdoURL & "sendusername") = csSendUserName
		'Your password on the SMTP server
		.Item(cdoURL & "sendpassword") = csSendPassword
		'Server port (typically 25)
		.Item(cdoURL & "smtpserverport") = ciSMTPServerPort 
		'Use SSL for the connection (False or True)
		.Item(cdoURL & "smtpusessl") = cbSMTPUseSSL
		'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
		.Item(cdoURL & "smtpconnectiontimeout") = 10
		'Update the config.
		.Update
	end with
	objMessage.Send 
	Set objMessage = nothing
End Sub ' mySendMail

function IsHostCscript()
	' Determines which program is used to run this script.
	' Returns true If the script host is cscript.exe
	' from WINDOWS\system32\prncnfg.vbs
	' Copyright (c) Microsoft Corporation. All rights reserved.
    If IsDebug <> 1 then On Error Resume Next
    dim strFullName
    dim strCommand
    dim i, j
    dim bReturn
    bReturn = False
    strFullName = WScript.FullName
    i = InStr(1, strFullName, ".exe", 1)
    If i <> 0 then
        j = InStrRev(strFullName, "\", i, 1)
        If j <> 0 then
            strCommand = Mid(strFullName, j+1, i-j-1)
            If LCase(strCommand) = "cscript" then
                bReturn = true
            End If
        End If
    End If
    If Err <> 0 then
        call wscript.echo("Error 0x" & hex(Err.Number) & " occurred. " & Err.Description _
                          & ". " & vbCRLF & "The scripting host could not be determined.")
    End If
    IsHostCscript = bReturn
end function

Sub WriteLog(sText, iDebug) 
	' ��������� ������ � ��� ��� �������� � ����� � �������, ���� iDebug <= glbDebug
	If IsDebug <> 1 then On Error Resume Next
	If iDebug <= glbDebug then
		Dim sTemp
		sTemp = FormatDateTime(Now,3) & " " & sText
		Wscript.Echo(sTemp)
		sLog = sLog & vbCRLF & sTemp
	End If
End Sub

Sub WriteLogError(sText, iNumber, sDescription, sSource)
	' ��������� ������ � ������� � ���
	If IsDebug <> 1 then On Error Resume Next
	Call WriteLog(sText & ", ERROR! : " & iNumber & ": " & sDescription & " generated by " & sSource, 2)
	If glbDebug >= 2 Then ' ���� ������� ����������� �������
		bNeedToSendLog = True ' �������� �� ������ ������
	End If
End Sub

function sRegRead (strComputer, regHive, sKeyPath, sKeyName)
	' ���������� �������� ����� ������� sKeyPath + sKeyName ���������� strComputer
	If IsDebug <> 1 then On Error Resume Next
	Call WriteLog("sRegRead, sComputer=" & strComputer & ", regHive=" & regHive & ", sKeyPath=" & sKeyPath & ", sKeyName=" & sKeyName, 3)	
	
	Dim objReg, sTemp

	If strComputer = "" Then strComputer = "."	
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	objReg.GetExpandedStringValue regHive, sKeyPath, sKeyName, sTemp
	If Err.Number <> 0 then Call WriteLogError("sRegRead", Err.Number, Err.Description, Err.Source)
	sRegRead = sTemp
End function


Function sFileRead (sFile)
	' ���������� ���������� �����
	If IsDebug <> 1 then On Error Resume Next
	Call WriteLog("sFileRead, sFile=" & sFile, 3)	
	
	
	Dim objFSO, fIn
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(sFile) then
		Set fIn = objFSO.OpenTextFile(sFile, 1)
		sFileRead = Replace(fIn.ReadAll, vbCrLf, "") ' ������ ���� ���� � ������� �������� ������
		fIn.Close
		Set fIn = nothing
	Else
		Call WriteLog("ERROR : sFileRead, file not found, sFile=" & sFile, 2)	
	End If
	
	Set objFSO = Nothing
End Function

Function sEnvGet (sEnv)
	' ���������� �������� ���������� ���������
	On Error Resume Next	
	
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	sEnvGet = objShell.ExpandEnvironmentStrings(sEnv)
	If Err.Number <> 0 then Call WriteLogError("sEnvGet", Err.Number, Err.Description, Err.Source)
End Function

Function myRun(sCommand)
	' ��������� sCommand � ��� ����������, ���������� ����� ��������
	If IsDebug <> 1 then On Error Resume Next
	Dim objShell, objShellExec
	
	Const ciShellExecStatusRun  = 0 '���������� ��������.
	Const ciShellExecStatusStop = 1 '1 - ���������� ���������.
	
	Set objShell = CreateObject("WScript.Shell")	
	Set objShellExec = objShell.Exec(sCommand)
	' ���, ���� ���������� ������
	While objShellExec.Status = ciShellExecStatusRun
		WScript.Sleep 100
	Wend
	myRun = objShellExec.StdOut.ReadAll
End Function

Function sHttpGet (sURL) 
	' ���������� ���������� ����������� �� ������ sURL
	If IsDebug <> 1 then On Error Resume Next
	Dim objHttp
	
	Call WriteLog("sHttpGet, sURL=" & sURL,3)
	Set objHttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")
	objHttp.Open "GET", sURL, False
	' ��������� IE10 64bit
	objHttp.setRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
	objHttp.Send
	If Err.Number <> 0 then Call WriteLogError("sHttpGet", Err.Number, Err.Description, Err.Source)
	sHttpGet = objHttp.ResponseText
	Set objHttp = nothing
End Function

Sub HttpGetSave(sURL, sFile) 
	' ��������� ���������� ����������� �� ������ sURL � ���� sFile
	If IsDebug <> 1 then On Error Resume Next
	Dim objHttp, objADOStream, objFSO
	
	Call WriteLog("HttpGetSave, sURL=" & sURL & ", sFile=" & sFile,3)
	Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")
	objHttp.Open "GET", sURL, False	
	' ��������� IE10 64bit
	objHttp.setRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
	objHttp.Send
	If objHttp.Status = 200 Then
		Set objADOStream = CreateObject("ADODB.Stream")
		objADOStream.Open
		objADOStream.Type = 1 		' adTypeBinary
		objADOStream.Write objHttp.ResponseBody
		objADOStream.Position = 0   ' Set the stream position to the start
		Set objFSO = Createobject("Scripting.FileSystemObject")
		If objFSO.Fileexists(sFile) Then objFSO.DeleteFile sFile
		objADOStream.SaveToFile sFile
		objADOStream.Close
		If Err.Number <> 0 then Call WriteLogError("HttpGetSave", Err.Number, Err.Description, Err.Source)
	End If
	Set objFSO = nothing
	Set objADOStream = nothing
	Set objHttp = nothing
End sub

Function Is64BitSystem (strComputer)
	' ���������� True, ���� ���������� strComputer 64������,  ���� strComputer = ����� �� �� ��������� ����������
	' http://stackoverflow.com/questions/3583604/how-can-i-use-vbscript-to-determine-whether-i-am-running-a-32-bit-or-64-bit-wind
	If IsDebug <> 1 then On Error Resume Next
	
	If strComputer = "" then strComputer = "." 
	If GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Processor='cpu0'").AddressWidth = 64 then
		Is64BitSystem = True
	else
		Is64BitSystem = False
	End If
End Function

Function IsPingSucsess (strComputer)
	' ���������� True, ���� ���������� strComputer �������� �� ����
	If IsDebug <> 1 then On Error Resume Next

	dim sTemp 
	sTemp = LCase(myRun ("%comspec% /c ping.exe -n 2 " & strComputer))
	If InStr(sTemp, "ttl=")  Then
		IsPingSucsess = True
	else
		IsPingSucsess  = False
	End If
End Function
