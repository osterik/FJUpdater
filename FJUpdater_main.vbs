'-----------------------------------------------------------------------------------------------------------------------
'
' Copyright (c) Ilya Kisleyko. All rights reserved.
' 
' AUTHOR  : Ilya Kisleyko
' E-MAIL  : osterik@gmail.com
' DATE    : 14.09.2014
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
'/debug:0 = �� �������� ������
'/debug:1 = ���������� ������ ��������� � ����������� ��������� (default) � ������
'/debug:2 = ���������� ��������� ��������� 
'/debug:3 = ���������� ��������� ��������� � ���������� ��������� � �� ���������
'
'/mail:1 = ���������� ����� ��� ������� ������ ��� ��� ������� ���������� 
'/mail:0 = ����� ����� (default)
'
'/WebMode:1 = ��������� ���������� �������� �� ��������� (default)
'/WebMode:0 = �� ��������� ����� (csInstallerPath)
'
'/WEBModeSaveInstall:1 = ��������� ���������� � ��������� ����� (csInstallerPath)
'/WEBModeSaveInstall:0 = ��������� ���������� � %TEMP% (default)
'���� /WEBModeSaveInstall ����� ����� ������ ��� /WEBMode:1
'
'/LogFile:0 = �� ��������� ������� � �� ���������� � ���� ������
'/LogFile:1 = ������� ������� (���� ����������� ��� ����������)
'
'/WEBModeSaveInstallForce = ������ �����, ������������� ������� ���������� ������ ������������
'� ��������� � ��������� ����� (csInstallerPath). �� ��������� ������������� ������ ��������.
'������������ �� ���������� �������������� ��� ������� ��������� ��������������� ������������ ���������� 
'
'/MailTest = ������ �����, �������� �������� �������� �����, ��������� �������� ���������
'
'/ShowVersion[:comp] = ������ �����, ������� ����� ����� ������������� �� ���������� �������� Java & Flash
' ��� ������� ��� ���������� - �� ��������� ����������, ���� ���� �������� - �� ���������� ��� ���\����� ����������
'
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
Const csJavaInstaller64		= "java_installer64.exe"
Const csJavaInstallerVers	= "java_current.txt"
Const csFlashPInstaller		= "flashP_installer.exe"	' Flash Plugin (Firefox, Mozilla, Netscape, Opera)
Const csFlashPInstallerVers	= "flashP_current.txt"
Const csFlashAInstaller		= "flashA_installer.exe"	' Flash ActiveX (IE)
Const csFlashAInstallerVers	= "flashA_current.txt"

'===================================================================================================
' ��� �������� ������ ���������� ������ ((��� /WEBMode+)
Const csJavaVersionCurrentLnk = "http://www.java.com/applet/JreCurrentVersion2.txt"
Const csFlashVersionCurrentLnk = "http://www.adobe.com/software/flash/about/"
' � ��� ������\������ �����������
Const csJavaInstallerLink   = "http://java.com/en/download/windows_manual.jsp"
Const csFlashPInstallerLink = "http://fpdownload.macromedia.com/pub/flashplayer/latest/help/install_flash_player.exe"
Const csFlashAInstallerLink = "http://fpdownload.macromedia.com/pub/flashplayer/latest/help/install_flash_player_ax.exe"
' ����� ������� ������������
Const csJavaInstallerParams = "/s"
Const csFlashInstallerParams = "/install"

'===================================================================================================
' ��� ������� � �������, ������� RegRead
Const HKEY_CLASSES_ROOT  	= &H80000000
Const HKEY_CURRENT_USER 	= &H80000001
Const HKEY_LOCAL_MACHINE 	= &H80000002
Const HKEY_USERS 		= &H80000003
Const HKEY_CURRENT_CONFIG  	= &H80000005

' ���������� ���������, ����� ���� �������������� �� ��������� ������
Dim glbDebug					' /debug0 = �� �������� ������
								' /debug1 = ���������� ������ ��������� � ����������� ��������� (default)
								' /debug2 = ���������� ��������� ��������� 
								' /debug3 = ���������� ��������� ��������� � ���������� ��������� � �� ���������
Dim glbMail 					' /mail+ = ���������� ����� ��� ������� ������ ��� ��� ������� ���������� 
								' /mail- = ����� ����� (default)
Dim glbWEBMode 					' /WEBMode+ = ��������� ���������� �������� �� ��������� (default)
								' /WEBMode- = �� ��������� ����� (csInstallerPath)
Dim glbWEBModeSaveInstall		' ����� ����� ������ ��� /WEBMode+
								' /WEBModeSaveInstall+ = ��������� ���������� � ��������� ����� (csInstallerPath)
								' /WEBModeSaveInstall- =  ��������� ���������� � %TEMP% (default)
Dim glbLogFile							' /LogFile+ = ��������� ����� � ���-����
								' /LogFile- = �� �������� ����� � ��� ����
'Dim glbComputer					' ���(��� �����) ����������, �� ������� ��������� ������ ��������
' ���������� ����������
Dim sInstallerPath							' ���� ���������� ����������� ��� /WEBModeSaveInstall+
Dim sLog 									' ����� ���, ���������� �� ����� ��� ������� ��� ��������� ����������
Dim bNeedToSendLog							' ���� ������������� �������� ���� �� �����
Dim sJavaNativeUpdateStatus, sJavaWoW6432UpdateStatus, sFlashAUpdateStatus, sFlashPUpdateStatus		' �������� ������ �������� ����������
Dim sJavaVersionCurrent, sFlashPVersionCurrent, sFlashAVersionCurrent ' ����� ������ ��� ��������� ����������, ������������ ��� ������ ����� � ������� ������

Sub Main
	If IsDebug <> 1 then On Error Resume Next
	
	' Abort If the host is not cscript
	If not IsHostCscript() then
		call wscript.echo("������������� ��������� ������ ������ � �������������� CScript.")
		'"������������� ��������� ������ ������ � �������������� CScript."  	& vbCRLF & vbCRLF &  
		'"��� �����:"														& vbCRLF & vbCRLF & _
		'"1. ����������� ""CScript " & WScript.ScriptName & " /���������"" "	& vbCRLF & vbCRLF & _
		'"���" 																& vbCRLF & vbCRLF & _
		'"2. ��������  WScript ��  CScript (�����"									& vbCRLF & _
		'"   ��� ����� ��������� ""CScript //H:CScript //S"" "				& vbCRLF & _
		'"   � ���������� ������ """ & WScript.ScriptName & " /���������""."	& vbCRLF
		wscript.quit
	End If
	
	If bParseCommandLine = false Then
		' �� ������ ��������� ��������� ��������� ������
		Call WriteLog("Main, ERROR! can't parse command line switches!", 1)
		Exit Sub 
	End If
	
	Call WriteLog(FormatDateTime(Now,2) & " JavaFlashUpdaterAdmin v2.6 (c) Ilya Kisleyko, osterik@gmail.com", 2)
	
	'���� ���������� �����������...
	If glbWEBMode then 
		If glbWEBModeSaveInstall Then
			sInstallerPath = csInstallerPath		' ...��� /WEBModeSaveInstall+
		Else
			sInstallerPath = sEnvGet("%temp%")		' ...��� /WEBModeSaveInstall-
		End If
	else
		sInstallerPath = csInstallerPath			' ..��� ��������� ������
	End If
	
	' ��� ������������� ��������� ���� � ����� �������� ��������
	If Right(sInstallerPath,1) <> "\" Then 
		sInstallerPath = sInstallerPath + "\"
	End If
	' ���� ��� ������������� ���������� ���� �� �����
	bNeedToSendLog = false
	
	' ������ ��������
	sJavaNativeUpdateStatus   = "UNKNOWN"
	sJavaWoW6432UpdateStatus  = "UNKNOWN"
	sFlashAUpdateStatus       = "UNKNOWN"
	sFlashPUpdateStatus       = "UNKNOWN"
	
	Call WriteLog("---------------------------------------------------------------------------------------------------",2)
	Call WriteLog("JavaNative - START",2)
	' JAVA Native (32 on 32, 64 on 64) - ������� ������ ���������� � �������������
	' ��� ��������� �� ���������� ������ ...
	If glbWEBMode Then ' WWW
		sJavaVersionCurrent = sJavaVersionWEBGet
	Else               ' ��������� ���������
		sJavaVersionCurrent = sJavaVersionLocalGet	
	End If
	If sJavaVersionCurrent <> "" then 
		If  sJavaVersionInstalledGet(".") <> "" then ' �� ���������� ������������ �����-�� ������
			If sJavaVersionInstalledGet(".") <> sJavaVersionCurrent then
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
				' �������������
				If Is64BitSystem (".") = true then
					Call myRun(sInstallerPath & csJavaInstaller64 & " " & csJavaInstallerParams)
				else
					Call myRun(sInstallerPath & csJavaInstaller & " " & csJavaInstallerParams)
				End If
			
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
	
	Call WriteLog("---------------------------------------------------------------------------------------------------",2)
	Call WriteLog("JavaWoW6432 - START",2)
	' JAVA 32bit on 64bit system) - ������� ������ ���������� � �������������
	' ��� ��������� �� ���������� ������ ...
	If glbWEBMode Then ' WWW
		sJavaVersionCurrent = sJavaVersionWEBGet
	Else               ' ��������� ���������
		sJavaVersionCurrent = sJavaVersionLocalGet	
	End If
	If sJavaVersionCurrent <> "" then 
		If  sJavaWoW6432VersionInstalledGet(".") <> "" then ' �� ���������� ������������ �����-�� ������
			If sJavaWoW6432VersionInstalledGet(".") <> sJavaVersionCurrent then
				' �� ���������, ���������
				Call WriteLog("WARNING! Installed and current WoW6432(32bit Java on 64bit OS) Java version is NOT identical",1)
				' ���� ����� ��������� �����
				bNeedToSendLog = True
				If glbWEBMode Then
					' �������� ������ �� ������ � ��������� ����������
					Call HttpGetSave(sJavaGetLinkToDownload("32"), sInstallerPath & csJavaInstaller) 
				End If
				' �������������
				Call myRun(sInstallerPath & csJavaInstaller & " " & csJavaInstallerParams)
				
				' �������� ��������� ���������
				If sJavaWoW6432VersionInstalledGet(".") = sJavaVersionCurrent then
					sJavaWoW6432UpdateStatus = "SUCCESS"
					If glbWEBModeSaveInstall then
						' ���������� ���� � ������� ���������� ������ �����������
						Call myRun("cmd /c echo " & sJavaVersionCurrent & "> " & sInstallerPath & csJavaInstallerVers)
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
					Call HttpGetSave(csFlashAInstallerLink, sInstallerPath & csFlashAInstaller)
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
					Call HttpGetSave(csFlashPInstallerLink, sInstallerPath & csFlashPInstaller)
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

	' ��� ������������� ���������� �����
	If bNeedToSendLog and glbMail then
		Call WriteLog("Sending log to e-mail",2)
		Call mySendMail("Updater Java=[" & sJavaNativeUpdateStatus &_
			"], FlashA =[" & sFlashAUpdateStatus & "], FlashP =[" & sFlashPUpdateStatus &_
			"] on " & sEnvGet("%COMPUTERNAME%"), sLog)
	End If

	' ���������� ����� � ��� ����
	If glbLogFile then 
		Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(ReportFilePath & "Update_flash_java.log",2,true)
		objFSO.WriteLine(sLog)
		objFSO.Close
		Set objFSO = Nothing
	End If

	Call WriteLog("===================================================================================================",2)
End Sub ' Main

Function sJavaGetLinkToDownload (sBits)
	' ������ ��������� csJavaInstallerLink
	' � ���������� ������ �� ������ ���������� ������ Java
	' ���� sBits = 64, �� ���������� ������ �� 64-������ ������, ����� = 32�
	If IsDebug <> 1 then On Error Resume Next
	Dim sHttpText				' ���������� ��������� ��������
	Dim objRegExp				' ��� ��������� ������ �� ����������
	Dim objMatches, objMatch	' ��������������� ������� ��� �������� html
	
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
	
	Dim sJavaMajorVersionInstalled, sJavaMicroVersion, sJavaBrowserVersionInstalled, sJavaVersionInstalled
	If strComputer = "" Then strComputer = "."
	' ������ ��������� ������������� Java
	sJavaMajorVersionInstalled = sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE" & sRegPathModifier & "\JavaSoft\Java Runtime Environment", "CurrentVersion")
	If sJavaMajorVersionInstalled = "1.8" then 
		sJavaMicroVersion = sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE" & sRegPathModifier & "\JavaSoft\Java Runtime Environment\"& sJavaMajorVersionInstalled, "MicroVersion")	
		sJavaBrowserVersionInstalled  = mid(sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE" & sRegPathModifier & "\JavaSoft\Java Runtime Environment", "BrowserJavaVersion"),4,2)
		sJavaVersionInstalled = sJavaMajorVersionInstalled & "." & sJavaMicroVersion & "_" &  sJavaBrowserVersionInstalled
	else
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
	
	sJavaVersionWEB = sHttpGet (csJavaVersionCurrentLnk)
	sJavaVersionWEB = Replace(sJavaVersionWEB, vbCrLf, "") ' ������� ������ ������� ������
	If sJavaVersionWEB = "" then ' �� ������ �������� ���������� ������ Java �� ������
		Call WriteLog("sJavaVersionWEBGet, ERROR! Can't get current version of Java by the link " & csJavaVersionCurrentLnk,1)
		bNeedToSendLog = True
		Exit Function
	End If
	sJavaVersionWEBGet = sJavaVersionWEB
	Call WriteLog("sJavaVersionWEBGet, JavaVersionWEB   = " & sJavaVersionWEB,2)
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
				sTemp = Mid(objMatches.Item(0).Value ,5,len(objMatches.Item(0).Value )-9) 
			Case "P"
				'Firefox, Mozilla, Netscape, Opera (and other plugin-based browsers)
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


' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' ��������� ��������� � ������� - ������ � ������, �������, http, �������� ����� � �.�.

Function bParseCommandLine
	' ��������� ��������� ��������� ������, ������������� �������� �� ��������� ��� ���������� ����������								
	' ���������� false ��� ��������� ������
	If IsDebug <> 1 then On Error Resume Next
	glbMail = False
	glbDebug = 1
	glbWEBMode = True
	glbWEBModeSaveInstall = false
	glbLogFile = False
	bParseCommandLine = true
	dim objNamed
	Set objNamed = WScript.Arguments.Named

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

	If objNamed.Exists("LogFile") Then 
		select case lcase(objNamed.Item("LogFile"))
			case "+", "1", "true"
				glbLogFile = true
				Call WriteLog ("bParseCommandLine, glbLogFile = true",3)
			case "-", "0", "false"
				glbLogFile = false
				Call WriteLog ("bParseCommandLine, glbLogFile = false",3)
			case else
				Call WriteLog ("bParseCommandLine, glbLogFile ERROR",1)
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
		Call WriteLog ("Java 32bit - START",2)		
 		Call HttpGetSave(sJavaGetLinkToDownload(""), sInstallerPath & csJavaInstaller) 
		Call WriteLog("---------------------------------------------------------------------------------------------------",2)
		Call WriteLog ("Java 64bit - START",2)		
		Call HttpGetSave(sJavaGetLinkToDownload("64"), sInstallerPath & csJavaInstaller64) 
		
		sJavaVersionCurrent = sJavaVersionWEBGet
		Call myRun("cmd /c echo " & sJavaVersionCurrent & "> " & sInstallerPath & csJavaInstallerVers)
		'flash ActicveX
		Call WriteLog("---------------------------------------------------------------------------------------------------",2)
		Call WriteLog ("FlashA - START",2)		
		Call HttpGetSave(csFlashAInstallerLink, sInstallerPath & csFlashAInstaller)
		sFlashAVersionCurrent = sFlashVersionWEBGet("A")
		Call myRun("cmd /c echo " & sFlashAVersionCurrent & "> " & sInstallerPath & csFlashAInstallerVers)
		'flash Plugin
		Call WriteLog("---------------------------------------------------------------------------------------------------",2)
		Call WriteLog ("FlashP - START",2)		
		Call HttpGetSave(csFlashPInstallerLink, sInstallerPath & csFlashPInstaller)
		sFlashPVersionCurrent = sFlashVersionWEBGet("P")
		Call myRun("cmd /c echo " & sFlashPVersionCurrent & "> " & sInstallerPath & csFlashPInstallerVers)
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
		If Is64BitSystem (".") = true then 
		   ' ������ ��������
			call WriteLog("JAVA64 = " & sJavaVersionInstalledGet(strComputer) ,1)
			' ��� 64 ������ ������ ������������� ���������� ������ 32� ���
			call WriteLog("JAVA32 = " & sJavaWoW6432VersionInstalledGet(strComputer) ,1)
		else
			call WriteLog("JAVA   = " & sJavaVersionInstalledGet(strComputer) ,1)
		End If

		If Is64BitSystem (".") = true then 
			
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
	Set fIn = objFSO.OpenTextFile(left (WScript.ScriptFullName, Len(WScript.ScriptFullName)-Len(WScript.ScriptName)) & "FJUpdater.vbs", 1)
	
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
		Call WriteLog("Main : ERROR! can't connect to SMTP server!", 1)
		Exit Sub
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
		.Item(cdoURL & "smtpconnectiontimeout") = 60
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
