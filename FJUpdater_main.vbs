'-----------------------------------------------------------------------------------------------------------------------
'
' Copyright (c) Ilya Kisleyko. All rights reserved.
' 
' AUTHOR  : Ilya Kisleyko
' E-MAIL  : osterik@gmail.com
' DATE    : 14.09.2014
' NAME    : FJUpdater_main.vbs
' COMMENT : Скрипт для проверки актуальности установленных на компьютере версий Java&Flash
'          и автоматического обновления (через интернет(HTTP) или локальную сеть(SMB)) при необходимости.
'          Скрипт может сохранять установщики актуальной версии Java&Flash в указанные папки (+ файл с номером версии),
'          что даёт возможность строить механизм автоматического развёртывания обновлений Java&Flash через GPO
'          (стартап скрипт компьютеров).
'          Об ошибках и результатах своей деятельности может присылать отчёт на почту.
'                                                                                                                       
'При запуске без параметров просто выполняет обновление установленных плагинов через интернет
'
'Параметры командной строки:
'/mail:1 = отправлять почту при ошибках работы или при наличии обновлений 
'/mail:0 = тихий режим (default)
'
'/debug:0 = не выводить ничего
'/debug:1 = записывать только сообщения о выполняемых действиях (default) и ошибки
'/debug:2 = записывать подробные сообщения 
'/debug:3 = записывать детальные сообщения о вызываемых функцияих и их парметрах
'
'/WebMode:1 = проверять обновления напрямую из интернета (default)
'/WebMode:0 = из локальной папки (csInstallerPath)
'
'/WEBModeSaveInstall:1 = сохранять обновления в локальную папку (csInstallerPath)
'/WEBModeSaveInstall:0 = сохранять обновления в %TEMP% (default)
'ключ /WEBModeSaveInstall имеет смысл только при /WEBMode:1
'
'/WEBModeSaveInstallForce = особый режим, принудительно скачать последение версии установщиков
'и сохранить в локальную папку (csInstallerPath). не проверяет установленные версии плагинов.
'Используется на компьютере администратора при запуске механизма автоматического развёртывания обновлений 
'
'/MailTest = особый режим, проверка настроек отправки почты, присылает тестовое сообщение
'
'/ShowVersion[:comp] = особый режим, выводит номер весий установленных на компьютере плагинов Java & Flash
' При запуске без параметров - на локальном компьютере, если есть параметр - он трактуется как имя\адрес компьютера
'
'Если при разборе параметров встречается параметр с примечанием "особый режим", то он выполняется немедленно,
'с учётом уже обработанных модификаторов (mail & debug) и производится выход.
'
'/? или /help = справка 
'-----------------------------------------------------------------------------------------------------------------------
'[HELP_END]
' строка "'[HELP_END]" используется как ограничитель для функциии вывода справки

Option Explicit

Const IsDebug = 1 ' флаг отладки
If IsDebug <> 1 then On Error Resume Next

' имена файлов с установщиками\номером актуальной версии, которые лежат в локальном репозитории
Const csJavaInstaller		= "java_installer.exe"
Const csJavaInstaller64		= "java_installer64.exe"
Const csJavaInstallerVers	= "java_current.txt"
Const csFlashPInstaller		= "flashP_installer.exe"	' Flash Plugin (Firefox, Mozilla, Netscape, Opera)
Const csFlashPInstallerVers	= "flashP_current.txt"
Const csFlashAInstaller		= "flashA_installer.exe"	' Flash ActiveX (IE)
Const csFlashAInstallerVers	= "flashA_current.txt"

'===================================================================================================
' где смотреть номера актуальных версий ((при /WEBMode+)
Const csJavaVersionCurrentLnk = "http://www.java.com/applet/JreCurrentVersion2.txt"
Const csFlashVersionCurrentLnk = "http://www.adobe.com/software/flash/about/"
' и где искать\качать инсталляшки
Const csJavaInstallerLink   = "http://java.com/en/download/windows_manual.jsp"
Const csFlashPInstallerLink = "http://fpdownload.macromedia.com/pub/flashplayer/latest/help/install_flash_player.exe"
Const csFlashAInstallerLink = "http://fpdownload.macromedia.com/pub/flashplayer/latest/help/install_flash_player_ax.exe"
' ключи запуска установщиков
Const csJavaInstallerParams = "/s /v /qn IEXPLORER=1 MOZILLA=1 REBOOT=ReallySuppress JAVAUPDATE=0 WEBSTARTICON=0"
Const csFlashInstallerParams = "/install"

'===================================================================================================
' для доступа к реестру, функция RegRead
Const HKEY_CLASSES_ROOT  	= &H80000000
Const HKEY_CURRENT_USER 	= &H80000001
Const HKEY_LOCAL_MACHINE 	= &H80000002
Const HKEY_USERS 		= &H80000003
Const HKEY_CURRENT_CONFIG  	= &H80000005

' глобальные настройки, могут быть переопределены из командной строки
Dim glbMail 					' /mail+ = отправлять почту при ошибках работы или при наличии обновлений 
								' /mail- = тихий режим (default)
Dim gliDebug					' /debug0 = не выводить ничего
								' /debug1 = записывать только сообщения о выполняемых действиях (default)
								' /debug2 = записывать подробные сообщения 
								' /debug3 = записывать детальные сообщения о вызываемых функцияих и их парметрах
Dim glbWEBMode 					' /WEBMode+ = проверять обновления напрямую из интернета (default)
								' /WEBMode- = из локальной папки (csInstallerPath)
Dim glbWEBModeSaveInstall		' имеет смысл только при /WEBMode+
								' /WEBModeSaveInstall+ = сохранять обновления в локальную папку (csInstallerPath)
								' /WEBModeSaveInstall- =  сохранять обновления в %TEMP% (default)
'Dim glbComputer					' имя(или адрес) компьютера, на котором проверяем версии плагинов
' глобальные переменные
Dim sInstallerPath							' путь сохранения инсталляшек при /WEBModeSaveInstall+
Dim sLog 									' общий лог, отсылается на почту при ошибках или появлении обновления
Dim bNeedToSendLog							' флаг необходимости отправки лога на почту
Dim sJavaNativeUpdateStatus, sJavaWoW6432UpdateStatus, sFlashAUpdateStatus, sFlashPUpdateStatus		' итоговый статус операций обновления
Dim sJavaVersionCurrent, sFlashPVersionCurrent, sFlashAVersionCurrent ' номер версии при получении обновлений, используется для записи файла с номером версии

Sub Main
	If IsDebug <> 1 then On Error Resume Next
	
	' Abort If the host is not cscript
	If not IsHostCscript() then
		call wscript.echo("Рекомендуется запускать данный скрипт с использованием CScript.")
		'"Рекомендуется запускать данный скрипт с использованием CScript."  	& vbCRLF & vbCRLF &  
		'"Для этого:"														& vbCRLF & vbCRLF & _
		'"1. Используйте ""CScript " & WScript.ScriptName & " /параметры"" "	& vbCRLF & vbCRLF & _
		'"или" 																& vbCRLF & vbCRLF & _
		'"2. Замените  WScript на  CScript (может"									& vbCRLF & _
		'"   Для этого выполните ""CScript //H:CScript //S"" "				& vbCRLF & _
		'"   и запускайте скрипт """ & WScript.ScriptName & " /параметры""."	& vbCRLF
		wscript.quit
	End If
	
	If bParseCommandLine = false Then
		' не смогли разобрать параметры командной строки
		Call WriteLog("Main, ERROR! can't parse command line switches!", 1)
		Exit Sub 
	End If
	
	Call WriteLog(FormatDateTime(Now,2) & " JavaFlashUpdaterAdmin v2.6 (c) Ilya Kisleyko, osterik@gmail.com", 2)
	
	'путь сохранения инсталляшек...
	If glbWEBMode then 
		If glbWEBModeSaveInstall Then
			sInstallerPath = csInstallerPath		' ...при /WEBModeSaveInstall+
		Else
			sInstallerPath = sEnvGet("%temp%")		' ...при /WEBModeSaveInstall-
		End If
	else
		sInstallerPath = csInstallerPath			' ..при локальном режиме
	End If
	
	' при необходимости добавляем слэш в конце рабочего каталога
	If Right(sInstallerPath,1) <> "\" Then 
		sInstallerPath = sInstallerPath + "\"
	End If
	' пока нет необходимости отправлять логи на почту
	bNeedToSendLog = false
	
	' статус операций
	sJavaNativeUpdateStatus   = "UNKNOWN"
	sJavaWoW6432UpdateStatus  = "UNKNOWN"
	sFlashAUpdateStatus       = "UNKNOWN"
	sFlashPUpdateStatus       = "UNKNOWN"
	
	Call WriteLog("---------------------------------------------------------------------------------------------------",2)
	Call WriteLog("JavaNative - START",2)
	' JAVA Native (32 on 32, 64 on 64) - сравним версии актуальную и установленную
	' что принимать за актуальную версию ...
	If glbWEBMode Then ' WWW
		sJavaVersionCurrent = sJavaVersionWEBGet
	Else               ' локальное хранилище
		sJavaVersionCurrent = sJavaVersionLocalGet	
	End If
	If sJavaVersionCurrent <> "" then 
		If  sJavaVersionInstalledGet(".") <> "" then ' на компьютере присутствует какая-то версия
			If sJavaVersionInstalledGet(".") <> sJavaVersionCurrent then
				' не совпадают, обновляем
				Call WriteLog("WARNING! Installed and current Native(32 on 32, 64 on 64) Java version is NOT identical",1)
				' надо будет отправить отчёт
				bNeedToSendLog = True
				If glbWEBMode Then
					' получаем ссылку на скачку и скачиваем установщик
					If Is64BitSystem (".") = true then
						Call HttpGetSave(sJavaGetLinkToDownload("64"), sInstallerPath & csJavaInstaller64) 
					else
						Call HttpGetSave(sJavaGetLinkToDownload("32"), sInstallerPath & csJavaInstaller) 
					End If
				End If
				' устанавливаем
				If Is64BitSystem (".") = true then
					Call myRun(sInstallerPath & csJavaInstaller64 & " " & csJavaInstallerParams)
				else
					Call myRun(sInstallerPath & csJavaInstaller & " " & csJavaInstallerParams)
				End If
			
				' повторно проверяем результат
				If sJavaVersionInstalledGet(".") = sJavaVersionCurrent then
					sJavaNativeUpdateStatus = "SUCCESS"
					If glbWEBModeSaveInstall then
						' записываем файл с номером актуальной версии установщика
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
	' JAVA 32bit on 64bit system) - сравним версии актуальную и установленную
	' что принимать за актуальную версию ...
	If glbWEBMode Then ' WWW
		sJavaVersionCurrent = sJavaVersionWEBGet
	Else               ' локальное хранилище
		sJavaVersionCurrent = sJavaVersionLocalGet	
	End If
	If sJavaVersionCurrent <> "" then 
		If  sJavaWoW6432VersionInstalledGet(".") <> "" then ' на компьютере присутствует какая-то версия
			If sJavaWoW6432VersionInstalledGet(".") <> sJavaVersionCurrent then
				' не совпадают, обновляем
				Call WriteLog("WARNING! Installed and current WoW6432(32bit Java on 64bit OS) Java version is NOT identical",1)
				' надо будет отправить отчёт
				bNeedToSendLog = True
				If glbWEBMode Then
					' получаем ссылку на скачку и скачиваем установщик
					Call HttpGetSave(sJavaGetLinkToDownload, sInstallerPath & csJavaInstaller) 
				End If
				' устанавливаем
				Call myRun(sInstallerPath & csJavaInstaller & " " & csJavaInstallerParams)
				
				' повторно проверяем результат
				If sJavaWoW6432VersionInstalledGet(".") = sJavaVersionCurrent then
					sJavaWoW6432UpdateStatus = "SUCCESS"
					If glbWEBModeSaveInstall then
						' записываем файл с номером актуальной версии установщика
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
	' FlashActiveX - сравним версии актуальную и установленную
	' что принимать за актуальную версию ...	
	If glbWEBMode Then
		sFlashAVersionCurrent = sFlashVersionWEBGet("A")
	Else
		sFlashAVersionCurrent = sFlashVersionLocalGet("A")
	End If
	
	If sFlashAVersionCurrent <> "" then 
		If sFlashVersionInstalledGet(".", "A") <> "" then ' на компьютере присутствует какая-то версия
			If sFlashVersionInstalledGet(".", "A") <> sFlashAVersionCurrent then
				' не совпадают, обновляем
				Call WriteLog("WARNING! Installed and current FlashActiveX version is NOT identical",1)
				' надо будет отправить отчёт
				bNeedToSendLog = True
				If glbWEBMode Then
					' получаем ссылку на скачку и скачиваем установщик
					Call HttpGetSave(csFlashAInstallerLink, sInstallerPath & csFlashAInstaller)
				End If
				' устанавливаем
				Call myRun(sInstallerPath & csFlashAInstaller & " " & csFlashInstallerParams)
				' повторно проверяем результат
				If sFlashVersionInstalledGet(".", "A") = sFlashAVersionCurrent then
					sFlashAUpdateStatus = "SUCCESS"
					' записываем файл с номером актуальной версии установщика
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
	' FlashPlugin - сравним версии актуальную и установленную
	' что принимать за актуальную версию ...	
	If glbWEBMode Then
		sFlashPVersionCurrent = sFlashVersionWEBGet("P")
	Else
		sFlashPVersionCurrent = sFlashVersionLocalGet("P")
	End If
	If sFlashPVersionCurrent <> "" Then 
		If sFlashVersionInstalledGet(".", "P") <> "" then ' на компьютере присутствует какая-то версия
			If sFlashVersionInstalledGet(".", "P") <> sFlashPVersionCurrent then
				' не совпадают, обновляем
				Call WriteLog("WARNING! Installed and current FlashPlugin version is NOT identical",1)
				' надо будет отправить отчёт
				bNeedToSendLog = True
				If glbWEBMode Then
					' получаем ссылку на скачку и скачиваем установщик
					Call HttpGetSave(csFlashPInstallerLink, sInstallerPath & csFlashPInstaller)
				End If
				' устанавливаем
				Call myRun(sInstallerPath & csFlashPInstaller & " " & csFlashInstallerParams)
				' повторно проверяем результат
				If sFlashVersionInstalledGet(".", "P") = sFlashPVersionCurrent then
					sFlashPUpdateStatus = "SUCCESS"
					' записываем файл с номером актуальной версии установщика
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
	
	' при необходимости отправляем отчёт
	If bNeedToSendLog and glbMail then
		Call WriteLog("Sending log to e-mail",2)
		Call mySendMail("Updater Java=[" & sJavaNativeUpdateStatus &_
			"], FlashA =[" & sFlashAUpdateStatus & "], FlashP =[" & sFlashPUpdateStatus &_
			"] on " & sEnvGet("%COMPUTERNAME%"), sLog)
	End If
	
	Call WriteLog("===================================================================================================",2)
End Sub ' Main

Function sJavaGetLinkToDownload (sBits)
	' парсит страничку csJavaInstallerLink
	' и возвращает ссылку на скачку актуальной версии Java
	' если sBits = 64, то возвращает ссылку на 64-битную версию, иначе = 32х
	If IsDebug <> 1 then On Error Resume Next
	Dim sHttpText				' содержимое скаченной страницы
	Dim objRegExp				' для получения ссылки на установщик
	Dim objMatches, objMatch	' вспомогательные объекты при парсинге html
	
	' скачиваем страницу выбора фалов для ручной загрузки Java в sHttpText
	Call WriteLog("sJavaGetLinkToDownload(" & sBits &"), Downloadning HTML",3)
	sHttpText = sHttpGet(csJavaInstallerLink)
	If sHttpText = "" then
		Call WriteLog("sJavaGetLinkToDownload, ERROR! Failed to get the html by the link " & csJavaInstallerLink,1)
		bNeedToSendLog = True
		Exit Function
	End If
	
	' начинаем разбирать HTML, ищем тег <a>
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
		'ищем href
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
			' обрезаем лишние символы
			sHttpText=Mid(objMatch.Value,7,len(objMatch.Value)-7)
			Call WriteLog("sJavaGetLinkToDownload, OK! Download link = " & sHttpText,3)
			sJavaGetLinkToDownload = sHttpText
		End If
	End If
End Function ' sJavaGetLinkToDownload

Function sJavaVersionInstalledGet (strComputer)
	' возвращает строку с номером версии Java, установленной на компьютере strComputer, или локальном компьютере (если пусто)
	' возвращает версию совпадающую с разрядностью системы (32 для 32, 64 для 64)
	sJavaVersionInstalledGet = sJavaVersionInstalledGetA (strComputer, "")
End Function ' sJavaVersionInstalledGet
Function sJavaWoW6432VersionInstalledGet (strComputer)
	' возвращает строку с номером версии Java, установленной на компьютере strComputer, или локальном компьютере (если пусто)
	' в sWoW6432 = возвращает версию 32-битной Javа на 64-битной системе
	' только на 64битных системых
	
	If Is64BitSystem(strComputer) then
		sJavaWoW6432VersionInstalledGet = sJavaVersionInstalledGetA (strComputer, "\Wow6432Node")
	Else
		sJavaWoW6432VersionInstalledGet = ""
	End If
End Function ' sJavaWoW6432VersionInstalledGet

	
Function sJavaVersionInstalledGetA (strComputer, sRegPathModifier)
	'sRegPathModifier -  добавляется в путь, если проверяется 32битная версия на 64битной системе. 
	' должен быть = "\Wow6432Node"

	If IsDebug <> 1 then On Error Resume Next
	
	Dim sJavaMajorVersionInstalled, sJavaVersionInstalled
	If strComputer = "" Then strComputer = "."
	' узнаем поколение установленной Java
	sJavaMajorVersionInstalled = sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE" & sRegPathModifier & "\JavaSoft\Java Runtime Environment", "CurrentVersion")
	sJavaMajorVersionInstalled = Right(sJavaMajorVersionInstalled,1)
	' узнаем полный номер версии установленной Java
	sJavaVersionInstalled  = sRegRead(strComputer, HKEY_LOCAL_MACHINE, "SOFTWARE" & sRegPathModifier & "\JavaSoft\Java Runtime Environment", "Java" & sJavaMajorVersionInstalled & "FamilyVersion")
	If sJavaVersionInstalled = "" then
		Call WriteLog("sJavaVersionInstalledGetA, INFO: Can't get the installed version of Java from the registry. Probably Java on this computer is not installed",3)
		'bNeedToSendLog = True
		'Exit Function
	End If
	Call WriteLog("sJavaVersionInstalledGetA (" & sRegPathModifier & "), JavaVersionInstalled = " & sJavaVersionInstalled,2)
	sJavaVersionInstalledGetA = sJavaVersionInstalled
End Function ' sJavaVersionInstalledGetA

Function sJavaVersionWEBGet
	' возвращает строку с номером актуальной версии JAVA, берёт с сайта
	If IsDebug <> 1 then On Error Resume Next
	Dim  sJavaVersionWEB
	
	sJavaVersionWEB = sHttpGet (csJavaVersionCurrentLnk)
	sJavaVersionWEB = Replace(sJavaVersionWEB, vbCrLf, "") ' убираем лишний перевод строки
	If sJavaVersionWEB = "" then ' не смогли получить актуальную версию Java по ссылке
		Call WriteLog("sJavaVersionWEBGet, ERROR! Can't get current version of Java by the link " & csJavaVersionCurrentLnk,1)
		bNeedToSendLog = True
		Exit Function
	End If
	sJavaVersionWEBGet = sJavaVersionWEB
	Call WriteLog("sJavaVersionWEBGet, JavaVersionWEB   = " & sJavaVersionWEB,2)
End Function ' sJavaVersionWEBGet

Function sJavaVersionLocalGet
	' возвращает строку с номером актуальной версии JAVA, берёт с локального репозитория
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
	' возвращает строку с номером версии Flash, установленной на компьютере strComputer, или локальном компьютере (если пусто)
	' sFlashType - тип флэш-проигрывателя, A=ActiveX, P=Plugin
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
	' возвращает строку с номером актуальной версии Flash, берёт с сайта
	' sFlashType - тип флэш-проигрывателя, A=ActiveX, P=Plugin	
	If IsDebug <> 1 then On Error Resume Next
	Dim sTemp
	Dim sHttpText				' содержимое загруженной страницы
	Dim objRegExp				' для получения номера актуальной версии
	Dim objMatches			 	' вспомогательный объект при парсинге html
	
	sHttpText = sHttpGet (csFlashVersionCurrentLnk)
	If sHttpText = "" then
		Call WriteLog("sFlashVersionWEBGet, ERROR! Can't get HTML by the link  " & csFlashVersionCurrentLnk,1)
		bNeedToSendLog = True
		Exit Function
	End If
	' начинаем разбирать HTML,ищем ячейки таблицы с последовательностью цифр, разделенных точками.
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
			' !!! изменить номер Item(x) при изменении на сайте Adobe по ссылке csFlashVersionCurrentLnk !!!
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
	' возвращает строку с номером актуальной версии Flash, берёт с локального репозитория
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
' служебные процедуры и функции - доступ к файлам, реестру, http, отправка почты и т.п.

Function bParseCommandLine
	' разбирает параметры командной строки, устанавливает значения по умолчанию для глобальных параметров								
	' возвращает false при появлении ошибок
	If IsDebug <> 1 then On Error Resume Next

	glbMail = False
	gliDebug = 1
	glbWEBMode = True
	glbWEBModeSaveInstall = false
	
	bParseCommandLine = true
	dim objNamed
	Set objNamed = WScript.Arguments.Named

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

	If objNamed.Exists("debug") Then 
		select case lcase(objNamed.Item("debug"))
			case "0","1","2","3"
				gliDebug = cint(objNamed.Item("debug"))
				Call WriteLog ("bParseCommandLine, gliDebug = " & gliDebug,3)
			case else
				Call WriteLog ("bParseCommandLine, gliDebug ERROR",1)
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
		' просто выкачиваем все обновления
		sInstallerPath = csInstallerPath		' путь сохранения инсталляшек при /WEBModeSaveInstall+
		' при необходимости добавляем слэш в конце рабочего каталога
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
			' Пингом проверим что комп включен
			If IsPingSucsess (strComputer) Then
				WScript.Echo ("компьютер " & strComputer & " доступен!")
			Else
				WScript.Echo ("компьютер " & strComputer & " НЕ доступен!")
				WScript.Quit
			End If
		End If
		' TODO - переделать вывод отчёта
		If Is64BitSystem (".") = true then 
		   ' родная битность
			call WriteLog("JAVA64 = " & sJavaVersionInstalledGet(strComputer) ,1)
			' для 64 битных систем дополнительно показываем версию 32й явы
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
		'todo : вывод Help'а
	End If
End Function

sub PrintHelp
	' выводит справку
	If IsDebug <> 1 then On Error Resume Next
	Call WriteLog("PrintHelp", 0)	
	Dim objFSO, fIn, sTemp, sTemp2

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	' TODO = сделать парсинг WScript.ScriptName и замену wsf на vbs (или txt)
	' из полного имени скрипта вырезаем имя скрипта, получаем путь запуска скрипта
	Set fIn = objFSO.OpenTextFile(left (WScript.ScriptFullName, Len(WScript.ScriptFullName)-Len(WScript.ScriptName)) & "FJUpdater.vbs", 1)
	
	Do While (Not fIn.AtEndOfStream) 
		sTemp = fIn.ReadLine
		If sTemp = "'[HELP_END]" then wscript.quit
		If sTemp <> "" then Wscript.Echo (Mid(sTemp,2,Len(sTemp)-1))
	Loop
	
end Sub

Sub mySendMail (sSubject, sTextBody)
	' отправляем почту через CDO
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
		'важность
		.fields.Item("urn:schemas:mailheader:importance").Value = "high"	
		'срочность 
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
	' добавляет строку в лог для отправки и вывод в консоль, если iDebug <= gliDebug 
	If IsDebug <> 1 then On Error Resume Next
	If iDebug <= gliDebug then
		Dim sTemp
		sTemp = FormatDateTime(Now,3) & " " & sText
		Wscript.Echo(sTemp)
		sLog = sLog & vbCRLF & sTemp
	End If
End Sub

Sub WriteLogError(sText, iNumber, sDescription, sSource)
	' добавляет строку с ошибкой в лог
	If IsDebug <> 1 then On Error Resume Next
	Call WriteLog(sText & ", ERROR! : " & iNumber & ": " & sDescription & " generated by " & sSource, 2)
	If gliDebug >= 2 Then ' если высокая детализация отладки
		bNeedToSendLog = True ' сообщить об ошибке почтой
	End If
End Sub

function sRegRead (strComputer, regHive, sKeyPath, sKeyName)
	' возвращает значение ключа реестра sKeyPath + sKeyName компьютера strComputer
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
	' возвращает содержимое файла
	If IsDebug <> 1 then On Error Resume Next
	Call WriteLog("sFileRead, sFile=" & sFile, 3)	
	
	
	Dim objFSO, fIn
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(sFile) then
		Set fIn = objFSO.OpenTextFile(sFile, 1)
		sFileRead = Replace(fIn.ReadAll, vbCrLf, "") ' читаем весь файл и убираем переводы строки
		fIn.Close
		Set fIn = nothing
	Else
		Call WriteLog("ERROR : sFileRead, file not found, sFile=" & sFile, 2)	
	End If
	
	Set objFSO = Nothing
End Function

Function sEnvGet (sEnv)
	' возвращает значение переменной окружения
	On Error Resume Next	
	
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	sEnvGet = objShell.ExpandEnvironmentStrings(sEnv)
	If Err.Number <> 0 then Call WriteLogError("sEnvGet", Err.Number, Err.Description, Err.Source)
End Function

Function myRun(sCommand)
	' запускает sCommand и ждёт выполнения, возвращает вывод комманды
	If IsDebug <> 1 then On Error Resume Next
	Dim objShell, objShellExec
	
	Const ciShellExecStatusRun  = 0 'приложение запущено.
	Const ciShellExecStatusStop = 1 '1 - приложение завершено.
	
	Set objShell = CreateObject("WScript.Shell")	
	Set objShellExec = objShell.Exec(sCommand)
	' ждём, пока отработает скрипт
	While objShellExec.Status = ciShellExecStatusRun
		WScript.Sleep 100
	Wend
	myRun = objShellExec.StdOut.ReadAll
End Function

Function sHttpGet (sURL) 
	' возвращает содержимое загруженное по ссылке sURL
	If IsDebug <> 1 then On Error Resume Next
	Dim objHttp
	
	Call WriteLog("sHttpGet, sURL=" & sURL,3)
	Set objHttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")
	objHttp.Open "GET", sURL, False
	' эмулируем IE10 64bit
	objHttp.setRequestHeader "User-Agent", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)"
	objHttp.Send
	If Err.Number <> 0 then Call WriteLogError("sHttpGet", Err.Number, Err.Description, Err.Source)
	sHttpGet = objHttp.ResponseText
	Set objHttp = nothing
End Function

Sub HttpGetSave(sURL, sFile) 
	' сохраняет содержимое загруженное по ссылке sURL в файл sFile
	If IsDebug <> 1 then On Error Resume Next
	Dim objHttp, objADOStream, objFSO
	
	Call WriteLog("HttpGetSave, sURL=" & sURL & ", sFile=" & sFile,3)
	Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")
	objHttp.Open "GET", sURL, False	
	' эмулируем IE10 64bit
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
	' возвращает True, если компьютере strComputer 64битный,  если strComputer = пусто то на локальном компьютере
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
	' возвращает True, если компьютере strComputer отвечает на пинг
	If IsDebug <> 1 then On Error Resume Next

	dim sTemp 
	sTemp = LCase(myRun ("%comspec% /c ping.exe -n 2 " & strComputer))
	If InStr(sTemp, "ttl=")  Then
		IsPingSucsess = True
	else
		IsPingSucsess  = False
	End If
End Function
