'-----------------------------------------------------------------------------------------------------------------------
'
' Copyright (c) Ilya Kisleyko. All rights reserved.
'
' AUTHOR  : Ilya Kisleyko
' E-MAIL  : osterik@gmail.com
' DATE    : 01.08.2015
' NAME    : strUtils.vbs
' COMMENT : Several functions for working with words, idea based on RXLib's strUtils.pas
'
' 1) Function WordCount(sIn, aDelims) : given a set of word delimiters, returns number of words in sIn.
' WScript.Echo WordCount("Hello, How are# you! today?", caDelims)
' =5
'2) Function ExtractWord (iPos, sIn, aDelims) :	returns the iPos'th word in sIn
' WScript.Echo ExtractWord(0, "Hello, How are# you! today?", caDelims)
' = "Hello"

Option Explicit
On Error Resume Next

'default set of word delimiters
dim caDelims
caDelims = Array(".", ",", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "+", "=", "|", "\", "/", "?", ":", ";", "(", ")", "{", "}", """", "'")

Function NormalizeString(sIn, aDelims)
	dim i,sOut, aOut, iMax

	'заменяем символы-разделители на пробелы
	sOut = sIn
	for i=LBound(aDelims) to UBound(aDelims)
		sOut = Replace(sOut, aDelims(i), " ")
	next

	'преобразовываем в массив по разделителям (пробелам)
	aOut = split(sOut, " ")

	'собираем строку обратно
	sOut = ""
	iMax = UBound(aOut)
	for i=LBound(aOut) to iMax
		if aOut(i) <> "" then
			sOut = sOut + aOut(i) + " "
		End if
	next
	'обрезаем лишний пробел в конце
	NormalizeString = Trim(sOut)
End Function

Function WordCount(sIn, aDelims)
	'WordCount given a set of word delimiters, returns number of words in sIn.
	dim sOut, aOut
	sOut = NormalizeString (sIn, aDelims)
	aOut = split(sOut, " ")
	WordCount = UBound(aOut) + 1
End Function

Function ExtractWord (iPos, sIn, aDelims)
	'returns the iPos'th word in sIn
	dim sOut, aOut
	sOut = NormalizeString (sIn, aDelims)
	aOut = split(sOut, " ")
	ExtractWord = aOut(iPos)
End Function
