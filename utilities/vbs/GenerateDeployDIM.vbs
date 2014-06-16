Option Explicit

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const Pattern0 = "(.*),(.+)"

' Entering main procedure
Call Main()

Sub Main()

	Dim regExp, rowMatch, row
	Dim objFSO, objFile, objWFile
	Dim fCsvFile, fTempFile
	Dim strLine, strPath, strFile
	
	if WScript.Arguments.Count < 2 Then
		WScript.Echo "Usage: CScript //Nologo GenerateDeployDIM.vbs <CSV-File> <Output-File>"
		WScript.Quit 1
	End if	

	fCsvFile=WScript.Arguments(0)
	fTempFile=WScript.Arguments(1)

	If Not FileExists(fCsvFile) Then
		WScript.Quit 1
	End If

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(fCsvFile, ForReading, False )
	Set objWFile = objFSO.OpenTextFile(fTempFile, ForWriting, True)
	
	Set regExp = New regExp
	
	Do While objFile.AtEndOfStream<>True
		'讀入檔案內容
		strLine = objFile.ReadLine()

		regExp.Pattern = Pattern0
		If regExp.Test(strLine) Then
			Set rowMatch = regExp.Execute(strLine)
			Set row = rowMatch(0)
			strPath = row.SubMatches(0)
			strFile = row.SubMatches(1)
			
			strPath = Replace(strPath, "\", "/")
			If Left(strPath, 1) <> "/" Then strPath = "/" & strPath
			If Right(strPath, 1) <> "/" Then strPath = strPath & "/"

			objWFile.WriteLine(strPath & strFile & " " & strPath & " CPY")
		End If			
	Loop

End Sub


' ================================================================================
' FileExists Function
'
' Arguments:
' strFilePath  [string] to check the path and file name whether exists
'
' Returns:
' [bool] True for exists; False for not exists
'
Function FileExists(ByVal strFilePath)

	Dim objFSO

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFilePath) Then
		FileExists = True
	Else
		WScript.Echo strFilePath & " doesn't exists."
		FileExists = False
	End If

End Function
' ================================================================================
