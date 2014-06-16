Option Explicit

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const Pattern0 = "(.*),(.+),""(.*)"""

' Entering main procedure
Call Main()

Sub Main()
	
	Dim regExp, rowMatch, row
	Dim objFSO, objFile, objWFile, strLine, strDMCLI, strCMD_UIA
	Dim fParamFile, fIniFile, fCsvFile, fTempFile, fLogFile, fLog2File
	Dim mProduct, mProject, mLocation, mChangeDocId
	Dim strPath, strFileName, strMemo
	
	if WScript.Arguments.Count < 3 Then
		WScript.Echo "Usage: CScript //Nologo UIA.vbe <PARAM-File> <INI-File> <CSV-File>"
		WScript.Quit 1
	End if	
	
	fParamFile=WScript.Arguments(0)
	fIniFile=WScript.Arguments(1)
	fCsvFile=WScript.Arguments(2)
	fTempFile=GetTempFile()
	fLogFile=GetTempFile()

	If Not FileExists(fCsvFile) Then
		WScript.Quit 1
	End If
	
	mProduct=ReadINI(fIniFile, "PROJECT_INFO", "PRODUCT")
	mProject=ReadINI(fIniFile, "PROJECT_INFO", "PROJECT")
	mLocation=ReadINI(fIniFile, "PROJECT_INFO", "LOCATION")
	mChangeDocID=ReadINI(fIniFile, "PROJECT_INFO", "CHANGE_DOC_ID")
	
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(fCsvFile, ForReading, False )
	Set objWFile = objFSO.OpenTextFile(fTempFile, ForWriting, True)
	
	Set regExp = New RegExp
	
	Do While objFile.AtEndOfStream<>True
		'讀入檔案內容
		strLine = objFile.ReadLine()
		
		regExp.Pattern = Pattern0
		If regExp.Test(strLine) Then
			Set rowMatch = regExp.Execute(strLine)
			Set row = rowMatch(0)
			strPath = row.SubMatches(0)
			strFileName = row.SubMatches(1)
			strMemo = row.SubMatches(2)

			strMemo = Replace(strMemo, "@", "@@")
			strMemo = Replace(strMemo, ",", "@,")
			strMemo = Replace(strMemo, "(", "@(")
			strMemo = Replace(strMemo, ")", "@)")
			strMemo = Replace(strMemo, """", "@""")
			strMemo = Replace(strMemo, "'", "@'")
			strMemo = Replace(strMemo, "/", "@/")
			strMemo = Replace(strMemo, "\", "@\")

			'UIA "QLARIUS:;" /WORKSET="QLARIUS:QUOTE_WEB_1.0" /FILENAME="Qlarius_Home\images\auto-off.gif" /ATTRIB=(Description="memo for auto-off") 
			strCMD_UIA="UIA ""##PRODUCT##"" /WORKSET=""##PROJECT##"" /FILENAME=""##FILENAME##"" /ATTRIB=(Description=""##MEMO##"")"
			strCMD_UIA=Replace(strCMD_UIA, "##PRODUCT##", mProduct)
			strCMD_UIA=Replace(strCMD_UIA, "##PROJECT##", mProject)
			strCMD_UIA=Replace(strCMD_UIA, "##FILENAME##", Trim(strPath) & Trim(strFileName))
			strCMD_UIA=Replace(strCMD_UIA, "##MEMO##", strMemo)
			objWFile.WriteLine(strCMD_UIA)
		End If
	Loop

	strDMCLI="dmcli -param ##PARAMFILE## -file ##CMDFILE## > ##LOG##"
	strDMCLI=Replace(strDMCLI, "##PARAMFILE##", fParamFile)
	strDMCLI=Replace(strDMCLI, "##CMDFILE##", fTempFile)
	strDMCLI=Replace(strDMCLI, "##LOG##", fLogFile)
	
	WScript.Echo "========== Command List =========="
	Call ReadFileToEcho(fTempFile)
	WScript.Echo "========== Command Execute =========="
	WScript.Echo strDMCLI
	CreateObject("WScript.Shell").Run "CMD /C " & strDMCLI, 0, True	
	WScript.Echo "========== Operation Detail =========="
	Call ReadFileToEcho(fLogFile)	
	WScript.Echo "========== Finished =========="
	
End Sub


' ================================================================================
' ReadINI Function
'
' Arguments:
' strFilePath  [string]  the (path and) file name of the INI file
' strSection   [string]  the section in the INI file to be searched
' strKey       [string]  the key which value is to be returned
'
' Returns:
' [string] the value for the specified key in the specified section
'
Function ReadINI( strFilePath, strSection, strKey )

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    'Dim strFilePath, strKey, strLeftString, strLine, strSection
    Dim strLeftString, strLine

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( strFilePath )
    strSection  = Trim( strSection )
    strKey      = Trim( strKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If

End Function
' ================================================================================


' ================================================================================
' WriteINI Procedure
'
' Arguments:
' strFilePath  [string]  the (path and) file name of the INI file
' strSection   [string]  the section in the INI file to be searched
' strKey       [string]  the key whose value is to be written
' strValue     [string]  the value to be written (strKey will be
'                       deleted if strValue is <DELETE_THIS_VALUE>)
'
' Returns:
' N/A
'
Sub WriteINI( strFilePath, strSection, strKey, strValue )

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
    Dim intEqualPos
    Dim objFSO, objNewIni, objOrgIni, wshShell
    Dim strFolderPath, strLeftString
    Dim strLine, strTempDir, strTempFile

    strFilePath = Trim( strFilePath )
    strSection  = Trim( strSection )
    strKey      = Trim( strKey )
    strValue    = Trim( strValue )

    Set objFSO   = CreateObject( "Scripting.FileSystemObject" )
    Set wshShell = CreateObject( "WScript.Shell" )

    strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )

    Set objOrgIni = objFSO.OpenTextFile( strFilePath, ForReading, True )
    Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )

    blnInSection     = False
    blnSectionExists = False
    ' Check if the specified key already exists
    blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
    blnWritten       = False

    ' Check if path to INI file exists, quit if not
    strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
    If Not objFSO.FolderExists ( strFolderPath ) Then
        WScript.Echo "Error: WriteIni failed, folder path (" _
                   & strFolderPath & ") to ini file " _
                   & strFilePath & " not found!"
        Set objOrgIni = Nothing
        Set objNewIni = Nothing
        Set objFSO    = Nothing
        WScript.Quit 1
    End If

    While objOrgIni.AtEndOfStream = False
        strLine = Trim( objOrgIni.ReadLine )
        If blnWritten = False Then
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                blnSectionExists = True
                blnInSection = True
            ElseIf InStr( strLine, "[" ) = 1 Then
                blnInSection = False
            End If
        End If

        If blnInSection Then
            If blnKeyExists Then
                intEqualPos = InStr( 1, strLine, "=", vbTextCompare )
                If intEqualPos > 0 Then
                    strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                    If LCase( strLeftString ) = LCase( strKey ) Then
                        ' Only write the key if the value isn't empty
                        ' Modification by Johan Pol
                        If strValue <> "<DELETE_THIS_VALUE>" Then
                            objNewIni.WriteLine strKey & "=" & strValue
                        End If
                        blnWritten   = True
                        blnInSection = False
                    End If
                End If
                If Not blnWritten Then
                    objNewIni.WriteLine strLine
                End If
            Else
                objNewIni.WriteLine strLine
                    ' Only write the key if the value isn't empty
                    ' Modification by Johan Pol
                    If strValue <> "<DELETE_THIS_VALUE>" Then
                        objNewIni.WriteLine strKey & "=" & strValue
                    End If
                blnWritten   = True
                blnInSection = False
            End If
        Else
            objNewIni.WriteLine strLine
        End If
    Wend

    If blnSectionExists = False Then ' section doesn't exist
        objNewIni.WriteLine
        objNewIni.WriteLine "[" & strSection & "]"
            ' Only write the key if the value isn't empty
            ' Modification by Johan Pol
            If strValue <> "<DELETE_THIS_VALUE>" Then
                objNewIni.WriteLine strKey & "=" & strValue
            End If
    End If

    objOrgIni.Close
    objNewIni.Close

    ' Delete old INI file
    objFSO.DeleteFile strFilePath, True
    ' Rename new INI file
    objFSO.MoveFile strTempFile, strFilePath

    Set objOrgIni = Nothing
    Set objNewIni = Nothing
    Set objFSO    = Nothing
    Set wshShell  = Nothing
End Sub
' ================================================================================

' ================================================================================
' GetTempFile Function
'
' Arguments:
' 	N/A
'
' Returns:
' [string] Temp File Name
'
Function GetTempFile()

	Dim objFSO, wshShell
	Dim strTempDir, strTempFile

	Set objFSO   = CreateObject( "Scripting.FileSystemObject" )
	Set wshShell = CreateObject( "WScript.Shell" )

	strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
	strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )

	GetTempFile = strTempFile

End Function
' ================================================================================

' ================================================================================
' ShellRun Procedure
'
' Arguments:
' strFileName  [string] the path and file of the file which user want to execute
'
' Returns:
' [integer] 0 for normal exit; 1 for file not exists.
'
Function ShellRun(ByVal strFileName, ByRef strTempFile)

	strTempFile = GetTempFile()

	If FileExists(strFileName) Then
		If CreateObject("WScript.Shell").Run(strFileName & " > " & strTempFile, 1, True)=0 Then
			ShellRun=0
		Else
			ShellRun=1
		End If
	Else
		ShellRun=1
	End If

End Function
' ================================================================================

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
' ================================================================================
' ReadFileToEcho Procedure
'
' Arguments:
' strFileName  [string]  the (path and) file name
'
' Returns:
' [string] the value for the specified key in the specified section
'
Sub ReadFileToEcho(strFileName)

    	Const ForReading   = 1
    	Const ForWriting   = 2
    	Const ForAppending = 8

    	Dim objFSO, objFile
    	Dim strLine

	Set objFSO = CreateObject( "Scripting.FileSystemObject" )

	strFileName = Trim(strFileName)

    	If objFSO.FileExists(strFileName) Then
        	Set objFile = objFSO.OpenTextFile(strFileName,ForReading,False)
	        Do While objFile.AtEndOfStream=False
			strLine = objFile.ReadLine
			WScript.Echo strLine
	        Loop
        	objFile.Close
    	Else
        	WScript.Echo strFileName & " doesn't exists. Exiting..."
    	End If

End Sub
' ================================================================================
