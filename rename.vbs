' rename.vbs: A VBScript file for renaming jpg files 
' in current directory to taken date. 
' Copyright (C) 2014 ITIIC <http://itiic.com/>
'
' rename.vbs is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' INIFile.vbs is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public License
' along with INIFile.vbs. If not, see <http://www.gnu.org/licenses/>.


' Declaring and setting objects
Dim objFSO 	: Set objFSO 	= Wscript.CreateObject("Scripting.FileSystemObject")
Dim objSh 	: Set objSh 	= Wscript.CreateObject("Wscript.Shell")
Dim objRE 	: Set objRE 	= New RegExp

iCnt 	= 0
strPath = objFSO.GetAbsolutePathName(".")
Dim objDirectory 	: Set objDirectory 	= objFSO.GetFolder(strPath)
Dim objEnv 			: Set objEnv 		= objSh.Environment("PROCESS")


Wscript.Echo "Processing Dir: " & strPath

' LOOP
For Each objFile In objDirectory.files

	ChangeName objFile, strPath
	
Next


' Get Exif data from file in dir
Function GetExif (filename, dir)
	' get tem file name
	strTempName = objFSO.GetTempName
	
	' place temp file in temp dir
    objTempFile = objEnv("tmp") & "\" & strTempName
    
	' cmd : execute exiv2.exe with param filename, redirect errors to nul, grep by Image timestamp and redirect to tmp file
	strCmd = "%comspec% /C "" """ & dir & "\exiv2.exe"" """ & filename & """ 2>nul | " & "find /I """ & "Image timestamp" & """ >" & objTempFile & """"
	
	' execute cmd
	objSh.Run strCmd, 0, True
	
	' open and read temp file
	Dim objTextFile	: Set objTextFile = objFSO.OpenTextFile(objTempFile, 1)
	Do While objTextFile.AtEndOfStream <> True
		strText = objTextFile.ReadLine
	Loop

	' close file and delete object
    objTextFile.Close
    objFSO.DeleteFile(objTempFile)
    
	' return text from temp file
    GetExif = strText
	
End Function


' Chance name of jpg file
Function ChangeName (objFile, Dir)
   
	' get extension of file
	strExtension = "." & lcase(objFSO.getExtensionName(objFile.path))
	
    if strExtension = ".jpg" then
	
		' get timestamp date from exif data of jpg file
		strTimestamp =  GetExif(objFile, Dir)
	
		' parametrise RegEx object
		With objRE
			.Pattern    = "Image timestamp : (\d+):(\d+):(\d+) (\d+):(\d+):(\d+)"
			.IgnoreCase = True
			.Global     = False
		End With		
		
		' execute matching - and get first row 
		Dim objDigMatch : Set objDigMatch = objRE.Execute( strTimestamp )(0)
		
		' create strNewFilename with masl YYYYMMDD_hhmmss
		arRange1 = Array(0, 1, 2)
		arRange2 = Array(3, 4, 5)
		
		For Each item in arRange1
			strNewFilename = strNewFilename & "" & objDigMatch.SubMatches(item)
		Next

		strNewFilename = strNewFilename & "_"
		
		For Each item in arRange2
			strNewFilename = strNewFilename & "" & objDigMatch.SubMatches(item)
		Next
		
		' sumarise 
		strPathOld = objFile.path
		strPathNew = Dir & "\" & strNewFilename & ".jpg"
		
		WScript.Echo "OLD : " & strPathOld
		WScript.Echo "NEW : " & strPathNew
			
		iCnt = iCnt + 1
		Wscript.Echo "Processing File : " & iCnt

		' rename
		objFSO.Movefile strPathOld, strPathNew
			
		
    End If
	
End Function
