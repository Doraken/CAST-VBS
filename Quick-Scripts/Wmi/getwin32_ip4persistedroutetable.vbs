'.NOTES
'	===========================================================================
'	 Created with: 	Notepad ++
'	 Created on:   	27/06/2013 23:22
'	 Created by:   	Arnaud Crampet 
'	 Mail to: 		arnaud@crampet.net 
'	 Organization: 	
'	 Filename:     	 
'	 Operatin system: Windows 2000/XP/VISTA/8/10 or windows server all versions
'	===========================================================================
'	.DESCRIPTION
'		A description of the file.
Option explicit 

Const ForReading = 1     	' used bye file writer as creator
Const ForWriting = 2	 	' used bye file writer as creator
Const ForAppending = 8	 	' used bye file writer as creator

' Default parameter il command line optiaon are omitted.
Const BaseComputerNameDefault 	= "."
Const DebugLevelDefault  		= "1"
Const StrFilePathToWriteDefault = "C:\Log"
Const WSCWriteModeDefault 		= "full"
Const WSCModeDefault 			= "full"

Dim colNamedArguments       	' Collection object to store Script options switch
Dim BaseDebugLevel          	' store debug level for this script.
Dim DebugLevel              	' Defile debug level 0 1 2 3 4 5 6
Dim WSCMode                 	' Define what kind of scan will be lauched 
Dim WSCWriteMode            	' Define what kind of file to write. CSV HTML JS ALL  





Dim objFSO 						' As FileSystemObject
Dim objTextFile 				' As Object
Dim StrFileDate             	' store filename date ext 
Dim StrDate    		        	' store  date 
Dim DbgPath						' Used to store StackTrace
Dim objDictionary
Dim BaseComputerName 			' String var used to store Computer Name 

Dim StrFileToRead       		' Storing Full Path of readed file.
Dim StrFilePathToWrite      	' Storing Base Full Path of writed file.
Dim StrFilePathToWriteCSV   	' Storing Full Path of CVS writed file.
Dim StrFilePathToWriteHTML  	' Storing Full Path of HTMl writed file.
Dim StrFilePathToWriteJS    	' Storing Full Path of JS writed file.
Dim StrFilePathToWriteCSVCmp	' Path for computer spécifique report files 
Dim StrFilePathToWriteHTMLCmp	' Path for computer spécifique report files 
Dim StrFilePathToWriteJSCmp		' Path for computer spécifique report files 

Function ProcessArgs
	'-------------------------------------------------------------------------------
	' ProcessArgs - Used to display script help information.
	' Parameters:
	'
	' Returns:
	'   var managment.
	'-------------------------------------------------------------------------------
	Dim colNamedArguments
	'Dim BaseComputerName
	
	Set colNamedArguments = WScript.Arguments.Named

	If WScript.Arguments.Count = "0" Then 
		Call  DisplayHelp 
	Else 
		If not colNamedArguments.Item("help") = ""  then
		   Call  DisplayHelp 
		End If 
		
		if not colNamedArguments.Item("server") = "" Then 
			BaseComputerName = LCase(colNamedArguments.Item("server"))
		Else 
			BaseComputerName = BaseComputerNameDefault
		End IF
		
		if not colNamedArguments.Item("debug") = "" then 
			DebugLevel = LCase(colNamedArguments.Item("debug"))
		Else 
			DebugLevel = DebugLevelDefault 
		End IF
		
		if not colNamedArguments.Item("reportpath") = "" then 
			StrFilePathToWrite = LCase(colNamedArguments.Item("reportpath"))
		Else 
			StrFilePathToWrite = StrFilePathToWriteDefault
		End IF 
		
		if not colNamedArguments.Item("reporttype") = "" then 
				WSCWriteMode = LCase(colNamedArguments.Item("reporttype"))
		Else 
				WSCWriteMode = WSCWriteModeDefault
		End IF
		if Not colNamedArguments.Item("mode") = "" Then 
			WSCMode = LCase(colNamedArguments.Item("mode"))
		Else 
			WSCMode = WSCModeDefault
		End If
	End If
		
	if WSCWriteMode = "csv" then 
			 Wscript.Echo "Export Mode : " & WSCWriteMode
		elseif WSCWriteMode = "html" then 
			Wscript.Echo "Export Mode : " & WSCWriteMode
		ElseIF WSCWriteMode = "js" then 
			Wscript.Echo "Export Mode : " & WSCWriteMode
		ElseIF WSCWriteMode = "full" then 
			Wscript.Echo "Export Mode : CSV"
			Wscript.Echo "Export Mode : HTML"
			Wscript.Echo "Export Mode : JS"		
		Else 
			WSCWriteMode = "full"   
			Wscript.Echo "Export Mode : CSV"
			Wscript.Echo "Export Mode : HTML"
			Wscript.Echo "Export Mode : JS"	 
	End If 
	
	If WSCMode = "full" Then
			Wscript.Echo "Scan Mode : [ " & WSCMode & " ] "
	ElseIf WSCMode = "perf" Then 
			Wscript.Echo "Scan Mode : [ " & WSCMode & " ] "
	ElseIf WSCMode = "config" Then 
			Wscript.Echo "Scan Mode : [ " & WSCMode & " ] "
	   Else 
			Call  DisplayHelp 
	End IF
	
	Wscript.Echo "Server Name: " & BaseComputerName
	Wscript.Echo "Reports Path: " & StrFilePathToWrite
	Wscript.Echo "Debug Level : " & DebugLevel
   
	
	
	
	
End Function

' Date and string manipulaitons -----------------------------------------------

Function GetDate 
	'-------------------------------------------------------------------------------
	' GetDate - Used to get actual system date ( in friendly human readable format ) 
	' Parameters:
	'	
	' Returns:
	'   vbdate string format.
	'-------------------------------------------------------------------------------
	Call StackTraceBuilder("DateFormatParse","ADD") 
	Call StackTraceBuilder("GetDate","ADD")
	Dim FileDate
	FileDate = FormatDateTime(now)
	FileDate = Replace(FileDate,"/","_")
	FileDate = Replace(FileDate," ","-")
	FileDate = Replace(FileDate,":","-")
	StrFileDate = FileDate
	StrDate = FormatDateTime(now)
	Call Debug("File Date   : StrFileDate","3") 
	Call StackTraceBuilder("GetDate","REMOVE")
End Function 

Function DateFormatParse ( byval DateString )
	'-------------------------------------------------------------------------------
	' DateFormatParse - Used to convert may kind of string ito vbdate format
	' Parameters:
	'	DateString - String to vonvert
	'
	' Returns:
	'   vbdate string format.
	'-------------------------------------------------------------------------------
	Call StackTraceBuilder("DateFormatParse","ADD") 
	Dim Ndatejj
	Dim Ndatemm
	Dim Ndateaaaa
	if not Datestring = ""  then 
		if mid(DateString,5,1) = "/" then 
			Call Debug("String Date test  for [ " & DateString & "]  : this is a date"  ,"6") 
			DateFormatParse = DateString ' Convertit en type Date. 
		Else 
			Call Debug("String Date test  for [ " & DateString & "]  : this is not a date"  ,"6") 
			Ndatejj =  mid(DateString,7,len(DateString)) 
			Ndatemm =  mid(DateString,5,2) 
			Ndateaaaa =  mid(DateString,1,4) 
			DateFormatParse = Ndateaaaa  & "/" & Ndatemm & "/" & Ndatejj
		end if
	end if 		
	Call StackTraceBuilder("DateFormatParse","Remove") 
end Function 

' Help and debug functions ----------------------------------------------------

Sub DisplayHelp
	'-------------------------------------------------------------------------------
	' DisplayHelp - Used to display script help information.
	' Parameters:
	'
	' Returns:
	'   how to message.
	'-------------------------------------------------------------------------------
	Wscript.Echo "___________________________________________________________________________" 
	Wscript.Echo "|                       WSC.VBS HELPER                                    |"
	Wscript.Echo "|_________________________________________________________________________|" 
	Wscript.Echo "|server     : Server to scan defaul is localhost ( or . )                 |" 
	Wscript.Echo "|debug      : Debug level defaul  is 0                                    |"  
	Wscript.Echo "|reportPath : where to writ reports files                                 |"  
	Wscript.Echo "|reporttype : Select wetween : full , csv , jr , html                     |"
	WScript.Echo "|mode       : Select between : full ( perf ans config) , config , perf    |"
	WScript.Echo "___________________________________________________________________________" 
	wscript.quit(1)
End Sub 

Sub CheckDebugState 
	'-------------------------------------------------------------------------------
	' CheckDebugState - Used to display curent debug level
	' Parameters:
	'
	' Returns:
	'   console Output
	'-------------------------------------------------------------------------------
	Call Debug("Debug Level 1 : Activated ", "1" )
	Call Debug("Debug Level 2 : Activated ", "2" )
	Call Debug("Debug Level 3 : Activated ", "3" )
	Call Debug("Debug Level 4 : Activated ", "4" )
	Call StackTraceBuilder("CheckDebugState","ADD")  
	Call Debug("Debug Level 5 : Activated ", "5" )
	Call Debug("Debug Level 6 : Activated ", "6" )
	Call Debug("Debug Level 7 : Activated ", "7" )
	Call Debug("Debug Level 8 : Activated ", "8" )
	Call Debug("Debug Level 9 : Activated ", "9" )
	Call StackTraceBuilder("CheckDebugState","REMOVE") 
end Sub 

Sub DebugInfo ( Byval MessageToPrint ) 
	'-------------------------------------------------------------------------------
	' DebugInfo - Used to prite generic messages
	' Parameters:
	'   MessageToPrint 	- String to print 
	' 
	' Returns:
	'   console Output.
	'-------------------------------------------------------------------------------
	Wscript.Echo "-----------------------------------"
	Wscript.Echo MessageToPrint
	Wscript.Echo "-----------------------------------"
	MessageToPrint  = "" 
End sub 

Sub Debug( Byval MessageToPrint, byval DebugLevelAsk )
	'-------------------------------------------------------------------------------
	' Debug - Used to display script exection information.
	' Parameters:
	'   MessageToPrint 	- Information string to display
	'	DebugLevelAsk	- Message debug level	
	'
	' Returns:
	'   console message.
	'-------------------------------------------------------------------------------
	  
		
	 if DebugLevelAsk <= DebugLevel then 
		Wscript.Echo "-----------------------------------"
		Wscript.Echo " Level : " & DebugLevelAsk   
		Wscript.Echo " " & MessageToPrint
		Wscript.Echo "-----------------------------------"
	 End If 
	 
	 MessageToPrint  = "" 
	 DebugLevelAsk = ""
end sub 


Sub StackTraceBuilder( byval FunctionName, Byval Mode ) 
	'-------------------------------------------------------------------------------
	' StackTraceBuilder - Used to build StackTrace dcebug helper.
	' Parameters:
	'   FunctionName 	- Function Name to write 
	'   Mode     		- Add or remove memeber of the stack
	'
	' Returns:
	'   Execution StackTrace.
	'-------------------------------------------------------------------------------
	if Mode = "ADD" then 
		DbgPath = DbgPath & "\" &  FunctionName 
		Call Debug("Entering Function PATH  : [ " & DbgPath & " ] ","5") 
	Elseif Mode = "REMOVE" then 
		DbgPath = Replace(DbgPath,"\" &  FunctionName,"") 
		Call Debug("Exiting : [ " & FunctionName & " ] to Function PATH  : [ " & DbgPath & " ] ","5")
	Else 
		Call Debug("Error on STACK builer at :[ " & DbgPath & " ] ","5")
	End If 
End Sub 

' Files and folders funtions --------------------------------------------------

Function CreateFolder(sFolderPath)
	'-------------------------------------------------------------------------------
	' CreateFolder - Used to create a folder.
	' Parameters:
	'   sFolderPath - The directory to create
	'
	' Returns:
	'   A New directory.
	'-------------------------------------------------------------------------------
	Call StackTraceBuilder("CreateFolder","ADD")
	Call Debug("Check folder  : sFolderPath","3") 
	Dim oFSO
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If Not oFSO.FolderExists(sFolderPath) Then
		oFSO.CreateFolder sFolderPath
		Call Debug("Folder " & sFolderPath & " : Created","3") 
	else 
		Call Debug("Folder " & sFolderPath & " : OK","3") 
	End If
	Call StackTraceBuilder("CreateFolder","REMOVE")
End Function 

Function CreateFile(sFilePath)
	'-------------------------------------------------------------------------------
	' WriteFileText - Used to write an item to a file in a folder.
	' Parameters:
	'   sFile - The file to write
	'
	' Returns:
	'   A string containing the content of the file.
	'-------------------------------------------------------------------------------
		Call StackTraceBuilder("CreateFile","ADD")
		Call Debug("Writing File : " & sFilePath ,"4")
		Set objFSO = Nothing
		Set objTextFile = Nothing
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objTextFile = objFSO.CreateTextFile(sFilePath, True)
	   
End Function

Function FileClose()
	'-------------------------------------------------------------------------------
	' WriteFileText - Used to write an item to a file in a folder.
	' Parameters:
	'   sFile - The file to write
	'
	' Returns:
	'   A string containing the content of the file.
	'-------------------------------------------------------------------------------
	Call StackTraceBuilder("FileClose","ADD")
	objTextFile.Close
	Call StackTraceBuilder("FileClose","REMOVE")
End Function

Function WriteLine(sText)
	'-------------------------------------------------------------------------------
	' WriteLine - Used to write an item to a file in a folder.
	' Returns:
	' Write a line to a text File
	'-------------------------------------------------------------------------------
		Call StackTraceBuilder("WriteLine","ADD")
		Call Debug("Writing Value : " & sText ,"4")
		' Write a line.
		 
		objTextFile.Write (sText & vbcrlf)
	   	Call StackTraceBuilder("WriteLine","REMOVE")
End Function

Function ReadFile(sFilePath)
	'-------------------------------------------------------------------------------
	' ReadFile - Used to read items from a file in a folder.
	' Returns:
	' File Lines
	'-------------------------------------------------------------------------------
		Call StackTraceBuilder("ReadFile","ADD")
		Set objFile = objFS.OpenTextFile(sFilePath)
		Do Until objFile.AtEndOfStream
			strLine= objFile.ReadLine
			Wscript.Echo strLine
		Loop
		objFile.Close
		'objFile = nothing 
		Call StackTraceBuilder("ReadFile","REMOVE")
End Function

Function WriteHeaderCSV ( sFilePath , Value )
	'-------------------------------------------------------------------------------
	' WriteHeaderCSV - Used to write CSV file header .
	' Parameters:
	'   sFilePath - The file to write
	'   Value     - header to write in file 
	'
	' Returns:
	'   A text in file.
	'-------------------------------------------------------------------------------
	Call StackTraceBuilder("WriteHeaderCSV","ADD")
	Call Debug("Writing File : " & SfilePath ,"3")
	Call Debug("Writing Value : " & Value ,"4")
	Call CreateFile(sFilePath) 
	Call WriteLine(Value)
	Call StackTraceBuilder("WriteHeaderCSV","REMOVE")
End Function

Sub BuildLogsDirs
	StrFilePathToWriteCSV  		= StrFilePathToWrite 		& "\CSV"
	StrFilePathToWriteHTML  	= StrFilePathToWrite 		& "\HTML"
	StrFilePathToWriteJS    	= StrFilePathToWrite 		& "\JS"
	StrFilePathToWriteCSVCmp	= StrFilePathToWriteCSV 	& "\" 		& BaseComputerName
	StrFilePathToWriteHTMLCmp	= StrFilePathToWriteHTML 	& "\"		& BaseComputerName
	StrFilePathToWriteJSCmp		= StrFilePathToWriteJS 		& "\" 		& BaseComputerName
	Call StackTraceBuilder("BuildLogsDirs","ADD") 
	Call Debug("Create Directory : [ " & StrFilePathToWrite 		& " ] " & "			: End","2")
	Call CreateFolder(StrFilePathToWrite)
	Call Debug("Create Directory : [ " & StrFilePathToWriteCSV 		& " ] " & "			: End","2")
	Call CreateFolder(StrFilePathToWriteCSV)
	Call Debug("Create Directory : [ " & StrFilePathToWriteHTML 	& " ] " & "			: End","2")
	Call CreateFolder(StrFilePathToWriteHTML)
	Call Debug("Create Directory : [ " & StrFilePathToWriteJS 		& " ] " & "			: End","2")
	Call CreateFolder(StrFilePathToWriteJS)
	Call Debug("Create Directory : [ " & StrFilePathToWriteCSVCmp 	& " ] " & "			: End","2")
	Call CreateFolder(StrFilePathToWriteCSVCmp)
	Call Debug("Create Directory : [ " & StrFilePathToWriteHTMLCmp 	& " ] " & "			: End","2")
	Call CreateFolder(StrFilePathToWriteHTMLCmp)
	Call Debug("Create Directory : [ " & StrFilePathToWriteJSCmp 	& " ] " & "			: End","2")
	Call CreateFolder(StrFilePathToWriteJS)
	Call StackTraceBuilder("BuildLogsDirs","REMOVE") 
End Sub



Dim HTML_TableStartLineLight
Dim HTML_TableStartLineDark
Dim HTML_TableStartLineHeader


Dim HTMLTableStart 		
Dim HTMLStartStartLine 	
Dim HTMLEndCell 			
Dim HTMLStartCell			
Dim HTMLLineEnd 		
Dim HTMLTableEnd 		
	
DIM	 HTML_PageAddLineH1S
DIM	 HTML_PageAddLineH1E
DIM	 HTML_PageAddLineH2S
DIM	 HTML_PageAddLineH2E
DIM	 HTML_PageAddLineH3S
DIM	 HTML_PageAddLineH3E

 
 Dim HTML_PageStS 
 Dim HTML_PageStE 
 Dim HTML_PageAddLineS
 Dim HTML_PageAddLineE 
 Dim HTML_TableStart 
 Dim HTML_TableStartLine 
 Dim HTML_TableEndCell 
 Dim HTML_TableStartCell
 Dim HTML_TableLineEnd 
 Dim HTML_TableEnd 
 Dim HTML_PageEnd 

' HTML vars 
 HTML_PageStS 				= chr(60)	& "!DOCTYPE html"	& chr(62)	& vbcrlf		& Chr(9)			& chr(60)	& "html" 		& chr(62)		& vbcrlf	& Chr(9)	& Chr(9)	&	chr(60)	& "head" 	& chr(62) 	& vbcrlf	& Chr(9)			& Chr(9)	& Chr(9)	& chr(60)	& "meta charset=" 	& chr(34) 	& "utf-8"	& chr(34)	& " /" 		& chr(62)	& vbcrlf	& Chr(9)	& Chr(9)	& Chr(9)	& chr(60)	& "title" 	& chr(62) 
 HTML_PageStE 				= chr(60)	& "/title"			& chr(62)	& vbcrlf		& Chr(9)			& Chr(9)	& chr(60)		& "/head" 		& chr(62)	& vbcrlf	& Chr(9)	&	chr(60)	& "body" 	& chr(62)	& vbcrlf
 HTML_PageEnd 				= chr(60)	& "/body"			& chr(62) 	& vbcrlf		& chr(60) 			& "/html"  	& chr(62)
 HTML_PageAddLineS			= chr(60)	& "BR"				& chr(62) 	 
 HTML_PageAddLineE			= chr(60) 	& "/BR"  			& chr(62)
 HTML_PageAddLineH1S		= chr(60)	& "H1"				& chr(62) 	 
 HTML_PageAddLineH1E		= chr(60) 	& "/H1"  			& chr(62)
 HTML_PageAddLineH2S		= chr(60)	& "H2 style=" 		& chr(34)	& "text-decoration: underline;" 	& chr(34)	& " "			& chr(62) 	 
 HTML_PageAddLineH2E		= chr(60) 	& "/H2"  			& chr(62)
 HTML_PageAddLineH3S		= chr(60)	& "H3 style=" 		& chr(34)	& "text-decoration: underline;" 	& chr(34)	& " "			& chr(62) 	 
 HTML_PageAddLineH3E		= chr(60) 	& "/H3"  			& chr(62)
 
 
 HTML_TableStart 			= chr(60)	& "table style="	& chr(34)	& "text-align: left; width: 900px;"	& chr(34) 	& " border="	& chr(34) 		& "1"		& chr(34) 	& " cellpadding="		& chr(34) 	& "1"		& chr(34) 	& " cellspacing="	& chr(34) 	& "1"		& chr(34) 	& " " 	& chr(62)	& " "		& vbcrlf	& "" 		& chr(60)	& "tbody" 	& chr(62)	& " "		& vbcrlf
 HTML_TableStartLine 		= Chr(9)	& chr(60)			& "tr"		& chr(62)		& vbcrlf			& Chr(9) 	& chr(60)		& "td style="	& chr(34) 	& "vertical-align: top;"			& chr(34)	& " " 		& chr(62)	& " " 
 
 HTML_TableEndCell 			= chr(60)	& "/td" 			& chr(62)	& vbcrlf
 HTML_TableStartCell		= Chr(9)	& Chr(9)			& chr(60)	& "td style="	& chr(34)			& "vertical-align: top;"	& chr(34) 		& " " 		& chr(62)	& " " 
 HTML_TableLineEnd 			= chr(60)	& "/tr"				& chr(62)	& " "
 HTML_TableEnd 				= chr(60)	& "/tbody"			& chr(62)	& " "			& vbcrlf			& chr(60)	& "/table"		& chr(62)		& vbcrlf


HTML_TableStartLineHeader  	= Chr(9)	& chr(60)			& "tr "	&	"style=" & chr(34) & "vertical-align: top; background-color: rgb(153, 153, 153);" & chr(34) & " " &  chr(62)		& vbcrlf			& Chr(9) 	& chr(60)		& "td style="	& chr(34) 	& "vertical-align: top;"			& chr(34)	& " " 		& chr(62)	& " " 
HTML_TableStartLineDark  	= Chr(9)	& chr(60)			& "tr "	&	"style=" & chr(34) & "vertical-align: top; font-family: monospace; font-weight: bold; background-color: rgb(102, 102, 102);" & chr(34) & " " & chr(62)		& vbcrlf			& Chr(9) 	& chr(60)		& "td style="	& chr(34) 	& "vertical-align: top;"			& chr(34)	& " " 		& chr(62)	& " " 
HTML_TableStartLineLight	= Chr(9)	& chr(60)			& "tr "	&	"style=" & chr(34) & "vertical-align: top; background-color: rgb(204, 204, 204);" & chr(34)	& " " & chr(62)		& vbcrlf			& Chr(9) 	& chr(60)		& "td style="	& chr(34) 	& "vertical-align: top;"			& chr(34)	& " " 		& chr(62)	& " " 

'   ----------------------------   End HTML vars

	
'	------------------------------ HTML Section -------------------------------------


Sub AddLineHtmlFile( byval Line , byval typeLine  ) 
     If typeLine = "header1" then 
		Call WriteLine ( HTML_PageAddLineH1S		& Line	&  HTML_PageAddLineH1E )
	 ElseIF typeLine = "header2" then 
	    Call WriteLine ( HTML_PageAddLineH2S		& Line	&  HTML_PageAddLineH2E )
	 ElseIF typeLine = "header3" then 
		Call WriteLine ( HTML_PageAddLineH3S		& Line	&  HTML_PageAddLineH3E )
     elseif typeLine = "dark" then
		Call WriteLine ( HTML_TableStartLineDark	& Line	&  HTML_PageAddLineE )
	 elseif   typeLine = "light" then 
	 	Call WriteLine ( HTML_TableStartLineLight	& Line	&  HTML_PageAddLineE )
	 else
		Call WriteLine ( HTML_PageAddLineS			& Line	&  HTML_PageAddLineE )
	 End If 
End Sub 

 
Sub InitHtmlFile( byval PageTitle ,  Byval File ) 
Call StackTraceBuilder("InitHtmlFile","ADD")

	Call CreateFile(File)
    Call WriteLine ( HTML_PageStS & PageTitle &  HTML_PageStE )
Call StackTraceBuilder("InitHtmlFile","REMOVE")
End Sub 

Sub EndHtmlPage
Call StackTraceBuilder("EndHtmlPage","ADD")
   Call WriteLine(HTML_PageEnd) 
   Call FileClose
  Call StackTraceBuilder("EndHtmlPage","REMOVE")
End Sub 	

 

Sub WriteTableLine ( byval  StrLinePart ,  LineNumber  ) 
	Call StackTraceBuilder("WriteTableLine","ADD")
	Dim SubCounter 
	Dim TblLine 
	Dim StrLinePartItem
	Dim UsedHEader
	
	if LineNumber = 0 then 
		UsedHEader = HTML_TableStartLineDark
    elseif LineNumber Mod 2 = 0  then 	
		UsedHEader = HTML_TableStartLineLight
	Else 
		UsedHEader = HTML_TableStartLineHeader
	End If 
	Subcounter = 0
	StrLinePart = Split(StrLinePart,";") 
	For Each StrLinePartItem in   StrLinePart
        If Subcounter = 2 then 
		  Call  WriteLine(UsedHEader & StrLinePartItem & HTML_TableEndCell & vbcrlf    ) 
		elseif not Subcounter < 2 then 
		   	WriteLine(HTML_TableStartCell &  StrLinePartItem & HTML_TableEndCell & vbcrlf) 
		End If
			 Subcounter = Subcounter + 1 
	next 
	 TblLine = TblLine & HTML_TableLineEnd
	Call  WriteLine(TblLine)
		Call StackTraceBuilder("WriteTableLine","REMOVE")
end Sub


 
 
' HTML convert function -------------------------------------------------------



Function ConvertCSVToHTML(byval sFilePath)
	'-------------------------------------------------------------------------------
	' ReadFile - Used to read items from a file in a folder and convert them to HTMLtable.
	' Returns:
	' File Lines
	'-------------------------------------------------------------------------------
	Call StackTraceBuilder("ConvertCSVToHTML","ADD")
	Dim i
	Dim strNextLine
	Dim objFSOR
	Dim objFile
	Set objFSOR = CreateObject("Scripting.FileSystemObject")
	Set objDictionary = CreateObject("Scripting.Dictionary")

	Const ForReading = 1
    wscript.echo "File " & sFilePath
	Set objFile = objFSOR.OpenTextFile (sFilePath, ForReading)
	i = 0
	Do Until objFile.AtEndOfStream
		strNextLine = objFile.Readline
		If strNextLine <> "" Then
			objDictionary.Add i, strNextLine
		End If
		i = i + 1
	Loop
	objFile.Close

	'Then you can iterate it like this

	Call LineCSVtoHTML(objDictionary)
			'objFile = nothing 
	Call StackTraceBuilder("ConvertCSVToHTML","REMOVE")
End Function


sub LineCSVtoHTML( byval objDictionary ) 
	Call StackTraceBuilder("LineCSVtoHTML","ADD")
	Dim StrLine
	Dim Counter 
	Counter = 0 
 	Call WriteLine (HTML_TableStart)
	For Each strLine in objDictionary.Items
	    Call WriteTableLine(StrLine, counter) 
		Counter = Counter + 1 
	Next
	Call  WriteLine(HTML_TableEnd)
	Call StackTraceBuilder("LineCSVtoHTML","REMOVE")
End sub 

Sub MakeTable ( Byval FileNameExpt , Byval FileNAmeToexport , Byval TableTitle, Byval SubtableTitle  ) 
	Call InitHtmlFile(TableTitle & " "  &  BaseComputerName , FileNameExpt )
	Call AddLineHtmlFile(SubtableTitle & " " & BaseComputerName , "header3")
	Call ConvertCSVToHTML(FileNAmeToexport)
	Call EndHtmlPage
End Sub
	
Dim colItems
Dim WMICSVReportFile
Dim WMIRoot

'"\root\CIMV2"

Sub GetWMIObj( byval StrComputerName, byval Wmiquery , WMIRoot  ) 
	'-------------------------------------------------------------------------------
	' GetBootConfiguration - Used to gather WMI object from Win32_BootConfiguration
	' Parameters:
	'   StrComputerName 	- IP or name  
	'	ExportMode			- export to csv HTML or js mode. ( options are : CVS , HTML , JS , FULL)
	' Returns:
	'   Execution Computername as tring.
	'-------------------------------------------------------------------------------
	'Call StackTraceBuilder("GetBootConfiguration","Add")
	'Call Debug("Get Wmi informations from : [ BootConfiguration scan instance] : Start","2") 
	Dim objWMIService
	Set objWMIService = GetObject("winmgmts:\\" & StrComputerName & "\root\Cimv2" ) 
	Set colItems = objWMIService.ExecQuery( Wmiquery,,48) 
End Sub

Function  GetComputerName( byval StrComputerName ) 
	'-------------------------------------------------------------------------------
	' GetComputerName - Used to get real computer displayname even if IP adresse is used for querying WMI angine.
	' Parameters:
	'   StrComputerName 	- IP or name  
	'
	' Returns:
	'   Execution Computername as tring.
	'-------------------------------------------------------------------------------
	Call StackTraceBuilder("GetComputerName","Add")
	Dim Result
	Dim objWMIService
	Dim objItem
	Set objWMIService = GetObject("winmgmts:\\" & StrComputerName & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery(  "SELECT * FROM Win32_ComputerSystem",,48) 
	For Each objItem in colItems 
		Call DebugInfo("Get computer name from WMI [ Win32_ComputerSystem ] instance" )
			Wscript.Echo "Name: " & objItem.Name
			Result = objItem.Name
	Next 
	GetComputerName = Result
	Call StackTraceBuilder("GetComputerName","REMOVE")
end Function 

Sub WMIObjwin32_ip4persistedroutetableDisplay( byval WmiItemCol)
	'-------------------------------------------------------------------------------
	' WMIObjwin32_ip4persistedroutetableDisplay - Used to gather WMI object from Win32_win32_ip4persistedroutetable
	' Parameters:
	'	StrComputerName 	- IP or name  
	'	ExportMode			- export to csv HTML or js mode. ( options are : CVS , HTML , JS , FULL)
	' Returns:
	'   Execution Computername as tring.
	'-------------------------------------------------------------------------------
	Call StackTraceBuilder("WMIObjwin32_ip4persistedroutetableDisplay","Add")

	Dim result
	Dim objItem
	Call Debug("Get Wmi informations from : [ BootConfiguration scan instance] : Start","2")


	Dim Header
	Header = "ComputerName;Date;description;destination;installdate;mask;metric1;name;nexthop;status"
	Call WriteHeaderCSV(StrFilePathToWriteCSVCmp & "\" & WMICSVReportFile & "-" & StrFileDate & ".csv", Header)
	For Each objItem in WmiItemCol
				Result = BaseComputerName & ";" & StrDate
				Result = Result & ";" & objItem.description
				Result = Result & ";" & objItem.destination
				Result = Result & ";" & objItem.installdate
				Result = Result & ";" & objItem.mask
				Result = Result & ";" & objItem.metric1
				Result = Result & ";" & objItem.name
				Result = Result & ";" & objItem.nexthop
				Result = Result & ";" & objItem.status
				Call  WriteLine(Result)
				Result = ""
			Next
	Call StackTraceBuilder("win32_ip4persistedroutetable","REMOVE")

	End Sub

Sub Main
	Call StackTraceBuilder("Main","ADD")
	Call ProcessArgs
	Call DebugInfo("Starting")
	Call CheckDebugState
	Call GetDate
	
	WMICSVReportFile =  "win32_ip4persistedroutetable"
	BaseComputerName = GetComputerName(BaseComputerName)
	Call BuildLogsDirs
	Call GetWMIObj(BaseComputerName, " SELECT * FROM win32_ip4persistedroutetable", "\root\cimv2")
	call WMIObjwin32_ip4persistedroutetableDisplay(colItems)
	Call MakeTable( StrFilePathToWriteHTMLCmp & "\" & WMICSVReportFile & "-" & StrFileDate & ".HTML", StrFilePathToWriteCSVCmp & "\" & WMICSVReportFile & "-" & StrFileDate & ".csv","win32_ip4persistedroutetable Repport","win32_ip4persistedroutetable Repport")
	Call StackTraceBuilder("Main","REMOVE")
End Sub  

 Call Main

