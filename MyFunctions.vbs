Dim strCurDir, xHttp, bStrm, appVersion, fileName
function checkProcess(process)
	Dim i, strComputer, FindProc
	strComputer = "."
	FindProc = process
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery _
		("Select Name from Win32_Process WHERE Name='" & FindProc & "'")
	If colProcessList.count>0 then
		wscript.echo FindProc & " is running"
		else
			wscript.echo FindProc & " is not running"
		End if
	Set objWMIService = Nothing
	Set colProcessList = Nothing
end function

function FileExists(FilePath)
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FileExists(FilePath) Then
		FileExists=CBool(1)
	Else
		FileExists=CBool(0)
	end If
end function

function checkUrl(chkurl)
	Set o = CreateObject("MSXML2.XMLHTTP")
	on error resume next
	o.open "GET", chkurl, False
	o.send
	if o.Status = 404 then checkUrl = o.Status
	on error goto 0 
end function
	
function downLoad(dldurl)
	fileName = "\" & appVersion & ".vbs"
	Set xHttp = createobject("Microsoft.XMLHTTP")
	Set bStrm = createobject("Adodb.Stream")
	xHttp.Open "GET", dldurl, False
	xHttp.Send
	with bStrm
		.type = 1 '//binary
		.open
		.write xHttp.responseBody
		.savetofile strCurDir & fileName , 2 '//overwrite
	end with 
end function
