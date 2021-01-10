Dim WshShell, strCurDir, xHttp, bStrm, srcAddress, dstName
	Set srcAddress = WScript.Arguments
	Set dstName = WScript.Arguments
	Set WshShell = CreateObject("WScript.Shell")
	strCurDir    = WshShell.CurrentDirectory
    Set xHttp = createobject("Microsoft.XMLHTTP")
	Set bStrm = createobject("Adodb.Stream")
	xHttp.Open "GET", srcAddress.item(0), False
	xHttp.Send
	with bStrm
		.type = 1 '//binary
		.open
		.write xHttp.responseBody
		.savetofile strCurDir & "\" & dstName.item(1), 2 '//overwrite
	end with
