Option Explicit

Dim objXML, objXMLHTTP, objArgs

Set objArgs = WScript.Arguments

WScript.Echo " "
WScript.Echo "=============== Start =============== "
WScript.Echo Now
WScript.Echo " "
WScript.Echo "=============== Arguments =============== "
WScript.Echo "     Product Name : " & objArgs(0)
WSCript.Echo "            CR ID : " & objArgs(1)
WScript.Echo " Dev. Baseline ID : " & objArgs(2)

Set objXML = CreateObject("Microsoft.XMLDOM")
objXML.async = False

Set objXMLHTTP = Createobject("MSXML2.XMLHTTP")
objXMLHTTP.Open "POST", "http://localhost/DimWeb/service.asmx/InvokeDeployProduction", False
objXMLHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objXMLHTTP.Send "ProdName=" & Trim(objArgs(0)) & "&CRID=" & Trim(objArgs(1)) & "&DevBaselineID=" & Trim(objArgs(2))

objXML.LoadXML(objXMLHTTP.ResponseText)
WScript.Echo " "
WScript.Echo "=============== Response =============== "
'WScript.Echo objXMLHTTP.ResponseText
WScript.Echo " Error Msg : " & objXML.GetElementsByTagName("ErrMsg")(0).Text
WScript.Echo "   API Log : " & objXML.GetElementsByTagName("APILog")(0).Text
WScript.Echo " Build Log : " & objXML.GetElementsByTagName("BuildLog")(0).Text

Set objXMLHTTP = Nothing
Set objXML = Nothing
Set objArgs = Nothing

WScript.Echo " "
WScript.Echo "=============== End =============== "
WScript.Echo Now
