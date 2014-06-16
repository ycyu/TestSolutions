Dim oSOAP

'Create an object of Soap Client
Set oSOAP = CreateObject("MSSOAP.SoapClient30")

'Initaialize the Web Service
oSOAP.mssoapinit("http://localhost/DimWeb/service.asmx?wsdl")

'Invoke the Web Service
set ss = oSOAP.InvokeDeployProduction("ECHECK", "ECHECK_CR_33", "ECHECK_CR_33_D_01")

WSCript.Echo ss.getElementsByTagName("APILog")(0).text



'set objXML = CreateObject("Microsoft.XMLDOM")
'objXML.async = False
'objXML.LoadXML(oSOAP.InvokeDeployProduction("ECHECK", "ECHECK_CR_33", "ECHECK_CR_33_D_01"))
'WScript.Echo objXML.getElementsByTagName("APILog")(0).text
'set objXML = Nothing