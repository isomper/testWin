Dim $objHTTP
Dim $strEnvelope
Dim $strReturn
Dim $objReturn
Dim $dblTax
Dim $strQuery
Dim $value

$value = InputBox("Testing", "Enter your new value here.", "20170724161ABC")

; Initialize COM error handler
$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")

$objHTTP = ObjCreate("Microsoft.XMLHTTP")
$objReturn = ObjCreate("Msxml2.DOMdocument.3.0")

; Create the SOAP Envelope
#cs
$strEnvelope = "<soap:envelope xmlns:soap=""urn:schemas-xmlsoap-org:soap.v1"">" & _
"<soap:header></soap:header>" & _
"<soap:body>" & _
"<m:getsalestax xmlns:m=""urn:myserver/soap:TaxCalculator"">" & _
"<salestotal>"&$value&"</salestotal>" & _
"</m:getsalestax>" & _
"</soap:body>" & _
"</soap:envelope>"
#ce
;"<soap:header></soap:header>" & _

$strEnvelope = "<S:Envelope xmlns:S=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
"<S:Body>" & _
"<ns2:ssoTokenGet xmlns:ns2=""http://webservice.subforeign.fort.sbr.com/"">" & _
"<authPwd>"&$value&"</authPwd>" & _
"<resId>147747419009616614383899</resId>" & _
"<userAccount>a</userAccount>" & _
"</ns2:ssoTokenGet>" & _
"</S:Body>" & _
"</S:Envelope>"

; Set up to post to our local server
;$objHTTP.open ("post", "http://localhost/soap.asp", False)
$objHTTP.open ("post", "http://192.168.23.234:8989/userOpen/", False)

; Set a standard SOAP/ XML header for the content-type
$objHTTP.setRequestHeader ("Content-Type", "text/xml")

; Set a header for the method to be called
;$objHTTP.setRequestHeader ("SOAPMethodName", "urn:myserver/soap:TaxCalculator#getsalestax")
$objHTTP.setRequestHeader ("SOAPMethodName", "http://webservice.subforeign.fort.sbr.com/#ssoTokenGet")

ConsoleWrite("Content of the Soap envelope : "& @CR & $strEnvelope & @CR & @CR)

; Make the SOAP call
$objHTTP.send ($strEnvelope)

; Get the return envelope
$strReturn = $objHTTP.responseText

; ConsoleWrite("Debug : "& $strReturn & @CR & @CR)

; Load the return envelope into a DOM
$objReturn.loadXML ($strReturn)

ConsoleWrite("Return of the SOAP Msg : " & @CR & $objReturn.XML & @CR & @CR)

; Query the return envelope
;$strQuery = "SOAP:Envelope/SOAP:Body/m:getsalestaxresponse/salestax"
$strQuery = "S:Envelope/S:Body/ns2:ssoTokenGetResponse/return/token"

$dblTax = $objReturn.selectSingleNode($strQuery)
$Soap = $objReturn.Text

MsgBox(0,"SOAP Response","The Sales Tax is : " & $Soap)

Func MyErrFunc()
$HexNumber=hex($oMyError.number,8)
Msgbox(0,"COM Test","We intercepted a COM Error !" & @CRLF & @CRLF & _
             "err.description is: " & @TAB & $oMyError.description & @CRLF & _
             "err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: " & @TAB & $HexNumber & @CRLF & _
             "err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
             "err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
             "err.source is: " & @TAB & $oMyError.source & @CRLF & _
             "err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext _
            )
SetError(1) ; to check for after this function returns
Endfunc