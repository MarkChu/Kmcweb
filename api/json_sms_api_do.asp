<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<%Response.ContentType= "application/json"%>
<!--#include file="json_common.asp"-->
<!--#include file="../inc/common.asp"-->
<!--#include file="../inc/Func.asp"-->
<!--#include file="../inc/JSON_2.0.4.asp"-->
<%
status = "0000"
group_id = request("group_id")

DIM Rs 
Set Rs=Server.CreateObject("ADODB.Recordset")
Set Rs1=Server.CreateObject("ADODB.Recordset")
 
 
SQL = "SELECT * FROM SMS_Result WHERE Group_ID = '"& Group_ID &"'"
Rs1.Open SQL,Objconn,3,2
I = 0

y = rs1.recordcount

allcnt = 0
okcnt = 0
ngcnt = 0

do while not Rs1.EOF
   Target = trim(Rs1("Cell_NO"))
   if (Target <> "") or (not isnull(Target)) then
		allcnt = allcnt + 1

		'Set ObjConvStr = server.CreateObject ("otastrconv.strconv")
		'strBody = objconvstr.str2Unicode(trim(Rs1("Description")))
		'set ObConvStr = nothing
		strBodyAry = split(str2Unicode(trim(Rs1("Description"))),",")
		strBody = ""
		strLength = 0
		if IsArray(strBodyAry) then
			for i=0 to ubound(strBodyAry)
				if i<=69 then
					strBody = strBody & strBodyAry(i)
				end if
				'strBody = strBody & strBodyAry(i)
			next
			strLength = (ubound(strBodyAry)+1)*2
		end if


	    str = "[MSISDN]" &  vbcrlf
        str = str & "List="	& ConvTEL(target) & vbcrlf
        str = str & "[MESSAGE]" & vbcrlf
        str = str & "Binary=" & strBody & vbcrlf
        str = str & "Length=" & strLength & vbcrlf
		str = str & "[SETUP]" & vbcrlf
		str = str & "DCS=UCS2" & vbcrlf
		str = str & "SplitText=yes" & vbcrlf
        str = str & "[END]"
        'response.write str
		if sendsms(str) = true then
			'session("strSendOK") = "OK"
			Rs1("Sent_Result")=1
			Rs1.Update

			okcnt = okcnt + 1
		else
			'session("strSendOK") = "False"
			Rs1("Sent_Result")=-10
			Rs1.Update   

			ngcnt = ngcnt + 1
		end if
   
   else
      Rs1("Sent_Result")=-10
      Rs1.Update   
   end if   
   Rs1.MoveNext 
loop 


FUNCTION str2Unicode(str)
    Dim objScript
    Set objScript = Server.CreateObject("ScriptControl")
    objScript.Language = "JavaScript"
	sqlstr = " if (!String.prototype.padStart) { " & vbcrlf 
	sqlstr = sqlstr & "     String.prototype.padStart = function padStart(targetLength,padString) { " & vbcrlf 
	sqlstr = sqlstr & "         targetLength = targetLength>>0; //truncate if number or convert non-number to 0; " & vbcrlf 
	sqlstr = sqlstr & "         padString = String((typeof padString !== 'undefined' ? padString : ' ')); " & vbcrlf 
	sqlstr = sqlstr & "         if (this.length > targetLength) { " & vbcrlf 
	sqlstr = sqlstr & "             return String(this); " & vbcrlf 
	sqlstr = sqlstr & "         } " & vbcrlf 
	sqlstr = sqlstr & "         else { " & vbcrlf 
	sqlstr = sqlstr & "             targetLength = targetLength-this.length; " & vbcrlf 
	sqlstr = sqlstr & "             if (targetLength > padString.length) { " & vbcrlf 
	sqlstr = sqlstr & "                 padString += padString.repeat(targetLength/padString.length); //append to original to ensure we are longer than needed " & vbcrlf 
	sqlstr = sqlstr & "             } " & vbcrlf 
	sqlstr = sqlstr & "             return padString.slice(0,targetLength) + String(this); " & vbcrlf 
	sqlstr = sqlstr & "         } " & vbcrlf 
	sqlstr = sqlstr & "     }; " & vbcrlf 
	sqlstr = sqlstr & " } " & vbcrlf 
	sqlstr = sqlstr & " if (!String.prototype.repeat) { " & vbcrlf 
	sqlstr = sqlstr & "   String.prototype.repeat = function(count) { " & vbcrlf 
	sqlstr = sqlstr & "     'use strict'; " & vbcrlf 
	sqlstr = sqlstr & "     if (this == null) { " & vbcrlf 
	sqlstr = sqlstr & "       throw new TypeError('can\'t convert ' + this + ' to object'); " & vbcrlf 
	sqlstr = sqlstr & "     } " & vbcrlf 
	sqlstr = sqlstr & "     var str = '' + this; " & vbcrlf 
	sqlstr = sqlstr & "     count = +count; " & vbcrlf 
	sqlstr = sqlstr & "     if (count != count) { " & vbcrlf 
	sqlstr = sqlstr & "       count = 0; " & vbcrlf 
	sqlstr = sqlstr & "     } " & vbcrlf 
	sqlstr = sqlstr & "     if (count < 0) { " & vbcrlf 
	sqlstr = sqlstr & "       throw new RangeError('repeat count must be non-negative'); " & vbcrlf 
	sqlstr = sqlstr & "     } " & vbcrlf 
	sqlstr = sqlstr & "     if (count == Infinity) { " & vbcrlf 
	sqlstr = sqlstr & "       throw new RangeError('repeat count must be less than infinity'); " & vbcrlf 
	sqlstr = sqlstr & "     } " & vbcrlf 
	sqlstr = sqlstr & "     count = Math.floor(count); " & vbcrlf 
	sqlstr = sqlstr & "     if (str.length == 0 || count == 0) { " & vbcrlf 
	sqlstr = sqlstr & "       return ''; " & vbcrlf 
	sqlstr = sqlstr & "     } " & vbcrlf 
	sqlstr = sqlstr & "     // Ensuring count is a 31-bit integer allows us to heavily optimize the " & vbcrlf 
	sqlstr = sqlstr & "     // main part. But anyway, most current (August 2014) browsers can't handle " & vbcrlf 
	sqlstr = sqlstr & "     // strings 1 << 28 chars or longer, so: " & vbcrlf 
	sqlstr = sqlstr & "     if (str.length * count >= 1 << 28) { " & vbcrlf 
	sqlstr = sqlstr & "       throw new RangeError('repeat count must not overflow maximum string size'); " & vbcrlf 
	sqlstr = sqlstr & "     } " & vbcrlf 
	sqlstr = sqlstr & "     var maxCount = str.length * count; " & vbcrlf 
	sqlstr = sqlstr & "     count = Math.floor(Math.log(count) / Math.log(2)); " & vbcrlf 
	sqlstr = sqlstr & "     while (count) { " & vbcrlf 
	sqlstr = sqlstr & "        str += str; " & vbcrlf 
	sqlstr = sqlstr & "        count--; " & vbcrlf 
	sqlstr = sqlstr & "     } " & vbcrlf 
	sqlstr = sqlstr & "     str += str.substring(0, maxCount - str.length); " & vbcrlf 
	sqlstr = sqlstr & "     return str; " & vbcrlf 
	sqlstr = sqlstr & "   } " & vbcrlf 
	sqlstr = sqlstr & " } " & vbcrlf 
	objScript.AddCode sqlstr
    objScript.AddCode "function dec2hex(dec, padding){var n = parseInt(dec,10);return n.toString(16).padStart(padding, '0');}"
    objScript.AddCode "function utf8StringToUtf16String(str){var utf16 = [];for(i=0;i<str.length;i++){utf16.push(dec2hex(str.charCodeAt(i),4))}return utf16.join(',');}"
    str2Unicode = objScript.Eval("utf8StringToUtf16String(""" & str & """);")
    Set objScript = NOTHING
END FUNCTION


function sendSMS(strSMSMessage)

	strAction="http://messaging.ota.com.tw/kmcgov71350/kmcgov71350.sms" 	
	strUsername="kmcgov71350"
	strPassWord="tBUq4pv0"
	Set xmlHTTP = Server.CreateObject("MicroSoft.XMLHTTP")
	xmlHTTP.open "POST",straction ,false ,strusername,strpassword
	XmlHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlHTTP.send  (strSMSMessage) '傳送的簡訊字串
	
	strRetval = XmlHTTP.ResponseText 
	'response.write strRetval
	set xmlHttp = nothing
	'sendSMS = strRetval
	IF instr(strRetval,"ORDERID=")>0 Then
		sendSMS = true
	ELSE
		response.write strSMSMessage
		response.write strRetval
		sendSMS = false
	End IF
End Function


function ConvTEL(strPhoneno)
	arrPhoneno= Split(strphoneno,", ")
	for i = lbound(arrPhoneno) to ubound(arrPhoneno)
		if len(trim(arrPhoneno(i)))<>0 then
			if mid(arrphoneno(i),1,1)="+" or mid(arrphoneno(i),1,1)<>"0" then
				if mid(arrphoneno(i),1,1)="+" then
					strTmpPhone = strTmpPhone + "," + trim(arrphoneno(i))
				else
					strTmpPhone = strTmpPhone + ",+" + trim(arrphoneno(i))
				end if	
			else
				strTmpPhone = strTmpPhone + ",+886" + mid(trim(arrphoneno(i)),2,len(trim(arrphoneno(i)))-1)
			end if
		end if	
	next
	if len(strtmpphone)<> 0 then
		strTmpPhone = Mid(strTmpPhone, 2, Len(strTmpPhone) - 1)
	end if	
	
	ConvTel=strTmpPhone
end function


Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Base64Decode(ByVal vCode)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.text = vCode
    Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Private Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = adTypeText
  BinaryStream.CharSet = "utf-8"
  BinaryStream.Open
  BinaryStream.WriteText Text
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary
  BinaryStream.Position = 0
  Stream_StringToBinary = BinaryStream.Read
  Set BinaryStream = Nothing
End Function

Private Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = adTypeBinary
  BinaryStream.Open
  BinaryStream.Write Binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText
  BinaryStream.CharSet = "us-ascii"
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function


Dim jsa
Set jsa = jsObject()

jsa("status") = status
jsa("status_desc") = status_desc
jsa("allcnt") = allcnt
jsa("okcnt") = okcnt
jsa("ngcnt") = ngcnt

jsa.Flush
response.end%>