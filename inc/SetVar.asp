<%
'logout
IF Request("logout")="True" Then
	session.contents.removeall   
  	session.abandon
	response.redirect "Default.asp"   
End IF

'產生TempCaseNo
Session("CaseNoTemp")=dt2Int(DateFormat(now(),1))&Rdn12()

'Default PageClass
MainPage = "MainPage"
leftFlag = True

'PageClass
Session("Page")=Request("p")
SELECT CASE Request("p")
	CASE "Case_Main"
		Session("Area")="CASE"
		MainPage = Request("p")
	CASE "SYSTEM"
		Session("Area")="SYSTEM"
		MainPage = Request("p")
	CASE ELSE
		MainPage = Request("p")
END SELECT

%>