<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<Script>
//window.resizeTo(200,350);
</Script>
<style type='text/css'>
	A.Default {COLOR: #000000; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	A.Default:linked {COLOR: #000061; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: Underline; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	A.Default:visited {COLOR: #000061; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: Underline; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	A.Default:hover {COLOR: #4000A2; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: Underline; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	.Default {COLOR: #000000; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	A.Cool {COLOR: #C20000; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px }
	A.Cool:linked {COLOR: #E01F25; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px }
	A.Cool:visited {COLOR: #FF0000; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px }
	A.Cool:hover {COLOR: #A00000; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px }
	.Cool {COLOR: #C20000; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px }
	A.cool2 {COLOR: #002F80; FONT-FAMILY: Arial; FONT-SIZE: 12px; FONT-WEIGHT: Bold; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
	A.cool2:linked {COLOR: #002F80; FONT-FAMILY: Arial; FONT-SIZE: 12px; FONT-WEIGHT: Bold; TEXT-DECORATION: Underline; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
	A.cool2:visited {COLOR: #004080; FONT-FAMILY: Arial; FONT-SIZE: 12px; FONT-WEIGHT: Bold; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
	A.cool2:hover {COLOR: #FF0000; FONT-FAMILY: Arial; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
	.cool2 {COLOR: #002F80; FONT-FAMILY: Arial; FONT-SIZE: 12px; FONT-WEIGHT: Bold; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
</style>

<%
YY = Request("YY")
MM = Request("MM")

IF YY = "" Then
	YY = Year(Now())
End IF
IF MM = "" Then
	MM = Month(Now())
End IF

YY = Cint(YY)
MM = Cint(MM)

SelectedDate=Request("SelectedDate")
IF SelectedDate <> "" Then
	YY = Year(CDATE(SelectedDate))
	MM = Month(CDATE(SelectedDate))
	Session("SelectDate")=SelectedDate
End IF


IF MM=1 Then
	LastStr = "?Flag="+Server.URLEncode(Request("Flag"))+"&YY=" & YY -1 & "&MM=12"
ELSE
	LastStr = "?Flag="+Server.URLEncode(Request("Flag"))+"&YY=" & YY & "&MM=" & MM-1
End IF

IF MM=12 Then
	NextStr = "?Flag="+Server.URLEncode(Request("Flag"))+"&YY=" & YY +1 & "&MM=1"
ELSE
	NextStr = "?Flag="+Server.URLEncode(Request("Flag"))+"&YY=" & YY & "&MM=" & MM + 1
End IF

FirstDay = CDATE(YY & "/" & MM & "/1")
n=WeekDay(FirstDay)-1
showDate = FirstDay

%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>小月曆</title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

      <TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0 align="center"><!--InsideTable-->
<Form action="" name="theForm" method="get">
        <TR>
          <TD ALIGN="center"><!--ConnectPool-->
<TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
  <TD align="center"><A href="Calendar.asp<%=LastStr%>">&lt;&lt;</A>&nbsp;<Font class="cool2">西元<Font class="cool"> <B><%=YY%></B> </Font>年<Font class="cool"> <B><%=MM%></B> </Font>月</Font>&nbsp;<A href="Calendar.asp<%=NextStr%>">&gt;&gt;</A></TD>
</TR>
<TR>
  <TD BGCOLOR="#5f5f5f">
    <TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=1>
      <TR ALIGN="" VALIGN="">
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="14%">日</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="14%">一</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="14%">二</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="14%">三</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="14%">四</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="14%">五</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="14%">六</TD>
      </TR>
      
<%
For i=0 to 5
	IF Month(showDate) = MM Then
	%>
	<TR>
	<%
	IF i = 0 and n<>0 Then
		For a=1 to n
			%>
			<TD ALIGN="center" BGCOLOR="#f7f7f7" BORDERCOLORDARK="#f7f7f7" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default">&nbsp;</TD>		
			<%
		Next
		For b=0 to 6-n
			IF WeekDay(showDate) = 7 Then
				colorstr = " BGCOLOR=""#e0ffbf"" BORDERCOLORDARK=""#e0ffbf"" "
			ELSE
				IF WeekDay(showDate) = 1 Then
					colorstr = " BGCOLOR=""#ffe1dc"" BORDERCOLORDARK=""#ffe1dc"" " 
				ELSE
					colorStr = " BGCOLOR=""#ffffff"" BORDERCOLORDARK=""#ffffff"" "
				End IF
			End IF
			IF showDate = CDate(Year(now()) & "/" & Month(now()) & "/" & Day(now())) Then
				colorStr = " BGCOLOR=""#c2efff"" BORDERCOLORDARK=""#c2efff"" "
			End IF
			IF showDate = CDate(Session("SelectDate")) Then
				FontStr = " class=""Cool"" "
			Else
				FontStr = " class=""cool2"" "	
			End IF
			%>
			<TD ALIGN="center" <%=colorStr%> BORDERCOLORLIGHT="#5f5f5f" CLASS="Default"><A href="#" onClick="ReturnDate('<%=YY%>','<%=MM%>','<%=Day(showDate)%>');"><Font <%=FontStr%>><%=Day(showDate)%></Font></a></TD>
			<%		
			showDate = showDate + 1
		Next
	ELSE
		For c=0 to 6
			IF Month(showDate) = MM Then
				IF WeekDay(showDate) = 7 Then
					colorstr = " BGCOLOR=""#e0ffbf"" BORDERCOLORDARK=""#e0ffbf"" "
				ELSE
					IF WeekDay(showDate) = 1 Then
						colorstr = " BGCOLOR=""#ffe1dc"" BORDERCOLORDARK=""#ffe1dc"" " 
					ELSE
						colorStr = " BGCOLOR=""#ffffff"" BORDERCOLORDARK=""#ffffff"" "
					End IF
				End IF
				IF showDate = CDate(Year(now()) & "/" & Month(now()) & "/" & Day(now())) Then
					colorStr = " BGCOLOR=""#c2efff"" BORDERCOLORDARK=""#c2efff"" "
				End IF
				IF showDate = CDate(Session("SelectDate")) Then
					FontStr = " class=""Cool"" "
				Else
					FontStr = " class=""cool2"" "	
				End IF
				%>
				<TD ALIGN="center" <%=colorStr%> BORDERCOLORLIGHT="#5f5f5f" CLASS="Default"><A href="#" onClick="ReturnDate('<%=YY%>','<%=MM%>','<%=Day(showDate)%>');"><Font <%=FontStr%>><%=Day(showDate)%></Font></a></TD>
				<%		
				showDate = showDate + 1
			ELSE
				%>
				<TD ALIGN="center" BGCOLOR="#f7f7f7" BORDERCOLORDARK="#f7f7f7" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default">&nbsp;</TD>
				<%
			End IF
		Next
	End IF
	%>
	</TR>
	<%
	End IF
Next
%>	  
	</TABLE></TD>		
	</TR>
</TABLE></TD>
	</TR>
	</Form>
</TABLE>
</body>

<SCRIPT language=javascript>

function ReturnDate(nYear,nMonth,nDay)
{
   if (opener!=null)
   {
      var Name=window.name;
      var p1=Name.indexOf("EINSTAND");
      var FormName=Name.substring(0,p1);
      var p2=Name.indexOf("EINSTAND",p1+1);
      if (p2!=-1)
      {
         var p3=Name.indexOf("EINSTAND",p2+1);
         var YCtrlName=Name.substring(p1+8,p2);
         var MCtrlName=Name.substring(p2+8,p3);
         var DCtrlName=Name.substring(p3+8,Name.length);
         var ParentCtrl=eval("opener.document."+FormName+"."+YCtrlName);
         if (ParentCtrl!=null) ParentCtrl.value=""+nYear; else alert("ParentCtrl is null->"+YCtrlName);
         ParentCtrl=eval("opener.document."+FormName+"."+MCtrlName);
         if (ParentCtrl!=null) ParentCtrl.value=""+nMonth; else alert("ParentCtrl is null->"+MCtrlName);
         ParentCtrl=eval("opener.document."+FormName+"."+DCtrlName);
         if (ParentCtrl!=null) ParentCtrl.value=""+nDay; else alert("ParentCtrl is null->"+DCtrlName);
      }
      else
      {
         var CtrlName=Name.substring(p1+8,Name.length);
         var ParentCtrl=eval("opener.document."+FormName+"."+CtrlName);
         if (ParentCtrl!=null)
         {
            ParentCtrl.value=""+nYear+"/"+nMonth+"/"+nDay;
         }
         else
         {
            alert("ParentCtrl is null->"+CtrlName);
         }
      }
   }
   else
   {
      alert("opner is null");
   }
   <%
   SELECT CASE request("Flag")
   		CASE "1"
		%>
	opener.GetSVNo();
		<%
		CASE "2"
		%>
	opener.DoDT();	
		<%
		CASE ELSE
   END SELECT
   %>
   self.close();
}

</SCRIPT>

