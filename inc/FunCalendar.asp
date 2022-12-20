<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<Script>
window.resizeTo(250,350);
</Script>
<style type='text/css'>
	A.Default {COLOR: #000000; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	A.Default:linked {COLOR: #000061; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	A.Default:visited {COLOR: #000061; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	A.Default:hover {COLOR: #4000A2; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	.Default {COLOR: #000000; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 20px }
	A.Cool {COLOR: #C20000; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px; FONT-WEIGHT:bold; }
	A.Cool:linked {COLOR: #E01F25; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px; FONT-WEIGHT:bold; }
	A.Cool:visited {COLOR: #FF0000; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px; FONT-WEIGHT:bold; }
	A.Cool:hover {COLOR: #A00000; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px; FONT-WEIGHT:bold; }
	.Cool {COLOR: #C20000; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 15px; FONT-WEIGHT:bold; }
	A.cool2 {COLOR: #002F80; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; FONT-WEIGHT: Bold; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
	A.cool2:linked {COLOR: #002F80; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; FONT-WEIGHT: Bold; TEXT-DECORATION: Underline; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
	A.cool2:visited {COLOR: #004080; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; FONT-WEIGHT: Bold; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
	A.cool2:hover {COLOR: #FF0000; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
	.cool2 {COLOR: #002F80; FONT-FAMILY: Tahoma; FONT-SIZE: 12px; FONT-WEIGHT: Bold; TEXT-DECORATION: none; LETTER-SPACING: 0px; LINE-HEIGHT: 16px }
</style>

<%
YYMM = Request("YYMM")
FormNM = Request("FormNM")
FieldNM = Request("FieldNM")
DTFormat = Request("DTFormat")
FunNM = Request("FunNM")

IF YYMM &"" = "" Then
	YY = Year(Now())
	MM = Month(Now())
ELSE
	YY = Left(YYMM,4)
	MM = Right(YYMM,2)
End IF

YY = Cint(YY)
MM = Cint(MM)


SelectedDate=Request("ThisDT")
IF Request("ThisDT")&"">"" Then
	YY = Year(DecodeDTFormat(Request("ThisDT"),DTFormat))
	MM = Month(DecodeDTFormat(Request("ThisDT"),DTFormat))
	Session("SelectDate")=DecodeDTFormat(Request("ThisDT"),DTFormat)
End IF


IF MM=1 Then
	LastYYMM = YY-1&"12"
ELSE
	LastYYMM = YY&right("0"&MM-1,2)
End IF

IF MM=12 Then
	NextYYMM = YY+1&"01"
ELSE
	NextYYMM = YY&right("0"&MM+1,2)
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

      <TABLE WIDTH="204" BORDER=0 CELLPADDING=0 CELLSPACING=0 align="center" style="margin-top:5px;"><!--InsideTable-->
<Form action="FunCalendar.asp" name="theForm" method="post">
        <TR>
          <TD ALIGN="center"><!--ConnectPool-->
<TABLE WIDTH="100%" BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
  <TD align="center" class="Default"><Img src="../Images/cal_3.gif" onClick="goLastYear();" align="absmiddle" style="cursor:hand;margin-right:3px;" alt="前一年">  <Img src="../Images/cal_1.gif" onClick="goLast();" align="absmiddle" style="cursor:hand;" alt="前一個月">
  <Font class="cool2">西元<Font class="cool"> <B><%=YY%></B> </Font>年<Font class="cool"> <B><%=right("0"&MM,2)%></B> </Font>月</Font>&nbsp;
  <Img src="../Images/cal_2.gif" onClick="goNext();" align="absmiddle" style="cursor:hand;margin-right:3px;" alt="後一個月"><Img src="../Images/cal_4.gif" onClick="goNextYear();" align="absmiddle" style="cursor:hand;" alt="後一年"></TD>
  
</TR>
<TR>
  <TD BGCOLOR="#5f5f5f">
    <TABLE WIDTH="204" BORDER=0 CELLPADDING=0 CELLSPACING=1>
      <TR ALIGN="" VALIGN="">
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="28">日</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="28">一</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="28">二</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="28">三</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="28">四</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="28">五</TD>
        <TD ALIGN="center" BGCOLOR="#c0c0c0" BORDERCOLORDARK="#c0c0c0" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" WIDTH="28">六</TD>
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
			<TD ALIGN="center" BGCOLOR="#f7f7f7" BORDERCOLORDARK="#f7f7f7" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" onClick="ReturnDate('');" style="cursor:hand;">&nbsp;</TD>		
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
			<TD ALIGN="center" <%=colorStr%> BORDERCOLORLIGHT="#5f5f5f" CLASS="Default"><A href="#" onClick="ReturnDate('<%=FunDtFormat(showDate,DTFormat)%>');"><Font <%=FontStr%>><%=Day(showDate)%></Font></a></TD>
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
					colorStr = " BGCOLOR=""#ffffff"" BORDERCOLORDARK=""#ffffff"" "
				Else
					FontStr = " class=""cool2"" "	
				End IF
				%>
				<TD ALIGN="center" <%=colorStr%> BORDERCOLORLIGHT="#5f5f5f" CLASS="Default"><A href="#" onClick="ReturnDate('<%=FunDtFormat(showDate,DTFormat)%>');"><Font <%=FontStr%>><%=Day(showDate)%></Font></a></TD>
				<%		
				showDate = showDate + 1
			ELSE
				%>
				<TD ALIGN="center" BGCOLOR="#f7f7f7" BORDERCOLORDARK="#f7f7f7" BORDERCOLORLIGHT="#5f5f5f" CLASS="Default" onClick="ReturnDate('');" style="cursor:hand;">&nbsp;</TD>
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
    	<Input type="hidden" name="YYMM" value="<%=Request("YYMM")%>">
    	<Input type="hidden" name="FunNM" value="<%=Request("FunNM")%>">
        <Input type="hidden" name="FormNM" value="<%=Request("FormNM")%>">
        <Input type="hidden" name="FieldNM" value="<%=Request("FieldNM")%>">
        <Input type="hidden" name="DTFormat" value="<%=Request("DTFormat")%>">
	</Form>
</TABLE>
</body>

<SCRIPT language=javascript>


function goLast(){
	document.theForm.YYMM.value="<%=LastYYMM%>";
	document.theForm.submit();
}

function goNext(){
	document.theForm.YYMM.value="<%=NextYYMM%>";
	document.theForm.submit();
}

function goLastYear(){
	document.theForm.YYMM.value="<%=YY-1&right("0"&MM,2)%>";
	document.theForm.submit();
}

function goNextYear(){
	document.theForm.YYMM.value="<%=YY+1&right("0"&MM,2)%>";
	document.theForm.submit();
}



function ReturnDate(dtstr)
{
	opener.document.<%=FormNM%>.<%=FieldNM%>.value=dtstr;
	<%IF FunNM&"">"" Then%>
	opener.<%=FunNM%>(dtstr);
	<%END IF%>
	window.close();
}

</SCRIPT>
<%
Function FunDTFormat(fundt,fundtformatstr)
	IF fundtformatstr&""="" Then
		fundtformatstr = "YYYY/MM/DD"
	End IF
	SELECT CASE UCASE(fundtformatstr)
		CASE "YYYYMMDD"
				FunDTFormat = Year(fundt) & right("0"&Month(fundt),2) & right("0"&day(fundt),2)
		CASE "YYYY/MM/DD"
				FunDTFormat = Year(fundt) & "/" & right("0"&Month(fundt),2) & "/" & right("0"&day(fundt),2)
		CASE "YYYY-MM-DD"
				FunDTFormat = Year(fundt) & "-" & right("0"&Month(fundt),2) & "-" & right("0"&day(fundt),2)
	END SELECT
End function

Function DecodeDTFormat(dtstr,fundtformatstr)
	IF fundtformatstr&""="" Then
		fundtformatstr = "YYYY/MM/DD"
	End IF
	SELECT CASE UCASE(fundtformatstr)
		CASE "YYYYMMDD"
				DecodeDTFormat = CDATE(left(dtstr,4)&"/"&right(left(dtstr,6),2)&"/"&right(dtstr,2))
		CASE "YYYY/MM/DD"
				DecodeDTFormat = CDATE(dtstr)
		CASE "YYYY-MM-DD"
				DecodeDTFormat = CDATE(replace(dtstr,"-","/"))
	END SELECT
End function

%>
