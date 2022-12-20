<%
' on error resume next 
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("Talks.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
%>