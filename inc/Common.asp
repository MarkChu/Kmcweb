<%
DBConnSTR = "Provider=sqloledb;Data Source=172.16.33.161;Initial Catalog=KMC;User Id=sa;Password=1234;"
DBConnSTR = "Provider=SQLNCLI11.1;Data Source=172.16.33.207;Initial Catalog=kmcweb;User Id=sa;Password=0926666905;"
Set OBJconn = Server.CreateObject("ADODB.Connection")
OBJconn.Open DBConnSTR
%>