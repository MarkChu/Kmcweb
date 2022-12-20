<%
if UCASE(Request.ServerVariables("REQUEST_METHOD"))="OPTIONS" then
  response.Status="200 OK"
  response.end
end if
Response.AddHeader "Access-Control-Allow-Origin","*"
Response.AddHeader "Access-Control-Allow-Headers","*"
Response.AddHeader "Access-Control-Allow-Credentials",true
Response.AddHeader "Access-Control-Allow-Methods","PUT, POST, GET, DELETE, OPTIONS"
response.ContentType="application/json"
%>