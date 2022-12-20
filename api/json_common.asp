<%
if trim(session("userid")&"")="" then
    session.contents.removeall   
    session.abandon
    'response.redirect "default.asp"
    response.end
end if
%>