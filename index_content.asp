<div class="ct_head">
  <div class="ct_head_title"><i class="fas fa-angle-double-right"></i><%=pagetitle%></div>
</div>
<%
if IsArray(submenu) then
%>
<div class="card-header">
  <ul class="nav nav-tabs card-header-tabs">
    <%
    for sn=0 to ubound(submenu,2)
      title = submenu(0,sn)
      url = submenu(1,sn)
      isactive = false
      if url=page then
        isactive  = true
      end if
      %>
      <li class="nav-item">
        <a class="nav-link <%if isactive then response.write "active"%>" href="?p=<%=url%>"><%=title%></a>
      </li>
      <%
    next
    %>
  </ul>
</div>
<%end if%>
<div class="card-body">
<%
if page>"" then
  if FunCheckFile(page&".asp") then
    Server.Execute(page&".asp")
  end if
end if
%>
</div>