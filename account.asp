<!--#include file="inc/Common.asp"-->
<!--#include file="inc/Func.asp"-->
<%
cnt = 0
sqlstr = " select uniqid,catego,convert(varchar,postdate,111) postdate,title,ison,attachfile1,attachname1 " & vbcrlf 
sqlstr = sqlstr & " ,convert(varchar,createdate,111)+' '+convert(varchar,createdate,108) createdate " & vbcrlf 
sqlstr = sqlstr & " from accountbbs " & vbcrlf 
if request("catego")>"" then
sqlstr = sqlstr & " where catego='"&request("catego")&"' " & vbcrlf   
end if
sqlstr = sqlstr & " order by createdate desc " & vbcrlf 
set rs= objconn.execute(sqlstr)
if not rs.eof then
  data_ary = rs.getrows()
  cnt = ubound(data_ary,2)+1
end if
rs.close
%>
<div>
  <div >公告類別：<select id="catego" onchange="window.location='?m=account&catego='+this.value">
          <option value="">全部</option>
          <option value="預算" <%if request("catego")="預算" then response.write "selected"%>>預算</option>
          <option value="決算" <%if request("catego")="決算" then response.write "selected"%>>決算</option>
          <option value="會計月報" <%if request("catego")="會計月報" then response.write "selected"%>>會計月報</option>
        </select>中目前有 <font color="red" style="font-weight: bold;"><%=cnt%></font> 項資料.</div>
  <div style="text-align:right;"><button type="button" class="btn btn-danger" onclick="del();"><i class="far fa-trash-alt"></i> 刪除</button></div>
  <table class="table table-striped footable" width="100%">
    <thead>
      <tr>
        <th data-classes="ftHeader" width="2%" data-sortable="false"><input class="chkbox" type="checkbox" onclick="checkall(this.checked);"></th>
        <th data-classes="ftHeader" width="8%" data-sortable="false">公告類別</th>
        <th data-classes="ftHeader" width="15%" data-sortable="false">公告日期</th>
        <th data-classes="ftHeader" width="33%">主旨</th>
        <th data-classes="ftHeader" width="20%">附件</th>
        <th data-classes="ftHeader" width="8%">狀態</th>
        <th data-classes="ftHeader" width="15%">建檔日期</th>
      </tr>
    </thead>
    <tbody>
    <%
    if isarray(data_ary) then
      for rows=0 to ubound(data_ary,2)
        %>
        <tr>
          <td data-classes="ftRow"><input class="chkbox rowchk" value="<%=data_ary(0,rows)%>" type="checkbox"></td>
          <td data-classes="ftRow"><%=data_ary(1,rows)%></td>
          <td data-classes="ftRow"><%=data_ary(2,rows)%></a></td>
          <td data-classes="ftRow"><a href="?p=account_add&uniqid=<%=data_ary(0,rows)%>"><%=data_ary(3,rows)%></a></td>
          <td data-classes="ftRow"><a href="accounts/<%=data_ary(5,rows)%>" target="_blank"><%=data_ary(6,rows)%></a></td>
          <td data-classes="ftRow"><%if trim(data_ary(4,rows))="Y" then%>上架<%else%><span style="color:#ff0000;">下架</span><%end if%></td>
          <td data-classes="ftRow"><%=data_ary(7,rows)%></td>
        </tr>
        <%
      next
    end if
    %>  
    </tbody>
  </table>
</div>
<form id="actionform" style="display:none;"></form>

<script type="text/javascript">
$(function()
{
  $('.footable').footable(ft_option);


});  

function checkall(chkstat){
  $('.rowchk').prop("checked",chkstat);
}

function del(){
  $('#actionform').empty();
  $('.rowchk:checked').each(function(idx,item){
      $('#actionform').append('<input type="text" name="uniqid" value="'+$(item).val()+'">');
  });
  var data = $('#actionform').serialize();
  if(data==""){
    alert("請勾選項目進行刪除作業!!");
    return false;
  }else{
    if(window.confirm("您確定要進行刪除?")){
      
      $.when( postAPI('api/json_account_del_do.asp',data) ).done(function(json){
        //console.log(json);
        if(json.status=="0000"){
          window.location = "?p=account";
        }else{
          alert(json.status_desc);         
        }
      });

    }
  }
}
</script>