<!--#include file="inc/Common.asp"-->
<!--#include file="inc/Func.asp"-->
<%
cnt = 0
sqlstr = " select uniqid,lawid,lawtitle,sortid,ison " & vbcrlf 
sqlstr = sqlstr & " ,convert(varchar,createdate,111)+' '+convert(varchar,createdate,108) createdate " & vbcrlf 
sqlstr = sqlstr & " from weblaw " & vbcrlf 
sqlstr = sqlstr & " where lawcatego=4 " & vbcrlf   
sqlstr = sqlstr & " order by sortid " & vbcrlf 
set rs= objconn.execute(sqlstr)
if not rs.eof then
  data_ary = rs.getrows()
  cnt = ubound(data_ary,2)+1
end if
rs.close
%>
<div>
  <div >其他法規 中目前有 <font color="red" style="font-weight: bold;"><%=cnt%></font> 項資料.</div>
  <div style="text-align:right;"><button type="button" class="btn btn-danger" onclick="del();"><i class="far fa-trash-alt"></i> 刪除</button></div>
  <table class="table table-striped footable" width="100%">
    <thead>
      <tr>
        <th data-classes="ftHeader" width="2%" data-sortable="false"><input class="chkbox" type="checkbox" onclick="checkall(this.checked);"></th>
        <th data-classes="ftHeader" width="63%">標題</th>
        <th data-classes="ftHeader" width="10%">排序</th>
        <th data-classes="ftHeader" width="10%">狀態</th>
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
          <td data-classes="ftRow"><a href="?p=weblaw4_add&uniqid=<%=data_ary(0,rows)%>"><%=data_ary(2,rows)%></a></td>
          <td data-classes="ftRow"><%=data_ary(3,rows)%></td>
          <td data-classes="ftRow"><%if trim(data_ary(4,rows))="Y" then%>上架<%else%><span style="color:#ff0000;">下架</span><%end if%></td>
          <td data-classes="ftRow"><%=data_ary(5,rows)%></td>
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
      
      $.when( postAPI('api/json_weblaw1_del_do.asp',data) ).done(function(json){
        //console.log(json);
        if(json.status=="0000"){
          window.location = "?p=weblaw4";
        }else{
          alert(json.status_desc);         
        }
      });

    }
  }
}
</script>