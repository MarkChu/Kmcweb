<!--#include file="inc/Common.asp"-->
<!--#include file="inc/Func.asp"-->
<%
cnt = 0
sqlstr = " select unique_id,group_id,[description],postdate,people_prompt,people_prompt1 " & vbcrlf 
sqlstr = sqlstr & " from sms " & vbcrlf 
sqlstr = sqlstr & " where account_id='"&session("userid")&"' " & vbcrlf 
sqlstr = sqlstr & " order by postdate desc " & vbcrlf 
set rs= objconn.execute(sqlstr)
if not rs.eof then
  data_ary = rs.getrows()
  cnt = ubound(data_ary,2)
end if
rs.close
%>
<div>
  <div >您共發出 <font color="red" style="font-weight: bold;"><%=cnt%></font> 項簡訊</div>
  <div style="text-align:right;"><button type="button" class="btn btn-danger" onclick="del();"><i class="far fa-trash-alt"></i> 刪除</button></div>
  <table class="table table-striped footable" width="100%">
    <thead>
      <tr>
        <th data-classes="ftHeader" width="2%" data-sortable="false"><input class="chkbox" type="checkbox" onclick="checkall(this.checked);"></th>
        <th data-classes="ftHeader" width="30%" data-sortable="false">發送對象</th>
        <th data-classes="ftHeader" width="50%">發送內容</th>
        <th data-classes="ftHeader" width="18%">發送時間</th>
      </tr>
    </thead>
    <tbody>
    <%
    if isarray(data_ary) then
      for rows=0 to ubound(data_ary,2)
        prompt = trim(data_ary(4,rows)&"")
        if trim(data_ary(5,rows)&"")>"" then
          if prompt>"" then
            prompt = prompt &","&trim(data_ary(5,rows)&"")
          else
            prompt = trim(data_ary(5,rows)&"")
          end if
        end if
        %>
        <tr>
          <td data-classes="ftRow"><input class="chkbox rowchk" value="<%=data_ary(0,rows)%>" type="checkbox"></td>
          <td data-classes="ftRow"><%=prompt%></td>
          <td data-classes="ftRow"><%=data_ary(2,rows)%></td>
          <td data-classes="ftRow"><%=data_ary(3,rows)%></td>
        </tr>
        <%
      next
    end if
    %>  
    </tbody>
  </table>
</div>
<form id="actionform" style="display: none;"></form>

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
      
      $.when( postAPI('api/json_sms_do.asp',data) ).done(function(json){
        //console.log(json);
        if(json.status=="0000"){
          window.location = "?p=sms";
        }else{
          alert(json.status_desc);         
        }
      });

    }
  }
}
</script>