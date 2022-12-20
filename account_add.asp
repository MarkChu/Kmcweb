<!--#include file="inc/Common.asp"-->
<!--#include file="inc/Func.asp"-->
<%
uniqid = request("uniqid")
act = "add"
act_label = "新增資料"
if uniqid>"" then
  act = "edit"
  act_label = "修改資料"
  sqlstr = " select uniqid,catego,convert(varchar,postdate,111) postdate,title,ison,attachfile1,attachname1 " & vbcrlf 
  sqlstr = sqlstr & " ,convert(varchar,createdate,111)+' '+convert(varchar,createdate,108) createdate " & vbcrlf 
  sqlstr = sqlstr & " from accountbbs " & vbcrlf 
  sqlstr = sqlstr & " where uniqid='"&uniqid&"' " & vbcrlf 
  set rs = objconn.execute(sqlstr)
  if not rs.eof then
    uniqid = trim(rs(0))
    catego = trim(rs(1))
    postdate = trim(rs(2))
    title = trim(rs(3))
    ison = trim(rs(4))
    attachfile1 = trim(rs(5))
    attachname1 = trim(rs(6))
  end if
  rs.close
end if
%>
<form id="myform" onsubmit="return false;" enctype="multipart/form-data">
  <input type="hidden" name="userid" value="<%=session("userid")%>">
  <input type="hidden" name="act" id="act" value="<%=act%>">
  <input type="hidden" name="uniqid" value="<%=uniqid%>">

  <div class="form-group row">
    <label for="catego" class="col-sm-2 col-form-label" style="text-align: right;">公告類別：</label>
    <div class="col-sm-3">
      <select id="catego" name="catego" class="form-control">
        <option value="預算" <%if catego="預算" then response.write "selected"%>>預算</option>
        <option value="決算" <%if catego="決算" then response.write "selected"%>>決算</option>
        <option value="會計月報" <%if catego="會計月報" then response.write "selected"%>>會計月報</option>
      </select> 
    </div>
  </div>    

  <div class="form-group row">
    <label for="postdate" class="col-sm-2 col-form-label" style="text-align: right;">公告日期：</label>
    <div class="col-sm-2"><input type="text" class="form-control" name="postdate" id="postdate" value="<%=postdate%>" placeholder="YYYY/MM/DD"></div>
  </div>

  <div class="form-group row">
    <label for="title" class="col-sm-2 col-form-label" style="text-align: right;">主旨：</label>
    <div class="col-sm-5"><input type="text" class="form-control" name="title" id="title" value="<%=title%>"></div>
  </div>

  <div class="form-group row">
    <label for="file1" class="col-sm-2 col-form-label" style="text-align: right;">附件：</label>
    <div class="col-sm-5">
      <%If attachname1>"" then%>
      <a href="accounts/<%=attachfile1%>" target="_blank"><%=attachname1%></a><br>
      <%end if%>
      <input type="file" class="form-control" name="file1" id="file1" accept=".pdf">

    </div>
  </div>

  <div class="form-group row">
    <label for="ison" class="col-sm-2 col-form-label" style="text-align: right;">狀態：</label>
    <div class="col-sm-2">
      <select id="ison" name="ison" class="form-control">
        <option value="Y" <%if ison="Y" then response.write "selected"%>>上架</option>
        <option value="N" <%if ison="N" then response.write "selected"%>>下架</option>
      </select> 
    </div>
  </div>  


  <button type="submit" class="btn btn-primary" onclick="dosubmit();"><%=act_label%></button>
</form>

<script type="text/javascript">

$(function(){
  $('#postdate').datepicker(dateformatOption);


});


function dosubmit(){
  
  if($('#postdate').val().length==""){
    alert("請輸入公告日期!!");
    return false;
  }

  if($('#title').val().length==""){
    alert("請輸入主旨!!");
    return false;
  }

  if($('#file1').val().length==""&&$('#act').val()=="add"){
    alert("請選擇上傳的檔案!!");
    return false;
  }

  var form = new FormData(document.getElementById('myform'));

  $.ajax({
    url: 'api/json_account_do.asp',
    cache: false,
    contentType: false,
    processData: false,
    //mimeType: 'multipart/form-data',
    data: form,     //data只能指定單一物件                 
    type: 'POST',
    success: function(json){
      if(json.status=="0000"){
        alert("資料處理完成");
        window.location = "?m=account"
      }else{
        alert(json_err(json.status_desc));
      }
    }
  });
  

  return false;
};

function dec2hex(dec, padding){
  return parseInt(dec, 10).toString(16).padStart(padding, '0');
}

function utf8StringToUtf16Array(str) {
  var utf16 = [];
  for (var i=0, strLen=str.length; i < strLen; i++) {
    utf16.push(dec2hex(str.charCodeAt(i), 4));
  }
  return utf16;
}
  
</script>