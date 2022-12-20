<!--#include file="inc/Common.asp"-->
<!--#include file="inc/Func.asp"-->
<%
uniqid = request("uniqid")
act = "add"
act_label = "新增資料"
if uniqid>"" then
  act = "edit"
  act_label = "修改資料"
  sqlstr = " select uniqid,lawid,lawtitle,sortid,ison " & vbcrlf 
  sqlstr = sqlstr & " ,convert(varchar,createdate,111)+' '+convert(varchar,createdate,108) createdate " & vbcrlf 
  sqlstr = sqlstr & " ,url " & vbcrlf 
  sqlstr = sqlstr & " from weblaw " & vbcrlf 
  sqlstr = sqlstr & " where uniqid="&uniqid&" " & vbcrlf 
  sqlstr = sqlstr & " order by sortid " & vbcrlf 
  set rs = objconn.execute(sqlstr)
  if not rs.eof then
    uniqid = trim(rs(0))
    lawid = trim(rs(1))
    lawtitle = trim(rs(2))
    sortid = trim(rs(3))
    ison = trim(rs(4))
    url = trim(rs(6))
  end if
  rs.close
else
  lawid = ""
  sortid = 1
  sqlstr = "select max(sortid) from weblaw where lawcatego=3"
  set rs = objconn.execute(sqlstr)
  if not rs.eof then
    if trim(rs(0))>"" then
      sortid = cdbl(rs(0)) + 1
    end if
  end if
  rs.close
end if


response.write "<script>var lawid='"&lawid&"';</script>"
%>

<form id="myform" onsubmit="return false;" enctype="multipart/form-data">
  <input type="hidden" name="userid" value="<%=session("userid")%>">
  <input type="hidden" name="act" id="act" value="<%=act%>">
  <input type="hidden" name="uniqid" value="<%=uniqid%>">
  <input type="hidden" name="lawid" id="lawid" value="<%=lawid%>">
  <input type="hidden" name="lawcatego" value="3">

  <div class="form-group row">
    <label for="title" class="col-sm-2 col-form-label" style="text-align: right;">標題：</label>
    <div class="col-sm-5"><input type="text" class="form-control" name="title" id="title" value="<%=lawtitle%>"></div>
  </div>


  <div class="form-group row">
    <label for="sortid" class="col-sm-2 col-form-label" style="text-align: right;">排序編號：</label>
    <div class="col-sm-1"><input type="number" class="form-control" name="sortid" id="sortid" value="<%=sortid%>"></div>
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

  <div class="form-group row">
    <label for="url" class="col-sm-2 col-form-label" style="text-align: right;">連結：</label>
    <div class="col-sm-10"><input class="form-control" name="url" id="url" value="<%=url%>"></div>
  </div>

  <button type="submit" class="btn btn-primary" onclick="dosubmit();"><%=act_label%></button>
</form>

  
<script type="text/javascript">

$(function(){
  if(lawid==""){
    lawid = uuidv4();
    $('#lawid').val(lawid);
  }
});

function dosubmit(){


  if($('#title').val().length==""){
    alert("請輸入標題!!");
    return false;
  }

  var form = new FormData(document.getElementById('myform'));

  $.ajax({
    url: 'api/json_weblaw3_do.asp',
    cache: false,
    contentType: false,
    processData: false,
    //mimeType: 'multipart/form-data',
    data: form,     //data只能指定單一物件                 
    type: 'POST',
    success: function(json){
      if(json.status=="0000"){
        alert("資料處理完成");
        window.location = "?m=weblaw3"
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