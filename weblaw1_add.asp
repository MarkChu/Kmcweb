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
  sqlstr = sqlstr & " ,lawcontent " & vbcrlf 
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
    lawcontent = trim(rs(6))
  end if
  rs.close
else
  lawid = ""
  sortid = 1
  sqlstr = "select max(sortid) from weblaw where lawcatego=1"
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
  <input type="hidden" name="lawcatego" value="1">
  <input type="hidden" name="detid" id="detid">
  <input type="hidden" name="isdel" id="isdel">
  <input type="hidden" name="chcatego" id="chcatego">
  <input type="hidden" name="chtitle" id="chtitle">
  <input type="hidden" name="chcontent" id="chcontent">
  <input type="hidden" name="detsortid" id="detsortid">

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
    <label for="lawcontent" class="col-sm-2 col-form-label" style="text-align: right;">內容/編修歷程：</label>
    <div class="col-sm-10"><textarea rows="5" class="form-control" name="lawcontent" id="lawcontent"><%=lawcontent%></textarea></div>
  </div>


  <div class="form-group row">
    <label class="col-sm-2 col-form-label" style="text-align: right;">條文內容：</label>
    <div class="col-sm-10">
      <button type="button" class="btn btn-primary" onclick="addrow();">新增條文</button>
      <table class="table table-striped footable" data-paging="false" id="det_list" width="100%">
        <thead>
          <tr>
            <th data-classes="ftHeader" data-sortable="false" width="5%" data-formatter="show_chk">刪除</th>
            <th data-classes="ftHeader" data-sortable="false" data-formatter="show_chcatego" width="15%">章節</th>
            <th data-classes="ftHeader" data-sortable="false" data-formatter="show_chtitle" width="10%">條目</th>
            <th data-classes="ftHeader" data-sortable="false" data-formatter="show_chcontent" width="65%">內容</th>
            <th data-classes="ftHeader" data-sortable="false" data-formatter="show_sortid" width="10%">排序</th>
            <th data-visible="false" data-name="lawid" data-filterable="false"></th>
            <th data-visible="false" data-name="isdel" data-filterable="false"></th>
            <th data-visible="false" data-name="detid" data-filterable="false"></th>
            <th data-visible="false" data-name="chcatego" data-filterable="false"></th>
            <th data-visible="false" data-name="chtitle" data-filterable="false"></th>
            <th data-visible="false" data-name="chcontent" data-filterable="false"></th>
            <th data-visible="false" data-name="detsortid" data-filterable="false"></th>
          </tr>
        </thead>
        <tbody>
        </tbody>
      </table>  
    </div>
  </div>

  <button type="submit" class="btn btn-primary" onclick="dosubmit();"><%=act_label%></button>
</form>

  
<script type="text/javascript">
var ft;

$(function(){
  ft = FooTable.init('#det_list',ft_option); 
  
  if(lawid==""){
    lawid = uuidv4();
    $('#lawid').val(lawid);
  }

  getlist();
});


function getlist(){
  var data = {
    act: "getdet",
    lawid: lawid
  };
  $.when( postAPI('api/json_weblaw1.asp',data) ).done(function(json){
    if(json.status=="0000"){
      $(json.data).each(function(idx,item){
        var values = {
          detid: item.detid,
          isdel: 'N',
          lawid: item.lawid,
          chcatego: item.chcatego,
          chtitle: item.chtitle,
          chcontent: item.chcontent,
          detsortid: item.sortid
        };
        ft.rows.add(values);
      });
    }else{
      alert(json.status_desc);         
    }
  });

}


function addrow(){
  var values = {
    detid: 0,
    isdel: 'N',
    lawid: lawid,
    chcatego: '',
    chtitle: '',
    chcontent: '',
    detsortid: 0
  };
  ft.rows.add(values);
}


function show_chk(value, options, rowData){
  var html = '<input type="hidden" class="detid" value="'+rowData.detid+'">';
  html += '<input class="isdel" type="hidden" value="'+rowData.isdel+'">';
  html += '<input type="checkbox" class="chkbox" value="'+rowData.detid+'">';
  return html;
}

function show_chcatego(value, options, rowData){
  var html = '<input type="text" class="form-control chcatego" value="'+rowData.chcatego+'">';
  return html;
}

function show_chtitle(value, options, rowData){
  var html = '<input type="text" class="form-control chtitle" value="'+rowData.chtitle+'">';
  return html;
}

function show_chcontent(value, options, rowData){
  var html = '<textarea class="form-control chcontent">'+rowData.chcontent+'</textarea>';
  return html;
}

function show_sortid(value, options, rowData){
  var html = '<input type="text" class="form-control detsortid" value="'+rowData.detsortid+'">';
  return html;
}


function genclass(clsname){
  var splitstr = '@kmc@';
  var final = '';
  $('.'+clsname).each(function(idx,item){
    var s = item.value;
    if(s!=undefined){
      s = s.replace(/,/g, '，');
      // s = s.replace(/\r\n|\n/g,'<br>');
    }else{
      s = '';
    }
    // item.value = s;
    final += splitstr + s;
  });
  if(final!=''){
    final = final.substr(splitstr.length)
  }
  return final;
}


function dosubmit(){

  $('.chkbox').each(function(idx,item){
    if(item.checked){
      $(item).closest('tr').find('.isdel').val('Y');
    }else{
      $(item).closest('tr').find('.isdel').val('N');
    }
  });
  
  $('#chcontent').val(genclass('chcontent'));
  $('#chcatego').val(genclass('chcatego'));
  $('#chtitle').val(genclass('chtitle'));
  $('#isdel').val(genclass('isdel'));
  $('#detid').val(genclass('detid'));
  $('#detsortid').val(genclass('detsortid'));


  if($('#title').val().length==""){
    alert("請輸入標題!!");
    return false;
  }

  var form = new FormData(document.getElementById('myform'));

  $.ajax({
    url: 'api/json_weblaw1_do.asp',
    cache: false,
    contentType: false,
    processData: false,
    //mimeType: 'multipart/form-data',
    data: form,     //data只能指定單一物件                 
    type: 'POST',
    success: function(json){
      if(json.status=="0000"){
        alert("資料處理完成");
        window.location = "?m=weblaw1"
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