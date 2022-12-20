<!--#include file="inc/Common.asp"-->
<!--#include file="inc/Func.asp"-->
<link rel="stylesheet" href="dist/themes/default/style.min.css" />
<script src="dist/jstree.min.js"></script>
<%

%>
<form id="myform" onsubmit="return false;">
  <input type="hidden" name="userid" value="<%=session("userid")%>">
  <div class="form-group row">
    <label for="dept" class="col-sm-2 col-form-label" style="text-align: right;">部門：</label>
    <div class="col-sm-5"><input type="text" class="form-control" id="dept" readonly value="<%=session("dept")%>"></div>
  </div>
  <div class="form-group row">
    <label for="username" class="col-sm-2 col-form-label" style="text-align: right;">發送人：</label>
    <div class="col-sm-5"><input type="text" class="form-control" id="username" readonly value="<%=session("username")%>"></div>
  </div>

  <div class="form-group row">
    <label class="col-sm-2 col-form-label" style="text-align: right;">發送對象：</label>
    <div class="col-sm-5" style="padding:0px 15px 0px 15px;">
      <div style="border:1px #cccccc solid;height:150px;overflow-y:auto; " id="emp_tree">
        <ul>
        <%
        sqlstr = "select dept_id,dept,expand_flag from vw_Dept_Login order by sort_id"
        set rs = objconn.execute(sqlstr)
        if not rs.eof then
          dept_ary = rs.getrows()
        end if
        rs.close
        if isArray(dept_ary) then
          for d_row=0 to ubound(dept_ary,2)
            %>
            <li id="0|<%=dept_ary(0,d_row)%>|<%=dept_ary(0,d_row)%>|<%=dept_ary(1,d_row)%>"><%=dept_ary(1,d_row)%>
            <%
              sqlstr = "select account_id,[Name] from vw_Members_Login where dept_id='"&dept_ary(0,d_row)&"' order by job_id desc,[Name] "
              set rs = objconn.execute(sqlstr)
              if not rs.eof then
                response.write "<ul>"
                memb_ary = rs.getrows()
                for m_row=0 to ubound(memb_ary,2)
                  response.write "<li id=""1|"&dept_ary(0,d_row)&"|"&memb_ary(0,m_row)&"|"&memb_ary(1,m_row)&""">"
                  response.write memb_ary(1,m_row)
                  response.write "</li>"
                next
                response.write "</ul>"
              end if
              rs.close
            %>  
            </li>
            <%
          next
        end if
        %>
        </ul>
      </div>
    </div>
    <div class="col-sm-5" style="padding:0px 15px 0px 15px;">
      <div style="border:1px #cccccc solid;height:150px;overflow-y:auto;padding:5px;" id="emp_sel"></div>
    </div>
  </div>

  <div class="form-group row">
    <label class="col-sm-2 col-form-label" style="text-align: right;">發送媒體：</label>
    <div class="col-sm-5" style="padding:0px 15px 0px 15px;">
      <div >
        <Input type="checkbox" class="chkbox" name="allmedia" id="allmedia_1" value="1">
        <label class="col-form-label" for="allmedia_1">全部媒體</label>
      </div>
      <div style="border:1px #cccccc solid;height:150px;overflow-y:auto; " id="media_tree">
        <ul>
        <%
        sqlstr = "select dept_id,dept,expand_flag from dept1 order by sort_id"
        set rs = objconn.execute(sqlstr)
        if not rs.eof then
          dept_ary = rs.getrows()
        end if
        rs.close
        if isArray(dept_ary) then
          for d_row=0 to ubound(dept_ary,2)
            %>
            <li id="0|<%=dept_ary(0,d_row)%>|<%=dept_ary(0,d_row)%>|<%=dept_ary(1,d_row)%>"><%=dept_ary(1,d_row)%>
            <%
              sqlstr = "select unique_id,[Name] from Members1 where dept_id='"&dept_ary(0,d_row)&"' order by [Name] "
              set rs = objconn.execute(sqlstr)
              if not rs.eof then
                response.write "<ul>"
                memb_ary = rs.getrows()
                for m_row=0 to ubound(memb_ary,2)
                  response.write "<li id=""1|"&dept_ary(0,d_row)&"|"&memb_ary(0,m_row)&"|"&memb_ary(1,m_row)&""">"
                  response.write memb_ary(1,m_row)
                  response.write "</li>"
                next
                response.write "</ul>"
              end if
              rs.close
            %>  
            </li>
            <%
          next
        end if
        %>
        </ul>
      </div>
    </div>
    <div class="col-sm-5" style="padding:0px 15px 0px 15px;">
      <div>&nbsp;</div>
      <div style="border:1px #cccccc solid;height:150px;overflow-y:auto;padding:5px;" id="media_sel"></div>
    </div>
    <label class="col-sm-2 col-form-label" style="text-align: right;"></label>
    <div class="col-sm-10" style="color:#ff0000;"><small>發送對象無登錄手機機號者,將不進行簡訊發送</small></div>
  </div>

  <div class="form-group row">
    <label for="senddate" class="col-sm-2 col-form-label" style="text-align: right;">發送時間：</label>
    <div class="col-sm-10 form-inline">
        <input name="send_y" type="text" value="<%=year(now)%>" class="form-control" style="width:65px;">&nbsp;年&nbsp;
        <select name="send_m" class="form-control" style="width:65px;">
          <%
          for i=1 to 12
            sel = ""
            if month(now)=i then
              sel = "selected"
            end if
            %>
            <option value="<%=right("0"&i,2)%>" <%=sel%>><%=right("0"&i,2)%></option>
            <%
          next
          %>
        </select>&nbsp;月&nbsp;
        <select name="send_d" class="form-control" style="width:65px;">
          <%
          for i=1 to 31
            sel = ""
            if day(now)=i then
              sel = "selected"
            end if
            %>
            <option value="<%=right("0"&i,2)%>" <%=sel%>><%=right("0"&i,2)%></option>
            <%
          next
          %>
        </select>&nbsp;日&nbsp;
        <select name="send_h" class="form-control" style="width:65px;">
          <%
          for i=0 to 23
            sel = ""
            if hour(now)=i then
              sel = "selected"
            end if
            %>
            <option value="<%=right("0"&i,2)%>" <%=sel%>><%=right("0"&i,2)%></option>
            <%
          next
          %>
        </select>&nbsp;時&nbsp;
        <select name="send_n" class="form-control" style="width:65px;">
          <%
          for i=0 to 59
            sel = ""
            if minute(now)=i then
              sel = "selected"
            end if
            %>
            <option value="<%=right("0"&i,2)%>" <%=sel%>><%=right("0"&i,2)%></option>
            <%
          next
          %>
        </select>&nbsp;分&nbsp;
        <select name="send_s" class="form-control" style="width:65px;">
          <%
          for i=0 to 59
            sel = ""
            if second(now)=i then
              sel = "selected"
            end if
            %>
            <option value="<%=right("0"&i,2)%>" <%=sel%>><%=right("0"&i,2)%></option>
            <%
          next
          %>
        </select>&nbsp;秒&nbsp;
    </div>



  </div>

  <div class="form-group row">
    <label class="col-sm-2 col-form-label" style="text-align: right;">注意事項：</label>
    <div class="col-sm-10" style="color:#ff0000;"><small>每1則簡訊內容包含中英文字、數字、符號等以70個字為限。<br>目前字數為 <span id="cnt">0</span> 字. </small></div>
  </div>
  <div class="form-group row">
    <label class="col-sm-2 col-form-label" style="text-align: right;">簡訊內容：</label>
    <div class="col-sm-10"><textarea id="sms_body" name="sms_body" rows="4" style="width:100%;" onkeyup="countcnt()" onchange="countcnt()"></textarea></div>
  </div>
  <button type="submit" class="btn btn-primary" onclick="dosubmit();">送出</button>
</form>

<script type="text/javascript">
function countcnt(){
  var cnt = 0;
  var text = $('#sms_body').val();
  var a = utf8StringToUtf16Array(text);
  cnt = a.length;
  $('#cnt').html(cnt);
}  


$(function()
{

  $("#emp_tree").on("changed.jstree", function (e, data) {
      //console.log(data.changed.selected); // newly selected
      //addtolist('emp_sel',data.changed.selected);
      changechk('emp_tree','emp_sel','topeople');
      //console.log(data.changed.deselected); // newly deselected
      //removefromlist('emp_sel',data.changed.deselected);
    }).jstree({
    "core" : {
      "themes" : { "icons" : false },
    },
    "checkbox" : {
      "keep_selected_style" : false,
      "three_state": false,
    },
    "plugins" : [ "checkbox","changed" ]
  });

  $("#media_tree").on("changed.jstree", function (e, data) {
      //console.log(data.changed.selected); // newly selected
      //addtolist('emp_sel',data.changed.selected);
      changechk('media_tree','media_sel','tomedia');
      //console.log(data.changed.deselected); // newly deselected
      //removefromlist('emp_sel',data.changed.deselected);
    }).jstree({
    "core" : {
      "themes" : { "icons" : false },
    },
    "checkbox" : {
      "keep_selected_style" : false,
      "three_state": false,
    },
    "plugins" : [ "checkbox","changed" ]
  });


});

function changechk(_source,_target,_fieldname){
  var s = $('#'+_source).jstree("get_selected");
  $('#'+_target).data("checklist",s);
  showchklist(_target,_fieldname)
}

function showchklist(_target,_fieldname){
  $('#'+_target).empty();
  var s = $('#'+_target).data("checklist");
  if(s!=undefined){
    for(r=0;r<s.length;r++){
      var opt = s[r].split('|');
      var val_str = "";
      if(opt[0]=="0"){
        val_str = opt[2]+'@';
      }else{
        val_str = opt[2]+'@'+opt[1];
      }
      var html = '<div><input type="hidden" name="'+_fieldname+'" value="'+val_str+'"><input type="hidden" name="'+_fieldname+'_txt" value="'+opt[3]+'">'+opt[3]+'</div>';
      $('#'+_target).append(html);
    }    
  }
}

function dosubmit(){
  var theform = $('#myform');
  var allmedia = $('#allmedia_1').attr("checked");
  
  if(!allmedia) {
    if(theform.find('input[name=topeople]').length==0&&theform.find('input[name=tomedia]').length==0){
      alert("請至少選擇一名發送對象!!!!");
      return false;
    }
  }

  if(theform.find('#sms_body').val()==""){
    alert("請輸入簡訊發送內容!!");
    return false;
  }

  var data = theform.serialize();
  var a = utf8StringToUtf16Array(theform.find('#sms_body').val());
  if(a.length > 70){
    alert("簡訊字數超過限制!!");
    return false;
  }

  $.when( postAPI('api/json_sms_send_do.asp',data) ).done(function(json){
      //console.log(json);
      if(json.status=="0000"){
        var group_id = json.group_id;
        var data = {
          "group_id":group_id 
        };
        $.when( postAPI('api/json_sms_api_do.asp',data) ).done(function(api_json){        
          if(api_json.status=="0000"){
            alert("簡訊發送共 "+api_json.allcnt+" 筆，成功 "+api_json.okcnt+" 筆，失敗 "+api_json.ngcnt+" 筆。 ");
            window.location = "?p=sms";
          }else{
            alert(api_json.status_desc);
          }
        });          
      }else{
        alert(json.status_desc);         
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