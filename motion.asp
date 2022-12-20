<!--#include file="inc/Common.asp"-->
<!--#include file="inc/Func.asp"-->
<form id="myform" onsubmit="return false;">
  	<div class="form-group row">
	    <label for="sm1_expkd" class="col-sm-2 col-form-label" style="text-align: right;">屆別：</label>
	    <div class="col-sm-2"><select class="form-control" id="sm1_expkd" name="sm1_expkd" data-sm1_expkd="" onchange="dogettitle();">
			<option value="">全部</option>
	    	<%
			sqlstr = "select distinct(sm1_expkd) sm1_expkd from prtms_project WHERE sys_opsts<>'D' order by 1 desc"
			call MakeCombo(sqlstr,"sm1_expkd","sm1_expkd","")		
	    	%></select></div>
	    <label for="sm1_seqkd" class="col-sm-2 col-form-label" style="text-align: right;">會別：</label>
	    <div class="col-sm-2"><select class="form-control" id="sm1_seqkd" name="sm1_seqkd" data-sm1_seqkd="" onchange="dogettitle();">
	    	<option value="">全部</option>
	    	<option value="1">定期會</option>
	    	<option value="2">臨時會</option>
	    	<option value="3">審查覆議案臨時會</option>
	    	</select></div>
	    <label for="sm1_seqno" class="col-sm-2 col-form-label" style="text-align: right;">會次：</label>
	    <div class="col-sm-2"><select class="form-control" id="sm1_seqno" name="sm1_seqno" data-sm1_title="" onchange="dofinal();"><option value="">全部</option></select></div>
  	</div>

	<div class="form-group row">
		<label for="sm1_sugkd" class="col-sm-2 control-label" style="text-align: right;">議案別：</label>
		<div class="col-sm-2 control-input"><select class="form-control" id="sm1_sugkd" name="sm1_sugkd">
			<option value="">全部</option>
	    	<option value="1">提案</option>
	    	<option value="2">動議案</option>
		</select></div>
		<label for="sm1_modkd" class="col-sm-2 control-label" style="text-align: right;">提案別：</label>
		<div class="col-sm-2 control-input"><select class="form-control" id="sm1_modkd" name="sm1_modkd">
			<option value="">全部</option>
	    	<option value="1">議決案</option>
	    	<option value="2">討論案</option>
	    	<option value="3">考察</option>			
		</select></div>
		<label for="sm1_chkkd" class="col-sm-2 control-label" style="text-align: right;">類別：</label>
		<div class="col-sm-2 control-input"><select class="form-control" id="sm1_chkkd" name="sm1_chkkd">
			<option value="">全部</option>
	    	<option value="1">民政類</option>
	    	<option value="2">財政類</option>
	    	<option value="3">建設類</option>			
			<option value="4">教育類</option>
		</select></div>
	</div>

	<div class="form-group row">
		<label for="search_content" class="col-sm-2 control-label" style="text-align: right;">關鍵字：</label>
		<div class="col-sm-2 control-input"><input type="text" class="form-control" maxlength="20" id="search_content" name="search_content" value=""></div>
		<label for="sm1_1C" class="col-sm-2 control-label" style="text-align: right;">提案人： </label>
		<div class="col-sm-2 control-input"><input type="text" class="form-control" maxlength="20" id="sm1_1C" name="sm1_1C" value=""></div>
		<label for="sm1_2C" class="col-sm-2 control-label" style="text-align: right;">連署人： </label>
		<div class="col-sm-2 control-input"><input type="text" class="form-control" maxlength="20" id="sm1_2C" name="sm1_2C" value=""></div>
	</div>


  	<button type="submit" class="btn btn-primary" onclick="dosearch();">查詢</button>
	<input type="hidden" name="sm1_title" id="sm1_title" value="">
</form>


<div id="list" class="form-group row">
	<div class="col-sm-6" >
	<div id="tablecnt" style="display:none;">查詢結果共有 <span class="cnt">0</span> 筆。</div>
	<table class="table table-striped footable" id="dtltb">
		<thead>
			<tr class="ft-head">
				<th data-formatter="show_expkd" data-sortable="false">屆別</th>
				<th data-name="sm1_seqno" data-sort-value="mySortValue">會次</th>
				<th data-formatter="show_seqkd" data-sortable="false">會別</th>
				<th data-formatter="show_sugkd" data-sortable="false">提/動</th>
				<th data-formatter="show_modkd" data-sortable="false">提/討</th>
				<th data-formatter="show_chkkd" data-sortable="false">類別</th>
				<th data-name="sm1_sugno" data-sort-value="mySortValue">案號</th>
				<th data-name="sm1_regyy">年度</th>
				<th data-formatter="show_regno" data-sortable="false">文號</th>
				<!--visible column start-->
				<th data-visible="false" data-name="uniqid" data-filterable="false"></th>
				<th data-visible="false" data-name="sm1_ln" data-filterable="false"></th>
				<th data-visible="false" data-name="sm1_expkd" data-filterable="false"></th>
				<th data-visible="false" data-name="sm1_seqkd" data-filterable="false"></th>
				<th data-visible="false" data-name="sm1_sugkd" data-filterable="false"></th>
				<th data-visible="false" data-name="sm1_modkd" data-filterable="false"></th>
				<th data-visible="false" data-name="sm1_chkkd" data-filterable="false"></th>
				<th data-visible="false" data-name="sm1_regno" data-filterable="false"></th>
			</tr>
		</thead>
		<tbody>
		</tbody>
	</table>
	</div>
	<div class="col-sm-6" id="dtl" style="display:none;">
	</div>
</div>


<script>
var ft;
$(function(){
	ft = FooTable.init('#dtltb',ft_option);

	dogettitle(); 
});	


function dosearch(){
	//alert('aa');
	$('#dtl').hide();
	var param = $('#myform').serialize();
	//console.log(param);
	$.ajax({
		url: "api/json_motion.asp",
		dataType : "json",
		data: "act=getlist&"+param,
		type: "GET",
		error: function() { 
			alert("error");
			$.unblockUI(); 
		},
		beforeSend:function(){
			$.blockUI();
		},
		success: function(json){
			$.unblockUI();
			var $data = json.data;
			ft.rows.load($data);
			$('#list').show();
			$('#tablecnt').find('.cnt').html($data.length);
			$('#tablecnt').show();
		}
  	}); 
}


function dogettitle(){
  var sm1_expkd = $('#sm1_expkd').val();
  var sm1_seqkd = $('#sm1_seqkd').val();

  if($('#sm1_expkd').data("sm1_expkd")!=sm1_expkd || $('#sm1_seqkd').data("sm1_seqkd")!=sm1_seqkd){
   	$('#sm1_title').data("sm1_title","");
  }

  var sm1_title = $('#sm1_seqno').data("sm1_title");
  $('#sm1_title').val('');
  $('#sm1_seqno').empty().append('<option value="">全部</option>');
  $.ajax({
      url: "api/json_motion.asp",
      dataType : "json",
      data: "act=getseq&sm1_expkd="+sm1_expkd+"&sm1_seqkd="+sm1_seqkd+'&sm1_title='+sm1_title,
      type: "GET",
      error: function() { alert("error") },
      beforeSend:function(){},
      success: function(json){
      	if(json.data.length>0){
      		$.each(json.data,function(idx,item){
      			var opt = '<option value="'+item.sm1_seqno+'" data-sm1_title="'+(item.sm1_title>""?item.sm1_title:"")+'" '+(item.is_selected?"selected":"")+'>'+(item.sm1_title>""?item.sm1_title:item.sm1_seqno)+'</option>'
      			$('#sm1_seqno').append(opt);
      		});
      	}
      }
  }); 

}  

function dofinal(){
  var seqno = $('#sm1_seqno').val();
  $('#sm1_seqno option:selected').each(function(idx,item){
    if($(this).val()!=""){
      $('#sm1_title').val($(this).data("sm1_title"));
    }else{
      $('#sm1_title').val('');
    }
  });
}

function show_expkd(value, options, rowData){
	var sm1_expkd=rowData.sm1_expkd;
	var sm1_ln=rowData.sm1_ln;
	var rtn = $('<a/>',{'href':'javascript:void(0);'}).html(sm1_expkd).click(function(){
				showdtl(this, sm1_ln);
			});
	return rtn;
}

function show_seqkd(value, options, rowData){
	var sm1_seqkd = trim(rowData.sm1_seqkd);
	switch(sm1_seqkd){
		case "1":
			return "定期會";
			break;
		case "2":
			return "臨時會";
			break;
		case "3":
			return "審查覆議案臨時會";
			break;
	}
}

function show_sugkd(value, options, rowData){
	var sm1_sugkd = trim(rowData.sm1_sugkd);
	switch(sm1_sugkd){
		case "1":
			return "提案";
			break;
		case "2":
			return "動議案";
			break;
	}
}

function show_modkd(value, options, rowData){
	var sm1_modkd = trim(rowData.sm1_modkd);
	switch(sm1_modkd){
		case "1":
			return "議決案";
			break;
		case "2":
			return "討論案";
			break;
		case "3":
			return "考察";
			break;
	}
}

function show_chkkd(value, options, rowData){
	var sm1_chkkd = trim(rowData.sm1_chkkd);
	switch(sm1_chkkd){
		case "1":
			return "民政";
			break;
		case "2":
			return "財政";
			break;
		case "3":
			return "建設";
			break;
		case "4":
			return "教育";
			break;	
	}
}

function show_regno(value, options, rowData){
	var sm1_regno = trim(rowData.sm1_regno);
	var sm1_ln =  rowData.sm1_ln;
	return sm1_regno;
}

function show_3C(value, options, rowData){
	return showdtlcontent(rowData.sm1_3C);
}

function show_chk(value, options, rowData){
	var html = '<input type="checkbox" class="chkbox" name="sm1_ln[]" value="'+rowData.sm1_ln+'">';
	return html;
}

function showdtl(linkobj,sm1_ln){
	var $tbody = $(linkobj).closest('tbody');
	$tbody.find('tr').css('background','');
	var $row = $(linkobj).closest('tr');
	$row.css('background','#fcffc4');
	var param = "sm1_ln="+sm1_ln;
	//console.log(param);
	$.ajax({
		url: "api/json_motion.asp",
		dataType : "json",
		data: "act=getdtl&"+param,
		type: "GET",
		error: function() { alert("error") },
		beforeSend:function(){},
		success: function(json){
			if(json.data.length>0){
				var val = json.data[0];
				var $sugkd_title = "";
				var $show6c = true;
		        var $show4c_empty = true;
		        switch (trim(val.sm1_sugkd)) {
		          case '1':
		            $sugkd_title = "提案人";
		            $show4c_empty = false;
		            break;
		          case '2':
		            $sugkd_title = "動議人";
		            $show6c = false;
		            break;          
		        }

		        var $show1c = true;
		        var $show2c = true;
		        var $show4c = true;
		        var $show5c = true;

		        if(trim(val.sm1_modkd)=="2"||trim(val.sm1_modkd)=="3")
		        {
		            $show1c = false;
		            $show2c = false;
		            $show4c = false;
		            $show5c = false;          
		        }



				var html  ='<table class="detailtable" style="margin-top:0px;">';
				if($show1c){
				html +='<tr>';
				html +='<td class="dtl-header">'+$sugkd_title+'</td>';
				html +='<td class="dtl-body">'+showdtlcontent(val.sm1_1C)+'</td>';
				html +='</tr>';	
				}
				if($show2c){
				html +='<tr>';
				html +='<td class="dtl-header">連署人</td>';
				html +='<td class="dtl-body">'+showdtlcontent(val.sm1_2C)+'</td>';
				html +='</tr>';
				}
				html +='<tr>';
				html +='<td class="dtl-header">案由</td>';
				html +='<td class="dtl-body">'+showdtlcontent(val.sm1_3C)+'</td>';
				html +='</tr>';
				html +='<tr>';
				html +='<td class="dtl-header">議決</td>';
				html +='<td class="dtl-body">'+showdtlcontent(val.sm1_7C)+'</td>';
				html +='</tr>';
				if($show4c_empty==false){
		          if(trim(val.sm1_4C)==""||val.sm1_4C==null){
		            $show4c = false;
		          }  
		        }
		        if($show4c){
				html +='<tr>';
				html +='<td class="dtl-header">說明</td>';
				html +='<td class="dtl-body">'+showdtlcontent(val.sm1_4C)+'</td>';
				html +='</tr>';	        	
		        }
		        if($show5c){
				html +='<tr>';
				html +='<td class="dtl-header">辦法</td>';
				html +='<td class="dtl-body">'+showdtlcontent(val.sm1_5C)+'</td>';
				html +='</tr>';
				}
				if($show6c){
				html +='<tr>';
				html +='<td class="dtl-header">審查意見</td>';
				html +='<td class="dtl-body">'+showdtlcontent(val.sm1_6C)+'</td>';
				html +='</tr>';
				}
				html +='<tr>';
				html +='<td class="dtl-header">處理情形</td>';
				html +='<td class="dtl-body">'+showdtlcontent(val.sm1_8C)+'</td>';
				html +='</tr>';
				html +='</table>';
				$('#dtl').empty().append(html);
				$('#dtl').show();	
			}
		}
  	}); 
}



</script>	
