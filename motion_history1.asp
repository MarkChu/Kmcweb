<!--#include file="inc/Common.asp"-->
<!--#include file="inc/Func.asp"-->
<form id="myform" onsubmit="return false;">
  	<div class="form-group row">
	    <label for="HN" class="col-sm-2 col-form-label" style="text-align: right;">冊別：</label>
	    <div class="col-sm-6"><select class="form-control" id="HN" name="HN">
	    	<%
			sqlstr = "select distinct(HN) HN from prtms_project_D0110 order by 1 desc"
			call MakeCombo(sqlstr,"HN","HN","")
	    	%>
		</select></div>
  	</div>
  	<button type="submit" class="btn btn-primary" onclick="dosearch();">查詢</button>
</form>


<div id="list" class="form-group row">
	<div class="col-sm-12" >
	<div id="tablecnt" style="display:none;">查詢結果共有 <span class="cnt">0</span> 頁。</div>
	<table class="table table-striped footable" id="dtltb" data-paging-size="1">
		<thead>
			<tr class="ft-head">
				<th data-formatter="show_img" data-sortable="false"></th>
				<!--visible column start-->
				<th data-visible="false" data-name="uniqid" data-filterable="false"></th>
				<th data-visible="false" data-name="sm1_img" data-filterable="false"></th>
				<th data-visible="false" data-name="sm1_page" data-filterable="false"></th>
			</tr>
		</thead>
		<tbody>
		</tbody>
	</table>
	</div>
</div>


<script>
var ft;
$(function(){
	ft = FooTable.init('#dtltb',ft_option);

});	


function dosearch(){
	//alert('aa');
	$('#dtl').hide();
	var param = $('#myform').serialize();
	//console.log(param);
	$.ajax({
		url: "api/json_motion.asp",
		dataType : "json",
		data: "act=getlistK&"+param,
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

function show_img(value, options, rowData){
	var sm1_img = trim(rowData.sm1_img);
	var sm1_page = trim(rowData.sm1_page);
	var str = '<div style="text-align:center;margin-top:10px;margin-bottom:10px;">頁次：'+sm1_page+'</div>';
	str += '<div style="text-align:center;"><img src="../DATA/TALK0110/'+sm1_img+'" style="max-width:800px;"></div>';
	return str;
}

</script>	
