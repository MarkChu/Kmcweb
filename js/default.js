function dologin()
{
  var a = $('#a1').val();
  var b = $('#a2').val();
  var c = $('#a3').val();
  if(a==""||b==""||c==""){
  	alert("登入資訊請填寫完整!!");
  	return false;
  }else{
  	var data = {
  		"a": a,
  		"b": b,
  		"c": c
  	};
  	$.when( postAPI('api/json_default_do.asp',data) ).done(function(json){
      //console.log(json);
      if(json.status=="0000"){
          window.location = "index.asp";
      }else{
        alert(json.status_desc);
        reloadimg();
        $('#a3').val("");
      }
    });
  }
}

function reloadimg(){
	d = new Date();
	$('#authimg').attr("src","GetAuthCode.asp?"+d.getTime());
}
