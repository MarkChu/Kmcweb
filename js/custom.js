//def option
var apiServer = '';

$.ajaxSetup({
    cache: false,
    error: function (x, e){
        if (x.status == 550)
            alert("550 Error Message");
        else if (x.status == "403")
            alert("403. Not Authorized");
        else if (x.status == "500")
            alert("500. Internal Server Error");
        else
            alert("Error...");
    },
    success: function (x){
        //do something global on success... 
    }
});

//object size
Object.size = function(obj) {
    var size = 0, key;
    for (key in obj) {
        if (obj.hasOwnProperty(key)) size++;
    }
    return size;
};

var dateformatOption = {
    closeText: "關閉",
    prevText: "&#x3C;上個月",
    nextText: "下個月&#x3E;",
    currentText: "今天",
    monthNames: [ "一月","二月","三月","四月","五月","六月",
    "七月","八月","九月","十月","十一月","十二月" ],
    monthNamesShort: [ "一月","二月","三月","四月","五月","六月",
    "七月","八月","九月","十月","十一月","十二月" ],
    dayNames: [ "星期日","星期一","星期二","星期三","星期四","星期五","星期六" ],
    dayNamesShort: [ "週日","週一","週二","週三","週四","週五","週六" ],
    dayNamesMin: [ "日","一","二","三","四","五","六" ],
    weekHeader: "週",
    dateFormat: "yy/mm/dd",
    firstDay: 1,
    isRTL: false,
    showMonthAfterYear: true,
    yearSuffix: "年",
};

var ft_option = {
    "sorting": {"enabled":true},
    "filtering": {"enabled":false},
    "paging": {"enabled":true ,"size":10},
    "state": {"enabled":false},
    "empty": "",
    "calculateWidthOverride": function() {
        return { width: $(window).width() };
    },
    "on": {
            "preinit.ft.table":function(e, ft){

            },
            "init.ft.table": function(e, ft){
                // bind to the plugin initialize event to do something
            },
            "predraw.ft.table": function(e, ft) {
               
            },
            "draw.ft.table": function(e, ft) {

            },
            "postdraw.ft.table": function(e, ft) {

            },
        },
}



function postAPI(_url, _formdata){
    
  return $.ajax({
    type: "POST",
    url: _url.replace('.asp&','.asp?'),
    data : _formdata,
    dataType: "json",
    error: function (x, e){
        if (x.status == 550)
            alert("550 Error Message");
        else if (x.status == "403")
            alert("403. Not Authorized");
        else if (x.status == "500")
            alert("500. Internal Server Error");
        else
            alert("Error...");
        $.unblockUI();  
    },
    beforeSend:function(){
      $.blockUI();
    },
    complete:function(){
      
    },  
    success: function(jData) {
      $.unblockUI();
    }
  });
}


function getAPI(_url){
  return $.ajax({
    type: "GET",
    url: _url,
    dataType: "json",
    error: function (x, e){
        if (x.status == 550)
            alert("550 Error Message");
        else if (x.status == "403")
            alert("403. Not Authorized");
        else if (x.status == "500")
            alert("500. Internal Server Error");
        else
            alert("Error...");
    },
    beforeSend:function(){
    },
    complete:function(){
    },  
    success: function(jData) {
    }
  });
}



function isDate(dtstr, format){
    var isValid = true;
    try{
      $.datepicker.parseDate(format,dtstr, null);
    }
    catch(error){
      isValid = false;
    }

    return isValid;
}




function IsNumber(num,min,max){
    var isVaild = true;
    if($.isNumeric(num)){
      if(min != null && $.isNumeric(min)){
        if(parseFloat(num)<parseFloat(min)){
          isVaild = false;
        }
      }
      if(max != null && $.isNumeric(max)){
        if(parseFloat(num)>parseFloat(max)){
          isVaild = false;
        }
      }
    }else{
      isVaild = false;
    }
   
    return isVaild;
 
}


function replaceall(fnsource,fnIn,fnOut){
    var a_str = fnsource;
    var call_x = eval('a_str.replace(/'+fnIn+'/g, "'+fnOut+'")');
    return call_x;
}


function addCommas(nStr)
{
    nStr += '';
    x = nStr.split('.');
    x1 = x[0];
    x2 = x.length > 1 ? '.' + x[1] : '';
    var rgx = /(\d+)(\d{3})/;
    while (rgx.test(x1)) {
        x1 = x1.replace(rgx, '$1' + ',' + '$2');
    }
    return x1 + x2;
}


function trim(x) {
    if(x==null){
        return '';
    }else{
        return x.replace(/^\s+|\s+$/gm,'');    
    }
}

function json_err(_arr){
    var $rtn = "";
    var s = Object.size(_arr);
    for(i=1;i<=s;i++){
      $rtn += _arr[i]+"\r";
    }
    return $rtn;
}

function showdtlcontent(jsondata){
    var s = jsondata;
    if(s==null){
        s = '';
    }else{
        s = s.replace(/\r\n/g,"<br />");
    }
    return s
}

function initDatePicker(jq_selector,addopt){
    var options;
    if(addopt!=null){
        options = $.extend({},dateformatOption,addopt);
    }else{
        options = dateformatOption;
    }
    //console.log(options);
    jq_selector.css({
        "width":"120px",
        "text-align":"center",
    })
    .attr({
        "autocomplete":"off",
        "placeholder":"YYYY-MM-DD",
        "maxlength":"10",
    })
    .datepicker(options);    
}


function autoHideFilterForm(fnformid){
    $('#'+fnformid).hide();
    var btn = $('<button/>',{'class':'btn expendbtn','style':'float:right;margin-right:10px;'}).html('展開查詢條件').click(function(e){
        e.preventDefault();
        $('#'+fnformid).show();
        $(this).remove();
    });
    $('#'+fnformid).parent().find('.expendbtn').remove();
    $('#'+fnformid).after(btn);    
}

function autoReNotify(){
    setTimeout(function() {
        oneReNotify();
        autoReNotify();
    }, 60*1000);
}

function oneReNotify(){
    $.when( getAPI('notify_json.php') ).done(function(json){
        if(json.status="0000"){
            var result = json.data;
            $('#index_user_menu').find('.notifymenu').remove();
            var logout = $('#index_user_menu').html();
            $('#index_user_menu').empty();
            $.each(result.list,function(i,v){
               var html = '<div class="notifymenu" ';
               html += ' onclick="window.location=\''+v.url+'\'" ';
               html += '>';
               html += v.text;
               html += '</div>';
               $('#index_user_menu').append(html);
            });
            $('#index_user_menu').append(logout);
            $('#notify').html(result.cnt);
            if(parseInt(result.cnt)>0){
                $('#notify').show();
            }else{
                $('#notify').hide();
            }
        }else{

        }
    });
}


function checkform(jq_obj){
    var err = '';
    $form = jq_obj;
    $form.find('input[type=text]:required,input[type=number]:required,input[type=password]:required,select:required').each(function(idx,item){
        if($(item).val()==""){
            var label = $form.find('label[for='+$(item).attr("id")+']').html();
            label = label.replace(/：/g, "");
            err += '['+label+'] 不可為空白!!\n';
        }
    });

    return err;
}


function uuidv4() {
  return ([1e7]+-1e3+-4e3+-8e3+-1e11).replace(/[018]/g, c =>
    (c ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> c / 4).toString(16)
  );
}


