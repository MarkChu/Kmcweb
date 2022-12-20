<%Response.CodePage = 65001%>
<%Response.Charset = "utf-8"%>
<!DOCTYPE html>
<head>
<meta charset="UTF-8">
<TITLE>基隆市議會會內資訊網</TITLE>
<link rel="stylesheet" href="fontawesome/css/all.css">
<link rel="stylesheet" href="css/jquery-ui.css">
<link rel="stylesheet" href="css/bootstrap.css">
<link rel="stylesheet" href="css/footable.standalone.css">
<link rel="stylesheet" href="css/style.css" type="text/css">
<script type="text/javascript" src="js/jquery.js"></script>
<script type="text/javascript" src="js/jquery-ui.js"></script>
<script type="text/javascript" src="js/moment-with-locales.js"></script>
<script type="text/javascript" src="js/bootstrap.js"></script>
<script type="text/javascript" src="js/footable.js"></script>
<script type="text/javascript" src="js/jquery.blockUI.js"></script>
<script type="text/javascript" src="js/custom.js"></script>
</head>
<!--#include file="inc/Common.asp"-->
<!--#include file="inc/Func.asp"-->
<!--#include file="index_setvar.asp"-->

<body>
<div id="containe">

<div id="header"><a href="index.asp"><img src="images/logo.png" class="logo" border="0"></a></div>

<div id="wrap">
<div id="bd">
<div id="index_left">
    <!--#include file="index_left.asp"-->
</div>

<div id="index_content">
  <div class="content_area">
  	<div class="tb">
        <!--#include file="index_content.asp"-->
     </div>
  </div>

</div>
 
<div class="clearboth"></div>


</div>
</div>

</div>
</body>
