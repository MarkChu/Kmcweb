<!DOCTYPE html>
<head>
<meta charset="UTF-8">
<TITLE>基隆市議會會內資訊網</TITLE>
<link rel="stylesheet" href="css/bootstrap.min.css">
<link href="css/style.css" rel="stylesheet" type="text/css">
<link href="fontawesome/css/all.css" rel="stylesheet">
<script src="js/jquery.js"></script>
<script src="js/jquery.blockUI.js"></script>
<script src="js/custom.js"></script>
<script src="js/default.js"></script>
<style>
.log {
    background: #f4f4f4;
    border-radius: 15px;
    width: 350px;
    height: 310px;
    margin: 30px auto;
    border: solid 1px #d3d3d3;
    padding: 20px;
    text-align: center;
}      
.log-headicon {
    width: 25px;
    height: 30px !important;
    font-size: 16px;
    text-align: center;
    background: #bebebe;
    border: solid 1px #b3b3b3;
}      
</style>
</head>
<body>
<div id="containe">

<div id="header"><img src="images/logo.png" class="logo"></div>

<div id="wrap">

      <div id="content_log">

        <div class="log">
            <div align="center"><img src="images/log-title.png"></div>
      	<div class="form-line">
                  <div class="log-headicon"><i class="fas fa-user"></i></div>
                  <div class=""><input id="a1" type="text" autocomplete="off" maxlength="20"></div>
                  
            </div>
            <div class="clearboth"></div>
            
            <div class="form-line">
                  <div class="log-headicon"><i class="fas fa-lock"></i></div>
                  <div class=""><input id="a2" type="password" autocomplete="new-password" maxlength="20"></div>    
                   
            </div>
            <div class="clearboth"></div>
            <div><img id="authimg" src="GetAuthCode.asp"></div>
            <div class="verification ">
                  <div class="">驗證碼</div>
                  <div class=""><input id="a3" type="text" autocomplete="off" maxlength="6"></div>
            </div>
            <div class="clearboth"></div>

            <button type="button" class="btn log-bt" onclick="dologin();"><i class="fas fa-user"></i>登入</button>
        </div>

      </div>

      <div class="clearboth"></div>

</div>

</div>

</body>