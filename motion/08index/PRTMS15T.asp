<%
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("../DATA/15Talks.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath


Set rsGN = Server.CreateObject("ADODB.Recordset")
rsGN.Open  "Select Distinct GN From D15T Order By GN Desc", conn, 3
	While Not rsGN.EOF
		GN = rsGN("GN")
		GNoption = GNoption & "<option value='" & GN & "'>" & GN & "</option>"
		rsGN.MoveNext  
	Wend
	
Set rsHN = Server.CreateObject("ADODB.Recordset")
rsHN.Open  "Select Distinct HN From D15T Order By HN Asc", conn, 3
	While Not rsHN.EOF
		HN = rsHN("HN")
		HNoption = HNoption & "<option value='" & HN & "'>" & HN & "</option>"  
		rsHN.MoveNext  
	Wend

Set rsHB = Server.CreateObject("ADODB.Recordset")
rsHB.Open  "Select Distinct HB From D15T Order By HB Asc", conn, 3
	While Not rsHB.EOF
		HB2 = rsHB("HB")
		IF HB2 = "1" Then     
			HB = "定期會"     
		END IF     
		IF HB2 = "2" Then     
			HB = "臨時會"     
		END IF     
		IF HB2 <> "" Then     
			HBoption = HBoption & "<option value='" & HB & "'>" & HB & "</option>"  
		END IF     
		rsHB.MoveNext  
	Wend
	
Set rsKIND = Server.CreateObject("ADODB.Recordset")
rsKIND.Open  "Select Distinct KIND From D15T Order By KIND Asc", conn, 3
	While Not rsKIND.EOF
		KIND = rsKIND("KIND")
		KINDoption = KINDoption & "<option value='" & KIND & "'>" & KIND & "</option>"  
		rsKIND.MoveNext  
	Wend
	
	
rsGN.Close
rsHN.Close
rsHB.Close
rsKIND.Close
conn.Close
%>
<html>

<head>

<title>議事檢索</title>

<link rel="stylesheet" href="../css/css.css" type="text/css">
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head><body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<table width="50%" border="0" cellpadding="0" cellspacing="0" height="100%">
  <tr> 
    <td background="../img/left-2.jpg" valign="top" height="90%"><img src="../img/top-tit.jpg" width="699" height="92"><br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="10" valign="top"><img src="../img/left-1.jpg" width="126" height="54"><br>
            <a href="../default-2.asp"><img src="../img/home-1.jpg" width="126" height="28" border="0" name="Image1" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','../img/home-2.jpg',0)"></a></td>
          <td>
            <table width="95%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td valign="top" align="center"> 
                  <blockquote>
                    <div align="left">[<a href="default.asp">議事檢索</a>] / 第十五屆討論報告案 </div>          
                  </blockquote>     
                  <table width="95%" border="0" align="center" cellspacing="0" cellpadding="0">     
                    <tr valign="top">      
                      <td width="1" height="1" align="left"><img src="../img/fra1.jpg" width="15" height="15"></td>     
                      <td background="../img/fra11.jpg" height="1"></td>     
                      <td width="1" height="1" align="right"><img src="../img/fra2.jpg" width="20" height="15"></td>     
                    </tr>     
                    <tr>      
                      <td background="../img/fra21.jpg" width="1" align="left"></td>     
                      <td>      
                        <table border="0" cellpadding="2" cellspacing="2" width="100%">     
                          <tr>      
                            <td valign="top" background="../img/flower-1.jpg" msnavigation>      
                              <form action="Search_15T.asp" method="POST">     
                                <div align="center">      
                                  <p>&nbsp;</p>     
                                  <table border="1" cellpadding="2" cellspacing="1" bordercolor="#CCCC99" width="95%" align="center">     
                                    <tr align="center">      
                                      <td colspan="2" height="25" bgcolor="#999966">      
                                        <div align="left">
                                          <p align="center"><font color="#FFFFFF" size="3">基　隆　市　議　會　　　</font><font color="#FFFFFF" size="3">第　十　五　屆<br>
                                          討論報告案暨處理情形資料庫網路查詢系統<br>
                                          </font></div>    
                                      </td>    
                                    </tr>    
                                    <tr>     
                                      <td align="center" nowrap bgcolor="CCCC99"><font color="#008080">&nbsp;&nbsp;&nbsp;&nbsp;              
                                        屆別：</font></td>     
                                      <td><font color="#666666">第              
                                        <select size="1" name="GN">     
                                          <option selected value="ALL">全部</option>     
                                          <% = GNoption %>     
                                        </select>     
                                        &nbsp; 屆</font></td>             
                                    </tr>     
                                    <tr>  
                                      <td align="center" nowrap bgcolor="CCCC99"><font color="#008080">&nbsp;&nbsp;&nbsp;&nbsp;          
                                        會期：</font></td>    
                                      <td><font color="#666666">第              
                                        <select size="1" name="HN">     
                                          <option selected value="ALL">全部</option>     
                                          <% = HNoption %>     
                                        </select> &nbsp; 次大會　</font></td>
                                    </tr>
                                    <tr>    
                                      <td align="center" nowrap bgcolor="CCCC99"><font color="#008080">&nbsp;&nbsp;&nbsp;&nbsp;          
                                        會別：</font></td>    
                                      <td><font color="#666666"> 　      
                                        <select size="1" name="HB">     
                                          <option selected value="ALL">全部</option>     
                                          <% = HBoption %>     
                                        </select>     
                                        </font></td>     
                                    </tr>     
                                    <tr>   
                                      <td align="center" nowrap bgcolor="CCCC99"><font color="#008080">&nbsp;&nbsp;&nbsp;&nbsp;          
                                        類別：</font></td>    
                                      <td><font color="#666666"> 　      
                                        <select size="1" name="KIND">     
                                          <option selected value="ALL">全部</option>     
                                          <% = KINDoption %>     
                                        </select>     
                                        </font></td>     
                                    </tr>     
                                    <tr>      
                                      <td height="80" colspan="2" valign="middle">      
                                        <p align="center"><font color="#008080">      
                                          &nbsp;&nbsp;&nbsp;&nbsp; <br>     
                                          <font color="#996600"> 包含下列文字者：</font><br>     
                                          </font><font color="#008080">&nbsp;              
                                          <input type="text" name="Key_Word" size="40">     
                                          <br>     
                                          <br>     
                                          </font></p>     
                                      </td>     
                                    </tr>     
                                    <tr valign="middle">      
                                      <td colspan="2" height="35" bgcolor="CCCC99">      
                                        <p align="center">      
                                          <input type="submit" value="開始搜尋" name="Search">     
                                          &nbsp;&nbsp;&nbsp;&nbsp;      
                                          <input type="reset" value="重設條件" name="Reset">     
                                          &nbsp;&nbsp;&nbsp;&nbsp; <a href="javascript:history.go( -1 );">      
                                          <input type="button" value="放棄搜尋" name="Back">     
                                          </a>      
                                      </td>     
                                    </tr>     
                                  </table>     
                                  <br>     
                                </div>     
                              </form>     
                              　 
                            </td> 
                          </tr> 
                        </table> 
                      </td> 
                      <td background="../img/fra31.jpg" width="1" align="right">&nbsp;</td> 
                    </tr> 
                    <tr valign="bottom">  
                      <td height="1" width="1" align="left"><img src="../img/fra3.jpg" width="15" height="21"></td> 
                      <td background="../img/fra41.jpg" height="1"></td> 
                      <td width="1" height="1" align="right"><img src="../img/fra4.jpg" width="20" height="21"></td> 
                    </tr> 
                  </table> 
                  <p>&nbsp;</p> 
                </td> 
              </tr> 
            </table> 
          </td> 
        </tr> 
      </table> 
    </td> 
  </tr> 
  <tr>  
    <td background="../img/left-2.jpg" valign="bottom" height="90%"><img src="../img/left-3.jpg" width="126" height="54"></td> 
  </tr> 
</table> 
<div id="Layer1" style="position:absolute; left:659px; top:0px; width:20px; height:41px; z-index:2"><img src="../img/ver-eng.jpg" width="110" height="31" usemap="#Map" border="0">  
  <map name="Map">  
    <area shape="rect" coords="27,4,89,20" href="../default.htm"> 
  </map> 
</div> 
<div id="img" style="position:absolute; left:0px; top:173px; width:78px; height:91px; z-index:1"><img src="../img/img-5.jpg" width="126" height="167"></div> 
</body> 
 
</html> 
