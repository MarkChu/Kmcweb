 <html>

<head>

<title>基 隆 市 議 會 討 論 報 告 案 查 詢 系 統</title>

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
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<table width="50%" border="0" cellpadding="0" cellspacing="0" height="100%">
  <tr> 
    <td background="../img/left-2.jpg" valign="top" height="90%"><img src="../img/top-tit.jpg"><br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="10" valign="top"><img src="../img/left-1.jpg" width="126" height="54"><br>
            <a href="../default-2.asp"><img src="../img/home-1.jpg" width="126" height="28" border="0" name="Image1" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','../img/home-2.jpg',0)"></a></td>
          <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td width="586" valign="top" align="center"> 
                  <table width="95%" border="0" cellspacing="2" cellpadding="2" class="text05_12" align="center">
                    <tr> 
                      <td width="170" >[<a href="default.asp">議事檢索</a>]/查 詢 結 果</td>                                                                          
                      <td >
                        <p align="right">[<a href="javascript:history.go( -1 );">重新查詢</a>]</p>
 </td>                                                      
                      <td > </td>                                                      
                    </tr>                                                      
                  </table>                                                      
                  <table border="0" cellpadding="2" cellspacing="2" width="95%" align="center">                                                      
                    <tr>                                                       
                      <td valign="top">                                                       
                        <%                                                          
Page= Request("Page")                                                          
If Page = "" Then                                                          
   Page = 1                                                          
End If                                                          
GN			= Request("GN")                                                          
HN 			= Request("HN")                                                          
HB			= Request("HB")                                                          
KIND		= Request("KIND")                                                          
Key_Word	= Request("Key_Word")                                                          
Set conn = Server.CreateObject("ADODB.Connection")                                                          
DBPath = Server.MapPath("../DATA/15Talks.mdb")                                                          
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath                                                          
	IF GN = "ALL" Then                                                          
		GNtmp = ""                                                          
		GNcfm = 1                                                          
	Else                                                          
		GNtmp = "GN =" & GN & " AND "                                                          
		GNcfm = 0                                                          
	END IF                                                          
                                                     
	IF HN = "ALL" Then                                                          
		HNtmp = ""                                                          
		HNcfm = 1                                                          
	Else                                                          
		HNtmp = "HN =" & HN & " AND "                                                          
		HNcfm = 0                                                          
	END IF                                                          
                                                     
	IF HB = "ALL" Then                                                          
		HBtmp = ""                                                          
		HBcfm = 1                                                          
	Else                                            
		IF HB = "定期會" Then                                                          
			HBtmp = "HB = 1 AND "                                                          
		END IF                                                          
		IF HB = "臨時會" Then                                                          
			HBtmp = "HB = 2 AND "                                                          
		END IF                                                          
		HBcfm = 0                                                          
	END IF                                                          
	                                                     
	IF KIND = "ALL" Then                                                          
		KINDtmp = ""                                                          
		KINDcfm = 1                                                          
	Else                                                          
		KINDtmp = "KIND = '" & KIND & "' AND "                                                          
		KINDcfm = 0                                                          
	END IF                                                          
	                                                     
	Set rs = Server.CreateObject("ADODB.Recordset")                                                          
	IF Key_Word = "" Then                                                        
		rs.Open	"Select * From D15 Where " & GNtmp & HNtmp & HBtmp & KINDtmp &        " ID <> 0 Order By ID", conn, 3                                                          
	Else                                                      
		Etmp = " ( EW Like '%" & Key_Word & "%' OR EH Like '%" & Key_Word & "%' OR EO Like '%" & Key_Word & "%' ) AND "                                                      
		rs.Open	"Select * From D15 Where " & GNtmp & HNtmp & HBtmp & KINDtmp & Etmp & " ID <> 0 Order By ID", conn, 3                         
	End IF                                                          
	rsnum = rs.RecordCount                                                       
	IF rsnum = 0 Then                                                          
		ShowEmpty                                                          
	END IF                                                          
	                                
	rs.PageSize = 20                                                                     
	rs.AbsolutePage = Page                                                                     
	                                
	If rs.AbsolutePage > rs.PageCount Then                                                                     
		rs.AbsolutePage = 1                                                                     
	End If                                                      
	For p = 1 to rs.PageCount                                                          
		PageLink = PageLink & "<a href='search_15.asp?Page=" & p & "&GN=" & GN & "&HN=" & HN & "&HB=" & HB & "&KIND=" & KIND & "&Key_Word=" & Key_Word & " '>[" & p & "]</a>"                                                          
	Next                                         
	                                             
	If Page > 1 Then                                                                                                                     
		BackPage = "<a href='search_15.asp?Page=" & Page-1 & "&GN=" & GN & "&HN=" & HN & "&HB=" & HB & "&KIND=" & KIND & "&Key_Word=" & Key_Word & " '>[上一頁]</a> "                                                                                                       
	Else                                                                                                       
		BackPage = ""                                                                                                       
	End If                                                                                                              
	                                
	If Page - rs.PageCount < 0 Then                                                                                                                     
		NextPage = "<a href='search_15.asp?Page=" & Page+1 & "&GN=" & GN & "&HN=" & HN & "&HB=" & HB & "&KIND=" & KIND & "&Key_Word=" & Key_Word & " '>[下一頁]</a> "                                                                                                       
	Else                                                                                                       
		NextPage = ""                                                                                                       
	End If                                                      
	ReSearch = "<a href='default.asp'>[重新查詢]</a>"                       
%>                                                      
                        <div align="center">                                                       
                          <table border="0" cellpadding="0" cellspacing="0" width="95%" align="center">                                                      
                            <tr>                                                       
                              <td bgcolor="#0066FF" align="center" class="text06_12">基　隆　市　議　會　&nbsp;          
                                第十五屆<br>
                                議 決 案 查 詢 系 統 - 查 詢 結 果</td>                                                                          
                            </tr>                                                      
                            <tr>                                                       
                              <td align="center"><font size="2" color="#0000FF"><br>                                                      
                                <br>                                                      
                                </font> <font size="2"> <font color="#0000FF">總共搜尋到</font> <font color="#FF0000" size="3"> <% = rsnum %>                                                                          
                                </font> <font color="#0000FF">筆資料 ： 每頁顯示 20 筆資料                                                                           
                                ： 目前顯示第 </font> <font color="#FF0000"><% = Page %></font> <font color="#0000FF">&nbsp頁<br> </font></font>                                                                           
                              </td>                                                      
                            </tr>                                                      
                          </table>                                                      
                        </div>                                                      
                        <div align="center">                                                       
                          <table border="0" cellpadding="0" cellspacing="0">                                                      
                            <tr>                                                       
                              <td>                                                       
                                <p align="center">                                                       
<%      no = (page-1) * 20	                                                                                                                                                             
		For i = 1 to rs.PageSize                                                                                             
 		PassChk = "NoPass"                      
		While PassChk = "NoPass"                 
		                         
    	  IDrs = rs("ID")                            
    	  GNrs = rs("GN")                            
    	  HNrs = rs("HN")                               
    	  HBchk = rs("HB")                               
    	  KINDrs = rs("KIND")                                                                                                                                                            
          IF HB <> "ALL" Then                                                                                                             
			 IF HBchk = "" Then                                                                                                                                                                                          
				HBrs = "</font><font size='2' color='#FF0000'> </font><font size='2'>"                                                                                                                                                                                          
			 End If                                                                                                                                                
			 IF HBchk = "1" Then                                                                                                                                                                                          
				HBrs = "</font><font size='2' color='#FF0000'> 定期 </font><font size='2'>"                                                                                                                                                                                          
			 End If                                                                                                                                                
			 IF HBchk = "2" Then                                                                                                                                                                                          
				HBrs = "</font><font size='2' color='#FF0000'> 臨時 </font><font size='2'>"                                                                                                                                                                                          
			 End If                                                                                                                                                
          Else                                                                 
			 IF HBchk = "" Then                                                                                                                                                                                          
				HBrs = ""                                                                                                                                                                                          
			 End If                                                                                                                                                               
			 IF HBchk = "1" Then                                                                                                                                                                                          
				HBrs = "定期"                                                                                                                                                                                          
			 End If                                                                                                                                                               
			 IF HBchk = "2" Then                                                                                                                                                                                          
				HBrs = "臨時"                                                                                                                                                                                          
			 End If                                                                                                                                                               
          End If                                                                 
                                                                                                                                                      
	  	  IF Key_Word = "" Then                                                                                                                                                           
        	EWrs = rs("EW")                                                                                                                                                                                          
			EHrs = rs("EH")                                                                                                                                                                                          
			EOrs = rs("EO")                                                                                                                                           
            no = no + 1                                                                                                                                                                                     
			ShowRS rs                                                                                                                                                                                        
    		PassChk = "Pass"                                                                                                                                            
		  Else                                                                                                                        
				EWtmp2 = rs("EW")                                                                                                                               
	          	IF EWtmp2 <> "" Then                                                                                                                               
				  	EWre = "</font><font size='2' color='#FF0000'>" & Key_Word & "</font><font size='2'>"                                                                                                                                                                                          
	        	  	EWrs = Replace (EWtmp2,Key_Word,EWre)                                                                                                                                                                                          
				Else                                                                                                                               
					EWrs = EWtmp2                                                                                                                              
				END IF                                                                                                                                         
                                                     
  				EHtmp2 = rs("EH")                                                                                                                                                                                          
	          	IF EHtmp2 <> "" Then                                                                                                                               
				  	EHre = "</font><font size='2' color='#FF0000'>" & Key_Word & "</font><font size='2'>"                                                                                                                                                                                          
				  	EHrs = Replace (EHtmp2,Key_Word,EHre)                                                                                                                                                                                          
				Else                                                                                                                               
					EHrs = EHtmp2                                                                                                                              
				END IF                                                                                          
                                                     
		   		EOtmp2 = rs("EO")                                                                                                                                                                                          
	          	IF EOtmp2 <> "" Then                                                                                                                               
				  	EOre = "</font><font size='2' color='#FF0000'>" & Key_Word & "</font><font size='2'>"                                                                                                                                                                                          
				  	EOrs = Replace (EOtmp2,Key_Word,EOre)                                                                                                                                                                                        
				Else                                                                                                                               
					EOrs = EOtmp2                                                                                                                              
				END IF                                                                                                                              
				                                                     
	            no = no + 1 	                                                                                                                                                                                   
	   			ShowRS rs                                                                                                                           
	   			PassChk = "Pass"                                                                                                                           
  			End IF                                                                                                                                                                                          
           	rs.MoveNext                                                                                                
			If rs.EOF Then                                                                                                                                                                                                         
  				Exit For                                                                                                                                                                                                         
			End If                                                                                                                                                                                 
			Wend                                                                                                                                                                                              
		Next                                                                                                                                                               
%>                                                      
                              </td>                                                      
                            </tr>                                                      
                          </table>                                                      
                        </div>                                                      
                        <p></p>                                                      
                        <div align="center">                                                       
                          <table border="0" cellpadding="0" cellspacing="0">                                                      
                            <tr>                                                       
                              <td height="30">                                                       
                                <p align="center"><font size="2">                                                       
                                  <% = BackPage %>                                                      
                                  <% = NextPage %>                                                      
                                  <% = ReSearch %>                                                      
                                  </font>                                                       
                              </td>                                                      
                            </tr>                                                      
                            <tr>                                                       
                              <td height="30">                                                       
                                <div align="center">                                                       
                                  <table border="0" cellpadding="0" cellspacing="0">                                                      
                                    <tr>                                                       
                                      <td>                                                       
                                        <p align="center"><font size="2">頁次：                                                       
                                          <% = PageLink %>                                                      
                                          </font>                                                       
                                      </td>                                                      
                                    </tr>                                                      
                                  </table>                                                      
                                </div>                                                      
                              </td>                                                      
                            </tr>                                                      
                            <tr>                                                       
                              <td height="30">                                                       
                                <p align="center"> <font size="2"> <font color="#0000FF">總共搜尋到</font>                                                       
                                  <font color="#FF0000" size="3">                                                                           
                                  <% = rsnum %>                                                      
                                  </font> <font color="#0000FF">筆資料 ： 每頁顯示 20                                                                           
                                  筆資料 ： 目前顯示第 </font> <font color="#FF0000">                                                                           
                                  <% = Page %>                                                      
                                  </font> <font color="#0000FF">&nbsp頁<br>                                                      
                                  </font></font> </p>                                                      
                              </td>                                                      
                            </tr>                                                      
                          </table>                                                      
                        </div>                                                      
                        <%                                                                                                                                                                                                             
		IF no=0 Then                                                                                                                                                                                                               
			ShowEmpty                                                                                                                                        
		End If                                                                                             
%>                                                      
                        <p></p>                 
                                 
                                    
                      </td>                                                   
                    </tr>                                                   
                  </table>                                                   
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
<!------------------------------------------------------------------------------------------------------------>       
                        <%  Sub ShowRS(rs)  %>                                                          
                        <table border="0" width="100%">                                                      
                          <tr>                                                       
                            <td colspan="2"> 　</td>                                                     
                          </tr>                                                     
                          <tr bgcolor="#CCCCCC">                                                      
                            <td width="294"> <font size="2">第                                                      
                              <% = no %>                                                     
                              筆查詢結果</font></td>                                                     
                            <td width="292"><font size="2">ＩＤ：                                                                               
                              <% = IDrs %>                                                          
                              &nbsp;</font></td>                                                                             
                          </tr>                                                         
                          <tr bgcolor="#CCCCCC">                                                          
                            <td colspan="2" align="center"><font size="2"> 第 <% = GNrs %> 屆                                                                     
                              第 <% = HNrs %> 次 <% = HBrs %> 會 </font>                                                                       
                            </td>                                                        
                          </tr>                                                        
                          <tr>                                                         
                            <td  colspan="2">                                                         
                              <table cellpadding="1" cellspacing="1" border="0" width="100%">                                                        
                                <tr bgcolor="#F2F1FC">                                                         
                                  <td width="15%" align="right"><font size="2"> 案　　由：</font></td>                                                        
                                  <td width="87%"><font size="2">                                                         
                                    <% = EWrs %>                                                        
                                    </font></td>                                                        
                                </tr>                                                        
                                <tr bgcolor="#F2F1FC">                                                         
                                  <td width="15%" align="right"><font size="2">提&nbsp;                                                                  
                                    案&nbsp; 人：</font></td>                                                                           
                                  <td width="87%"><font size="2">                                                        
                                    <% = EHrs %>                                                       
                                    </font></td>                                                       
                                </tr>                                                       
                                <tr bgcolor="#F2F1FC">                                                        
                                  <td width="15%" align="right"><font size="2"> 處理情形：</font></td>                                                       
                                  <td width="87%"><font size="2">                                                        
                                    <% = EOrs %>                                                       
                                    </font></td>                                                       
                                </tr>                                                       
                              </table>                                                       
                            </td>                                                       
                          </tr>                                                       
                        </table>                                                       
                        <% End Sub %>               
       
                        <%  Sub ShowEmpty  %>                                                       
                        <p align="center">　</p>                                                       
                        <p align="center"><font size="3" color="#FF0000">查無符合相關條件的資料！</font></p>                                                       
                        <p align="center"><font color="#FF0000"><font size="3">請重新設定條件查詢！</font></font></p>                                                       
                        <p align="center"><font size="3" color="#0000FF"></font><a href="javascript:history.go( -1 );">重新查詢</a></p>                                                       
                        <p align="center">　</p>                                                       
                        <% Response.End %>                                                       
						<% End Sub %>                                                    
</body>                                                                   
</html>       
       
