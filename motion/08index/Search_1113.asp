<!-- #include file="Cn.asp" -->
<table WIDTH="100%" CELLPADDING="0" border="0" CELLSPACING="0"><!--OutsideBorderTable-->
  <tr><td WIDTH="1" HEIGHT="5"><img SRC="/ImgSys/NT-Spacer.gif" HEIGHT="5" WIDTH="1"></td></tr><!--TopBorderZone-->
  <tr>
    <td WIDTH="5" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" HEIGHT="1" WIDTH="5"></td><!--LeftBorderZone-->
    <td ALIGN="center" WIDTH="100%"><!--InsideTable-->
      <table WIDTH="95%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
        <tr>
          <td ALIGN="center" COLSPAN="5"><table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="30%" height="20" bgcolor="#848484">
&nbsp;&nbsp;&nbsp;<img src="object/icon/record_icon.gif" align="absmiddle">
<font class="w12B">議事檢索 </font>
</td>
   <td width="70%" height="20" align="right" valign="bottom">  <!-- banner_menu -->
  
</td>
  <tr>
    <td width="100%" height="2" bgcolor="#848484" colspan="2"><img src="object/icon/2pix.gif" align="absmiddle"></td>
</table></td>
        </tr>
        <tr><!--InsideMargin:top-->
          <td BGCOLOR="#848484" WIDTH="1" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" WIDTH="1"></td><!--LeftLine-->
          <td COLSPAN="3" WIDTH="1" HEIGHT="5"><img SRC="/ImgSys/NT-Spacer.gif" HEIGHT="5" WIDTH="1"></td>
          <td BGCOLOR="#848484" WIDTH="1" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" WIDTH="1"></td><!-- RightLine-->
        </tr>
        <tr>
          <td BGCOLOR="#848484" WIDTH="1" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" WIDTH="1"></td><!--LeftLine-->
          <td WIDTH="10"><img SRC="/ImgSys/NT-Spacer.gif" HEIGHT="1" WIDTH="10"></td><!--InsideMargin.left-->
          <td WIDTH="100%">          
          
          
            <table width="100%" border="0" cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td width="586" valign="top" align="center"> 
                                                     
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
		rs.Open	"Select * From D1113 Where " & GNtmp & HNtmp & HBtmp & KINDtmp &        " ID <> 0 Order By ID", conn, 3                                                          
	Else                                                      
		Etmp = " ( EW Like '%" & Key_Word & "%' OR EH Like '%" & Key_Word & "%' OR EO Like '%" & Key_Word & "%' ) AND "                                                      
		rs.Open	"Select * From D1113 Where " & GNtmp & HNtmp & HBtmp & KINDtmp & Etmp & " ID <> 0 Order By ID", conn, 3                         
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
		PageLink = PageLink & "<a href='Discuss13R.asp?Page=" & p & "&GN=" & GN & "&HN=" & HN & "&HB=" & HB & "&KIND=" & KIND & "&Key_Word=" & Key_Word & " '>[" & p & "]</a>"                                                          
	Next                                         
	                                             
	If Page > 1 Then                                                                                                                     
		BackPage = "<a href='Discuss13R.asp?Page=" & Page-1 & "&GN=" & GN & "&HN=" & HN & "&HB=" & HB & "&KIND=" & KIND & "&Key_Word=" & Key_Word & " '>[上一頁]</a> "                                                                                                       
	Else                                                                                                       
		BackPage = ""                                                                                                       
	End If                                                                                                              
	                                
	If Page - rs.PageCount < 0 Then                                                                                                                     
		NextPage = "<a href='Discuss13R.asp?Page=" & Page+1 & "&GN=" & GN & "&HN=" & HN & "&HB=" & HB & "&KIND=" & KIND & "&Key_Word=" & Key_Word & " '>[下一頁]</a> "                                                                                                       
	Else                                                                                                       
		NextPage = ""                                                                                                       
	End If                                                      
%>                                                      
                                                                
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
%>                    <p></p>                 
                                 
                                    
                      </td>                                                   
                    </tr>   
                    </table>     </td>                                                   
                    </tr>   
                    </table>          
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
                        <p align="center">　</p> 
                         </td>                                                       
                          </tr>                                                       
                        </table>  </td>                                                       
                          </tr>                                                       
                        </table> 
                        
                         </td>
          <td WIDTH="10"><img SRC="/ImgSys/NT-Spacer.gif" HEIGHT="1" WIDTH="10"></td><!--InsideMargin.right-->
          <td BGCOLOR="#848484" WIDTH="1" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" WIDTH="1"></td><!-- RightLine-->
        </tr>
        <tr><!--InsideMargin:bottom-->
          <td BGCOLOR="#848484" WIDTH="1" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" WIDTH="1"></td><!--LeftLine-->
          <td COLSPAN="3" WIDTH="1" HEIGHT="5"><img SRC="/ImgSys/NT-Spacer.gif" HEIGHT="5" WIDTH="1"></td>
          <td BGCOLOR="#848484" WIDTH="1" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" WIDTH="1"></td><!-- RightLine-->
        </tr>
        <tr>
          <td ALIGN="center" COLSPAN="5"><table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" height="1" bgcolor="#848484"><img src="object/icon/1pix.gif" align="absmiddle"></td>
  </tr>
</table></td>
        </tr>
      </table></td>
  </tr>
</table>                                                                                     
                        <% Response.End %>                                                       
						<% End Sub %>                                                    
            </td>
          <td WIDTH="10"><img SRC="/ImgSys/NT-Spacer.gif" HEIGHT="1" WIDTH="10"></td><!--InsideMargin.right-->
          <td BGCOLOR="#848484" WIDTH="1" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" WIDTH="1"></td><!-- RightLine-->
        </tr>
        <tr><!--InsideMargin:bottom-->
          <td BGCOLOR="#848484" WIDTH="1" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" WIDTH="1"></td><!--LeftLine-->
          <td COLSPAN="3" WIDTH="1" HEIGHT="5"><img SRC="/ImgSys/NT-Spacer.gif" HEIGHT="5" WIDTH="1"></td>
          <td BGCOLOR="#848484" WIDTH="1" HEIGHT="1"><img SRC="/ImgSys/NT-Spacer.gif" WIDTH="1"></td><!-- RightLine-->
        </tr>
        <tr>
          <td ALIGN="center" COLSPAN="5"><table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" height="1" bgcolor="#848484"><img src="object/icon/1pix.gif" align="absmiddle"></td>
  </tr>
</table></td>
        </tr>
      </table></td>
  </tr>
</table>    
       
