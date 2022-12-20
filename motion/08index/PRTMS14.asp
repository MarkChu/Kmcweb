<!-- #include file="Cn.asp" -->
<%



Set rsGN = Server.CreateObject("ADODB.Recordset")
rsGN.Open  "Select Distinct GN From D14 Order By GN Desc", conn, 3
	While Not rsGN.EOF
		GN = rsGN("GN")
		GNoption = GNoption & "<option value='" & GN & "'>" & GN & "</option>"
		rsGN.MoveNext  
	Wend
	
Set rsHN = Server.CreateObject("ADODB.Recordset")
rsHN.Open  "Select Distinct HN From D14 Order By HN Asc", conn, 3
	While Not rsHN.EOF
		HN = rsHN("HN")
		HNoption = HNoption & "<option value='" & HN & "'>" & HN & "</option>"  
		rsHN.MoveNext  
	Wend

Set rsHB = Server.CreateObject("ADODB.Recordset")
rsHB.Open  "Select Distinct HB From D14 Order By HB Asc", conn, 3
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
rsKIND.Open  "Select Distinct KIND From D14 Order By KIND Asc", conn, 3
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

                              <form action="Discuss14R.asp" method="POST"> 
                                <INPUT type="hidden" name="GN" value="14">  
                                <div align="center">      
                                  <p>&nbsp;</p>     
                                  <table border="1" cellpadding="2" cellspacing="1" bordercolor="#CCCC99" width="95%" align="center">     
                                    <tr align="center">      
                                      <td colspan="2" height="25" bgcolor="#999966">      
                                        <div align="left">
                                          <p align="center"><font color="#FFFFFF" size="3">基　隆　市　議　會　第　十　四　屆<br>
                                          </font><font color="#FFFFFF" size="3">議決案暨處理情形資料庫網路查詢系統<br>
                                          </font></div>    
                                      </td>    
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