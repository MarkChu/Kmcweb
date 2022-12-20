<!-- #include file="Cn.asp" -->
<%


Set rsHN = Server.CreateObject("ADODB.Recordset")

rsHN.Open  "Select Distinct HN From D0110 Order By HN ", conn, 3
	While Not rsHN.EOF
		HN = rsHN("HN")
		HNoption = HNoption & "<option value='" & HN & "'>" & HN & "</option>"  
		rsHN.MoveNext  
	Wend
rsHN.Close
conn.Close
%>    
                              <form action="Discuss10R.asp" method="POST">     
                                <div align="center">      
                                  <p>&nbsp;</p>     
                                  <table border="1" cellpadding="2" cellspacing="1" bordercolor="#CCCC99" width="95%" align="center">     
                                    <tr align="center">      
                                      <td colspan="2" height="25" bgcolor="#999966">      
                                        <div align="left">
                                          <p align="center"><font color="#FFFFFF" size="3">基隆市議會　　　　第一屆至第十屆<br>
                                          </font><font color="#FFFFFF" size="3">議案暨處理情形資料庫網路查詢系統<br>
                                          </font></div>    
                                      </td>    
                                    </tr>    
                                    <tr>  
                                      <td align="center" nowrap bgcolor="CCCC99"><font color="#008080">&nbsp;&nbsp;&nbsp;&nbsp;      
                                        冊別：</font></td>    
                                      <td><font color="#666666">第                      
                                        <select size="1" name="HN"><% = HNoption %>         
                                        </select> &nbsp; 冊</font></td>     
                                    </tr>
                                    <tr valign="middle">      
                                      <td colspan="2" height="35" bgcolor="CCCC99">      
                                        <p align="center">      
                                          <input type="submit" value="開始搜尋" name="Search"> 
                                          &nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="重設條件" name="Reset">          
                                          &nbsp;&nbsp;&nbsp;&nbsp; <a href="javascript:history.go( -1 );">              
                                          <input type="button" value="放棄搜尋" name="Back">     
                                          </a>      
                                      </td>     
                                    </tr>     
                                  </table>     
                                  <br>     
                                </div>     
                              </form>
