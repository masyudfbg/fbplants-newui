
<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<html>
<head>
	<title>ePlaS - Get Data From WeighBridge</title>
	<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
	<script language="JavaScript" src="date-picker.js"></script>
	<script language="JavaScript">
		var newwin;
		function launchwin(winurl,winname,winfeatures)
		{
			//This launches a new window and then
			//focuses it if window.focus() is supported.
			newwin = window.open(winurl,winname,winfeatures);
		}
	</script> 
</head>
<body>
	<!--#include file="top.jsp"-->
	<table width="100%">
		<tr>
			<td width="22%" valign="top">
				<!--#include file="menu.asp"-->  
			</td>
			<td width="78%" valign="top">
			    <%
			    Dim txtToDayDate,txtFromDayDate, txtyear, txtmonth, txtfcba
				Dim strSQLPreParameter, rsSQLPreParameter,v_txtFromDayDate, v_txtToDayDate, v_txtfcba, v_txtyear, v_txtmonth,BRG
				dim  cmdView
				v_txtFromDayDate = Request("txtFromDayDate")
				v_txtToDayDate   = Request("txtToDayDate")
               	cmdView = request("cmdView")			
				%>
					<form name="frmEntry" method="post" action="p_WeighBridge_GetdataCPO.Asp">
						<table class="data" width="80%" cellpadding="0" cellspacing="0">
						            <tr>
										<th>&nbsp;</th>
										<th>&nbsp;</th>
									</tr>
						            <tr>
										<h2>Ambil Data CPO dari WeighBridge ke IPlaS ( Delivery Note dan WeighBridge Ticket )</h2>
										
									</tr>
									<tr>
										<th>&nbsp;</th>
									</tr>
									<tr>
										<td ><div align="center"><font size="3" face="Verdana"><strong>Tanggal Awal</strong></font></div></td>
										<td ><div align="center"><font size="3" face="Verdana"><strong>Tanggal Akhir</strong></font></div></td>
									</tr>
									

									<tr>
										<td align="center">
										<% if v_txtFromDayDate <> "" and v_txtFromDayDate = v_txtFromDayDate then %>
											<input type="text" maxlength="10" name="txtFromDayDate" size="10" value="<%=v_txtFromDayDate%>" title="Type Date mm/dd/yyyy" >
										<% else %>
											<input type="text" maxlength="10" name="txtFromDayDate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
										<% end if%>
											<a href="javascript:show_calendar('frmEntry.txtFromDayDate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
											<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
											</a>
										</td>
										
										<td align="center">
										<% if v_txtToDayDate <> "" and v_txtToDayDate = v_txtToDayDate then %>
												<input type="text" maxlength="10" name="txtToDayDate" size="10" value="<%=v_txtToDayDate%>" title="Type Date mm/dd/yyyy" >
										<% else %>
												<input type="text" maxlength="10" name="txtToDayDate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
										<% end if%>
										<a href="javascript:show_calendar('frmEntry.txtToDayDate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
										<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
										</a>
										</td>	
										<input Type="Hidden" Value="True" Name="cmdView">
										<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
										<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">	
                                        									
										<td align="center"><input type="submit" value="..Process.." /></td>	
									</tr>
									<tr>
										<th>&nbsp;</th>
										<th>&nbsp;</th>
									</tr>
								
						</table>	
					
				
				    <%
			        IF cmdView = "True" Then
					    Const adOpenForwardOnly = 0
						Const adLockReadOnly = 1
						Const adCmdTableDirect = &H0200
						Const adUseClient = 3
						Dim  strSQL,rsAsal,strSQLwb,rswb,q,strSQLCust,rscust
						Dim  strSQLweight,rsweight , CONTRACTCODE
						Dim TimeAwal,TimeAkhir,Hitung,Lama,CHITNO
						TimeAwal=Time()
						Dim Startingtime, Endingtime
						Startingtime = FormatDateTime(CDate(Now))
					%>
						Please Wait...<br><br>
					<%
						
						
						Response.Write "Collecting Data From Date ..." & v_txtFromDayDate & " To.. " & v_txtToDayDate & "<br>" 
						Response.Write "Start ..." & Time() & "<br><br>" 
						
						strSQL = " SELECT to_char(t.tanggalm,'YYYYMMDD')  tanggalm,t.waktu1, t.tanggalk,t.waktu2,t.bruto,t.tarra,replace(t.nopolisi ,'''',' ') nopolisi,t.tanggalm dnote_date,t.nodo deliverynotecode, t.kontrak contractcode,t.nosegel gatepass, " &_
						         " '00' WBT_CHITNO,'BST1' STORAGECODE , T.Tarra - t.bruto QUANTITY,t.ffa ,t.kotoran DIRT,t.moisture, " &_
								 " t.angkutan transportco, replace(t.supir ,'''',' ') drivername ,t.sim ICNO ,t.kd_brg" &_
						         "  FROM WBST_CPO T WHERE  t.tanggalm between to_date('" & v_txtFromDayDate & " ','MM/DD/YYYY') and to_date('" & v_txtToDayDate & "','MM/DD/YYYY')  and t.tarra is not null and   t.kontrak IS NOT NULL and ( t.error_flag is null or  t.error_flag <>  'Success' )  order by  to_char(t.tanggalm,'YYYYMMDD')"
						'Response.Write strSQL  & "<br><br>"
						Set rsAsal = Server.CreateObject("ADODB.Recordset")
						rsAsal.CursorType = 3
						rsAsal.CursorLocation = 3
						rsAsal.Open strSQL, connstr
						if not rsAsal.EOF then
						   Response.Write "Insert ke WEIGHBRIDGETICKET dan DELIVERYNOTE " & "<br><br>" 
			               while not rsAsal.EOF
							 CONTRACTCODE = rsAsal("contractcode")  
								 'Response.Write rsAsal("contractcode") & ", " 
								 'Check apakah No Contract ada di IPlaS
								 strSQLwb = "select * from deliverycontract t  where t.contractcode = '" & CONTRACTCODE  &"'"
								 set rswb = Server.CreateObject("ADODB.Recordset")
								 rswb.CursorType = 3
						         rswb.CursorLocation = 3
								 rswb.Open strSQLwb, connstr
								 'Response.Write strSQLwb & "<br>" 
								 if not rswb.eof then
								        
								       strSQLCust = "  select * from DELIVERYNOTE t  where t.contractcode ='" & CONTRACTCODE  & "'  and to_char(t.dnote_date,'YYYYMMDD') = '" & rsAsal("tanggalm") & "' and  t.icno = '" & rsAsal("ICNO") & "' and t.icno ='' "
									   set rscust   = Server.CreateObject("ADODB.Recordset")
									   rscust.CursorType = 3
						               rscust.CursorLocation = 3
									   'Response.Write strSQLCust & "<br>" 
								       rscust.Open strSQLCust, connstr
									   if rscust.eof then
									         if  rsAsal("kd_brg") = "PK" then
													    BRG = "KER"
											else
													    BRG = rsAsal("kd_brg")
										    end if
									        strSQLweight = " select NVL(QTY,0) QTY,CHITNO  from ( select sum(t.quantity) qty ,weight_bridge_api.get_chitno('" & Session("fcBA") & "','" & BRG  & "') chitno from DELIVERYNOTE t  where t.contractcode ='" & CONTRACTCODE  & "'  )  where   ( get_totweightDeliverycontract('" & CONTRACTCODE  & "','" & Session("fcBA") & "') - nvl(qty,0)) <> 0 "											
											set rsweight  = Server.CreateObject("ADODB.Recordset")
											rsweight.CursorType = 3
						                    rsweight.CursorLocation = 3
											'response.write  strSQLweight & "<br>"
											rsweight.Open strSQLweight, connstr
											if not rsweight.eof then
											        if  rsAsal("kd_brg") = "PK" then
													    BRG = "KER"
												    else
													    BRG = rsAsal("kd_brg")
													end if
									                Response.Write "Kd Barang : " & BRG & " , No Contract : " & CONTRACTCODE & " , tgl "& rsasal("dnote_date") & " , dgn No Polisi " & rsAsal("nopolisi") & " berhasil di Insert " & "<br>" 
													CHITNO = rsweight("chitno")
													
													'insert ke WEIGHBRIDGETICKET
													q=  " Insert into WEIGHBRIDGETICKET (TRANSACTIONTYPE, CHITNO, PRODUCT_PRODUCTCODE, VEHICLENO, CUST_CUSTOMERCODE, FIELD_ESTATE_ESTATECODE, " & _
 													    " FIELD_FIELDCODE, DATEIN, WBRIDGE_WEIGHBRIDGECODEIN, WEIGHTIN, DATEOUT, " & _
														" WBRIDGE_WEIGHBRIDGECODEOUT, WEIGHTOUT, " & _
														" FCENTRY, FCEDIT, FCIP, FCBA,TIMEIN,TIMEOUT,DIVISION,DRIVERNAME ) " & _
														" values ('2','" &  CHITNO & "','" &  BRG & "','" &  rsAsal("nopolisi") & "','10001',''," & _
														" '',to_Date('" & rsAsal("dnote_date") & "','MM/DD/YYYY'),'WB01'," &  rsAsal("bruto") & ",to_Date('" & rsAsal("tanggalk") & "','MM/DD/YYYY')," & _
														" 'WB01'," &  rsAsal("tarra") & "," & _
														" '" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
														Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','" &  rsAsal("waktu1") & "','" &  rsAsal("waktu2") & "','','" &  rsAsal("DRIVERNAME") & "')"
														connStr.execute(q)	
											 		   ' response.write  q & "<br>"
													'insert ke Delivery Note
												    q=  "Insert into DELIVERYNOTE (DNOTE_DATE, DELIVERYNOTECODE, CONTRACTCODE, GATEPASS, WBT_CHITNO, STORAGECODE, QUANTITY, FFA, DIRT, MOISTURE, TRANSPORTCO, DRIVERNAME, ICNO, FCENTRY, FCEDIT, FCIP, FCBA) " & _
														"values (to_Date('" & rsAsal("DNOTE_DATE") & "','MM/DD/YYYY'),'" &  rsAsal("DELIVERYNOTECODE") & "','" &  rsAsal("CONTRACTCODE") & "','" & rsAsal("GATEPASS") & "','"& CHITNO &"'," & _
														"'" &  rsAsal("STORAGECODE") & "','" &  rsAsal("QUANTITY") & "','" &  rsAsal("FFA") & "','" &  rsAsal("DIRT") & "','" &  rsAsal("MOISTURE") & "','" & rsAsal("TRANSPORTCO") & "','" &  rsAsal("DRIVERNAME") & "','" &  rsAsal("ICNO") & "'," & _
														"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
														Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
														'Response.write q & "<br>"
														connStr.execute(q)	
													q = " update wbst_cpo t set t.nochit = '" &  CHITNO & "', t.error_flag ='Success' ,t.error_desc ='' where to_char(t.tanggalm,'YYYYMMDD') = '" & rsAsal("tanggalm") & "' and t.sim  = '" & rsAsal("ICNO") & "'  and t.nopolisi ='" & rsAsal("nopolisi") & "' " 
												   ' Response.write q
													connStr.execute(q) 
										    else
											        ' jika Quota sdh terpenuhi
													 q = " update WBST_CPO  t set t.error_flag ='Not Success' ,t.error_desc ='No Contract : " & CONTRACTCODE & " ,  sdh Over Quota!!. ' where to_char(t.tanggalm,'YYYYMMDD') = '" & rsAsal("tanggalm") & "' and t.sim  = '" & rsAsal("ICNO") & "'  and t.nopolisi '" & rsAsal("nopolisi") & "' " 
													' connStr.execute(q)	
													 Response.Write "No Contract : " & CONTRACTCODE & " , tgl "&  rsAsal("dnote_date") &  " , dgn No Polisi " & rsAsal("nopolisi") & " <blink> ,  Over Quota!!. </blink> "& "<br>"	

                                            end if											
									   else 
                                             ' jika sudah di input ada no contract,tgl di IPlaS
											 'q = " update WBST_CPO  t set t.error_flag ='Not Success' ,t.error_desc ='No Contract : " & CONTRACTCODE & " , tgl "& rsasal("dnote_date") & " , dgn No Polisi " & rsAsal("nopolisi") & " sdh ada di IPlaS !!. ' where to_char(t.tanggalm,'YYYYMMDD') = '" & rsAsal("tanggalm") & "' and t.sim  = '" & rsAsal("ICNO") & "'  and t.nopolisi '" & rsAsal("nopolisi") & "' " 
										     'connStr.execute(q)	
											 Response.Write "( INSERT DATA GAGAL !!) No Contract : " & CONTRACTCODE & " , tanggal "& rsasal("dnote_date") & " , dgn No Polisi " & rsAsal("nopolisi") & " sudah diinput di IPlaS !!. "& "<br>"	
                                       end if
									   
                                       'Responsee.Write  q 		
								 else
                                      'jika tidak ada no contract di IPlaS
									   'q = " update WBST_CPO  t set t.error_flag ='Not Success' ,t.error_desc ='No Contract : " & CONTRACTCODE & " , tidak ditemukan di IPlaS !!. ' where to_char(t.tanggalm,'YYYYMMDD') = '" & rsAsal("tanggalm") & "' and t.sim  = '" & rsAsal("ICNO") & "'  and t.nopolisi '" & rsAsal("nopolisi") & "' " 
									   'connStr.execute(q)	
									   Response.Write "No Contract : " & CONTRACTCODE & " , tidak ditemukan di IPlaS !!. "& "<br>"							 
								 end if 
								 
						     rsAsal.MoveNext
							Wend
						end if
						rsAsal.Close
						set rsAsal = Nothing
						
					    TimeAkhir=Time()
						Lama =TimeAkhir-TimeAwal
						Response.Write "<br>Finish ... " & time() & "<br><br>"
						Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."
				    END IF
					%>
					
					</form>
			</td>
        </tr>		 
		<tr>
		  <td colspan="2"> 
		  <!--#include file="bottom.jsp"-->
		  </td>
		</tr>
    </table>

</body>
</html>