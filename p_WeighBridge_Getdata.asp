
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
				Dim strSQLPreParameter, rsSQLPreParameter,v_txtFromDayDate, v_txtToDayDate, v_txtfcba, v_txtyear, v_txtmonth
				dim  cmdView
				v_txtFromDayDate = Request("txtFromDayDate")
				v_txtToDayDate   = Request("txtToDayDate")
               	cmdView = request("cmdView")			
				%>
					<form name="frmEntry" method="post" action="p_WeighBridge_Getdata.Asp">
						<table class="data" width="80%" cellpadding="0" cellspacing="0">
						            <tr>
										<th>&nbsp;</th>
										<th>&nbsp;</th>
									</tr>
						            <tr>
										<h2>Ambil Data TBS dari WeighBridge ke IPlaS ( Grading Record dan Weigh Bridge Ticket )</h2>
										
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
						Dim  QC, strSQL,rsAsal,strSQLwb,rswb,q,strSQLCust,rscust
						Dim TimeAwal,TimeAkhir,Hitung,Lama,CHITNO
						TimeAwal=Time()
						Dim Startingtime, Endingtime
						Startingtime = FormatDateTime(CDate(Now))
					%>
						Please Wait...<br><br>
					<%
						
						
						Response.Write "Collecting Data From Date ..." & v_txtFromDayDate & " To.. " & v_txtToDayDate & "<br>" 
						Response.Write "Start ..." & Time() & "<br>" 
						Response.Write "Insert ke WEIGHBRIDGETICKET dan GRADINGRECORD " & "<br>" 
						strSQL = " select to_number(t.noseri) noseri,replace(t.nopolisi,'''',' ') nopolisi,replace(t.supir ,'''',' ') drivername,t.kd_rls,t.tanggalm,t.tanggalm,T.WAKTU1 TIMEIN,t.tanggalk,T.WAKTU2  TIMEOUT,t.bruto weightin ,t.tarra weightout,(t.bruto - t.tarra) netto1,round((t.f00 + t.f5 + t.tkosong + t.tpanjang + t.dura + t.bkecil) * (t.bruto - t.tarra)/100,0)  persen1 ,round((t.air + t.pasir + t.sampah + t.lain),0) tambah1,janjang bunches, " &_
						         " F00, F5,TKOSONG,TPANJANG, DURA,BKECIL,AIR,SAMPAH,PASIR,LAIN from wbst_tbs t where t.tanggalm between to_date('" & v_txtFromDayDate & " ','MM/DD/YYYY') and to_date('" & v_txtToDayDate & "','MM/DD/YYYY')  and   t.kd_brg in ('TBS','FFB') and t.tarra is not null and to_number(t.noseri) not in (select to_number(substr(w.chitno,-5)) from weighbridgeticket w where substr(w.chitno,-1) not in ('A','B'))"
						'Response.Write strSQL  & "<br><br>"
						Set rsAsal = Server.CreateObject("ADODB.Recordset")
						rsAsal.CursorType = 3
						rsAsal.CursorLocation = 3
						rsAsal.Open strSQL, connstr
						
						if not rsAsal.EOF then
			               while not rsAsal.EOF
							 CHITNO = rsAsal("noseri")
								 'Response.Write CHITNO & ", " 
								 'check apakah nochit sdh ada di IPlaS
								 strSQLwb = "select t.chitno from weighbridgeticket t where t.chitno = '" & CHITNO &"'"
								 set rswb = Server.CreateObject("ADODB.Recordset")
								 rswb.Open strSQLwb, connstr
								' Response.Write strSQLwb & "<br>" 
								 if rswb.eof then
								       strSQLCust = "select t.kdcust,t.devision  from WBST_RELASI t where t.kode = '" &  rsAsal("kd_rls") & "' and t.kdcust is not null "
									   set rscust   = Server.CreateObject("ADODB.Recordset")
									   'Response.Write strSQLCust & "<br>" 
								       rscust.Open strSQLCust, connstr
									   if not rscust.eof then
									                Response.Write "No Chit : " & CHITNO & ",Kode Customer = " & rsAsal("kd_rls") & " = "& rscust("kdcust") & " afdeling " & rscust("devision")  &" berhasil di Insert " & "<br>" 
													'insert ke WEIGHBRIDGETICKET
													
												  q =   " Insert into WEIGHBRIDGETICKET (TRANSACTIONTYPE, CHITNO, PRODUCT_PRODUCTCODE, VEHICLENO, CUST_CUSTOMERCODE, FIELD_ESTATE_ESTATECODE, FIELD_FIELDCODE, DATEIN, WBRIDGE_WEIGHBRIDGECODEIN, WEIGHTIN, DATEOUT, WBRIDGE_WEIGHBRIDGECODEOUT, WEIGHTOUT, FCENTRY, FCEDIT, FCIP, FCBA ,TIMEIN,TIMEOUT,DIVISION,DRIVERNAME) " & _
														 " values ('1','TBS" &  CHITNO & "','FFB','" &  rsAsal("nopolisi") & "','" &  rscust("kdcust") & "','',"   & _
														 " '',to_Date('" & rsAsal("tanggalm") & "','MM/DD/YYYY'),'WB01'," &  rsAsal("weightin") & ",to_Date('" & rsAsal("tanggalk") & "','MM/DD/YYYY'),"  & _
													 	 " 'WB01'," &  rsAsal("weightout") & ","  & _
													 	 " '" & Session("CodeName") & "','" & Session("CodeName") & "','"  & _
													 Request.ServerVariables("REMOTE_ADDR") & "','" & Session("fcBA") & "','" &  rsAsal("TIMEIN") & "','" &  rsAsal("TIMEOUT") & "','" &  rscust("devision") & "','" &  rsAsal("DRIVERNAME") & "')"
													 connStr.execute(q)
                                                    'Response.Write q & "<br>" 														
													'insert ke GRADINGRECORD
												  q =   " Insert into GRADINGRECORD (GRADINGNO, DATETIME, VEHICLENO, WEIGHT, BUNCHES, REMARKS, CHECKEDBY, WBT_CHITNO, FCENTRY, FCEDIT, FCIP, FCBA,F00, F5,TKOSONG,TPANJANG, DURA,BKECIL,AIR,SAMPAH,PASIR,LAIN) " & _
														" values ('04',to_date('" &  rsAsal("tanggalm")& "','mm/dd/yyyy'),'" &  rsAsal("nopolisi") & "','" &  cdbl(rsAsal("persen1")) + cdbl(rsAsal("tambah1")) & "','" & rsAsal("bunches") & "','','','TBS" &  CHITNO & "'," & _
														" '" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
														Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','" &  rsAsal("F00") & "','" &  rsAsal("F5") & "','" &  rsAsal("TKOSONG") & "','" &  rsAsal("TPANJANG") & "','" &  rsAsal("DURA") & "','" &  rsAsal("BKECIL") & "','" &  rsAsal("AIR") & "','" &  rsAsal("SAMPAH") & "','" &  rsAsal("PASIR") & "','" &  rsAsal("LAIN") & "')"
													connStr.execute(q)	
													'Response.Write q & "<br>"
												  q = " update wbst_tbs t set t.nochit = 'TBS" &  CHITNO & "', t.error_flag ='Success' ,t.error_desc ='' where to_number(t.noseri) = "&  CHITNO 
												  connStr.execute(q)
									   else 
                                                  q = " update wbst_tbs t set t.error_flag ='Not Success' ,t.error_desc ='Kode Customer tidak ditemukan di iplas' where to_number(t.noseri)  = "&  CHITNO 
												  connStr.execute(q)	
												  
												  Response.Write "No Chit : " & CHITNO & ",Kode Customer " & rsAsal("kd_rls") & " , Kode Customer tidak ditemukan di iplas "& "<br>"
                                       end if									   
                                       'Response.Write  q 		
									   
								 end if 
							  'message = "This is my message"
                              'Response.Write("<script language=VBScript>MsgBox """ + "This is my message" + """</script>")	 
							  'response.write("<script language=""javascript"">alert('Hello!');</script>") 
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