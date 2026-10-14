<%@ Language=VBScript %>
<%Option Explicit%>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim rsModule, strModule

Dim v_txtDate
v_txtDate		=	trim(request("txtDate"))

strModule = "SELECT CODE, NAME FROM MODULE WHERE ACTIVE = 'Y'  AND CODE NOT IN ('17','31','1','33','30','15','16','32') ORDER BY SEQUENCE ASC"

Set rsModule = Server.CreateObject("ADODB.Recordset")
rsModule.Open strModule, connstr, adOpenStatic 
%>
<html>
	<head>
		<script language="JavaScript" src="/js/date-picker.js"></script> 
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
<table id="form-table" style="margin: 0 auto">
	<tr>
		<th>TANGGAL</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post">
		<td align="center">
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
			<input type="text" maxlength="10" name="txtDate" size="10" value="<%=v_txtDate%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtDate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtDate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

			<table border="1" cellpadding="0" cellspacing="0" width="600PX"  style="margin: 0 auto">
				<tr>
					<th align="center" bgcolor="F0F0B0">NO</th>
					<td>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<td>
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<th width="10%" align="center" bgcolor="F0F0B0">&nbsp;</th>
										<th width="70%" align="center" bgcolor="F0F0B0">MODULE</th>
										<th align="center" bgcolor="F0F0B0">STATUS</th>
									</tr>
								</table>
							</td>
						</table>
					</td>
				</tr>
				<%
					if not rsModule.EOF then
						while not rsModule.EOF
				%>
							<tr>
								<td align="center" bgcolor="F0F0B0">&nbsp;</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td width="5%" align="center" bgcolor="F0F0B0">&nbsp;</td>
											<td align="left" bgcolor="F0F0B0"><%=rsModule("NAME")%></td>
										</tr>
									</table>
								</td>
							</tr>
								<%
									Dim rsSubModule, strSubModule
									strSubModule = 	" SELECT CODE, NAME, MODULE FROM SUBMODULE WHERE ACTIVE = 'Y' AND" & _
													" MODULE ='" & rsModule("CODE") & "' AND" & _
													" NAME NOT LIKE '%Closing%' AND NAME NOT LIKE '%Ledger%' AND" & _
													" NAME NOT LIKE '%Transfer%' AND NAME NOT LIKE '%Seal%' AND" & _
													" NAME NOT LIKE '%Process%' AND NAME NOT LIKE '%WeighBridge%' AND" & _
													" NAME NOT LIKE '%Master%' AND NAME NOT LIKE '%Adjustment%' AND" & _
													" NAME NOT LIKE '%Return%' AND NAME NOT LIKE '%Product%' AND" & _
													" NAME NOT LIKE '%Customer%' ORDER BY SEQUENCE ASC"

									Set rsSubModule = Server.CreateObject("ADODB.Recordset")
									rsSubModule.Open strSubModule, connstr, adOpenStatic
									
									if not rsSubModule.EOF then
										Dim NO
										NO = 1
										while not rsSubModule.EOF
								%>
									<tr>
										<td align="center" valign="top"><%=NO%></td>										
											<%
												Dim rsDetail, strDetail, rsStatusDetail, strStatusDetail, rsStatus, strStatus
												if rsSubModule("MODULE") = "3" and rsSubModule("CODE") = "11" then
													strDetail = " SELECT FCCODE FROM GANG WHERE FCBA = '" & Session("Fcba") & "' ORDER BY FCCODE"
												elseif rsSubModule("MODULE") = "9" and rsSubModule("CODE") = "11" then
													strDetail = " SELECT FCCODE FROM VEHICLE WHERE FCBA = '" & Session("Fcba") & "' ORDER BY FCCODE"
												else
													strDetail = " SELECT '' FCCODE FROM DUAL"
												end if
												
												Set rsDetail = Server.CreateObject("ADODB.Recordset")
												rsDetail.Open strDetail, connstr, adOpenStatic
												if not rsDetail.EOF then
													if rsDetail("FCCODE") <> "" then
											%>
										<td>
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
												<tr>
													<td width="10%" align="center" bgcolor="EBEA95">&nbsp;</td>
													<td width="70%" align="left" bgcolor="EBEA95"><font size="2" color="#0E0D03"><%=rsSubModule("NAME")%></font></td>
												</tr>
												<%
												Do While Not rsDetail.EOF
													if rsSubModule("MODULE") = "3" and rsSubModule("CODE") = "11" then
														strStatusDetail = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, GANGCODE, FDDATE FROM ATTENDANCE_GAD" & _
																	" WHERE FDDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, GANGCODE, FDDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "' AND GANGCODE = '" & rsDetail("FCCODE") & "'"
													elseif rsSubModule("MODULE") = "9" and rsSubModule("CODE") = "11" then
														strStatusDetail = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, VEHICLECODE, VEHICLEDATE FROM VEHICLEACTIVITY" & _
																	" WHERE VEHICLEDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, VEHICLECODE, VEHICLEDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "' AND VEHICLECODE = '" & rsDetail("FCCODE") & "'"
													else
														strStatusDetail =	" SELECT 'No Data' STATUS FROM DUAL"
													end if
													Set rsStatusDetail = Server.CreateObject("ADODB.Recordset")
													rsStatusDetail.Open strStatusDetail, connstr, adOpenStatic
													if not rsStatusDetail.EOF then
														Do While Not rsStatusDetail.EOF

												%>
												<tr>
													<td width="10%" align="center" bgcolor="EBEA95">&nbsp;</td>
													<td>
														<table border="0" cellpadding="0" cellspacing="0" width="100%">
															<tr>
																<td width="10%" align="center" bgcolor="EBEA95">&nbsp;</td>
																<td width="70%" align="left" bgcolor="EBEA95"><font size="2" color="#0E0D03"><%=rsDetail("FCCODE")%></font></td>
																<td align="center" bgcolor="EBEA95"><%=rsStatusDetail("STATUS")%></td>
															</tr>
														</table>
													</td>													
												</tr>
												<%
															rsStatusDetail.MoveNext
														loop
													end if
												rsDetail.MoveNext%>
												<%loop%>
											</table>
										</td>
												<% else %>
										<%
													if rsSubModule("MODULE") = "14" and rsSubModule("CODE") = "11" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, DATECREATED FROM PAYMENTVOUCHER" & _
																	" WHERE DATECREATED = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, DATECREATED" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "5" and rsSubModule("CODE") = "11" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, PODATE FROM PURCHASEORDER" & _
																	" WHERE PODATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, PODATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "5" and rsSubModule("CODE") = "12" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, RNDATE FROM RECEIVENOTE" & _
																	" WHERE RNDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, RNDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "5" and rsSubModule("CODE") = "14" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, INVOICEDATE FROM INVOICE" & _
																	" WHERE INVOICEDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, INVOICEDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "6" and rsSubModule("CODE") = "54" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, SIVDATE FROM ISSUE" & _
																	" WHERE SIVDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, SIVDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "6" and rsSubModule("CODE") = "55" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, WODATE FROM WRITEOFF" & _
																	" WHERE WODATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, WODATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "3" and rsSubModule("CODE") = "12" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, ADHOCDATE FROM EMPALLDED" & _
																	" WHERE ADHOCDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, ADHOCDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "3" and rsSubModule("CODE") = "13" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, FDDATE FROM HARVESTINGQUALITY" & _
																	" WHERE FDDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, FDDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "8" and rsSubModule("CODE") = "11" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, WORKSHOPDATE FROM WORKSHOPACTIVITY" & _
																	" WHERE WORKSHOPDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, WORKSHOPDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "4" and rsSubModule("CODE") = "11" then
														strStatus = " SELECT 'No Query' STATUS FROM DUAL"
													elseif rsSubModule("MODULE") = "4" and rsSubModule("CODE") = "12" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, WIPDATE FROM WORKINPROGRESS" & _
																	" WHERE WIPDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, WIPDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "10" and rsSubModule("CODE") = "11" then
														strStatus = " SELECT 'No Query' STATUS FROM DUAL"
													elseif rsSubModule("MODULE") = "10" and rsSubModule("CODE") = "13" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, CULLINGDATE FROM NURSERYCULLING" & _
																	" WHERE CULLINGDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, CULLINGDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "10" and rsSubModule("CODE") = "14" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, ISSUEDATE FROM NURSERYISSUE" & _
																	" WHERE ISSUEDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, ISSUEDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "2" and rsSubModule("CODE") = "11" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, RECEPTIONDATE FROM HARVESTING" & _
																	" WHERE RECEPTIONDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, RECEPTIONDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "2" and rsSubModule("CODE") = "12" then
														strStatus = " SELECT 'No Query' STATUS FROM DUAL"
													elseif rsSubModule("MODULE") = "2" and rsSubModule("CODE") = "13" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, WIDATE FROM WEATHER" & _
																	" WHERE WIDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, WIDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "29" and rsSubModule("CODE") = "16" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, DATEIN FROM WEIGHBRIDGETICKET" & _
																	" WHERE DATEIN = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, DATEIN" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "29" and rsSubModule("CODE") = "17" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, DATETIME FROM GRADINGRECORD" & _
																	" WHERE DATETIME = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, DATETIME" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "23" and rsSubModule("CODE") = "12" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, PRODUCTIONDATE FROM SHIFTREPORT" & _
																	" WHERE PRODUCTIONDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, PRODUCTIONDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "23" and rsSubModule("CODE") = "16" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, READINGDATE FROM PRODUCTSTORAGEREADING" & _
																	" WHERE READINGDATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, READINGDATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													elseif rsSubModule("MODULE") = "22" and rsSubModule("CODE") = "11" then
														strStatus = " SELECT 'No Query' STATUS FROM DUAL"
													elseif rsSubModule("MODULE") = "22" and rsSubModule("CODE") = "13" then
														strStatus = " SELECT TRIM(DECODE(COUNT(*),1,'OK','No Data')) STATUS FROM(" & _
																	" SELECT FCBA, DNOTE_DATE FROM DELIVERYNOTE" & _
																	" WHERE DNOTE_DATE = TO_DATE('" & v_txtDate & "','MM/DD/YYYY')" & _
																	" GROUP BY FCBA, DNOTE_DATE" & _
																	" )" & _
																	" WHERE FCBA = '" & Session("Fcba") & "'"
													else
														strStatus =	" SELECT 'No Data' STATUS FROM DUAL"
													end if
													Set rsStatus = Server.CreateObject("ADODB.Recordset")
													rsStatus.Open strStatus, connstr, adOpenStatic
													if not rsStatus.EOF then
														Do While Not rsStatus.EOF
										%>
										<td>
											<table border="0" cellpadding="0" cellspacing="0" width="100%">
												<tr>
													<td width="10%" align="center" bgcolor="EBEA95">&nbsp;</td>
													<td width="70%" align="left" bgcolor="EBEA95"><font size="2" color="#0E0D03"><%=rsSubModule("NAME")%></font></td>
													<td align="center" bgcolor="EBEA95"><%=rsStatus("STATUS")%></td>
												</tr>
												<%
															rsStatus.MoveNext
														loop
													end if%>
											</table>
										</td>
									</tr>
									<% end if
									end if
									%>									
								<%
										NO = NO + 1
										rsSubModule.MoveNext
										wend
									end if
							rsModule.MoveNext
						Wend
					end if
					rsModule.Close
				%>
		</table>
	</body>
<html>