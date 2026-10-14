<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno,POSTINGFLG,totalpremi,bkmdate
dim ScriptName,btnApprove,bkmno1,SQL,fcbaview
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
bkmno		= trim(request("bkmno"))
businessunit= trim(session("fcba"))
fcbaview		= trim(request("fcba"))
'response.write bkmno

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if



' get data header
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,SECTIONNAME,GANGCODE,ROWSTATEID,SUPERVISI1,SUPERVISINAME1,Clasification,Clasificationname,ROWSTATE,VEHICLECODE,VEHICLENAME
sqlheader = "select h.BKMNO,TO_CHAR(h.BKMDATE,'dd/mm/yyyy')BKMDATE,h.GANGCODE,h.SUPERVISI1,(SELECT s1.FCNAME FROM EMPLOYEE  s1 WHERE h.SUPERVISI1=s1.FCCODE and s1.fcba=h.fcba ) AS SUPERVISINAME1, " & _
" h.SUPERVISI2,(SELECT s2.FCNAME FROM EMPLOYEE s2 WHERE h.SUPERVISI2=s2.FCCODE and s2.fcba=h.fcba) AS SUPERVISINAME2,h.SUPERVISI3, " & _
" (SELECT s3.FCNAME FROM EMPLOYEE s3 WHERE h.SUPERVISI3=s3.FCCODE and s3.fcba=h.fcba) AS SUPERVISINAME3,h.SUPERVISI4, " & _
" (SELECT s4.FCNAME FROM EMPLOYEE s4 WHERE h.SUPERVISI4=s4.FCCODE and s4.fcba=h.fcba) AS SUPERVISINAME4," & _ 
" h.SECTIONCODE,h.DOCUMENT_CLASSIFICATION,c.fcname, h.ROWSTATE,h.POSTINGFLG,h.VEHICLECODE,v.FCCODE || '-' || v.FCNAME VEHICLENAME " & _
" from bkmhdr h LEFT JOIN DOCUMENT_CLASSIFICATION c ON (h.DOCUMENT_CLASSIFICATION=c.FCCODE) " & _ 
" left join vehicle v on (h.VEHICLECODE=v.FCCODE and h.fcba=v.fcba) where h.BKMNO='" & bkmno & "'"
'response.write sqlheader
set rsHeader = connStr.Execute(sqlheader)

if not rsHeader.eof then 
	bkmnohdr = rsHeader("BKMNO")
	DATECREATED = rsHeader("BKMDATE")
	SECTIONNAME = rsHeader("SECTIONCODE")
	GANGCODE = rsHeader("GANGCODE")
	SUPERVISI1 = rsHeader("SUPERVISI1")
	
	SUPERVISINAME1 = rsHeader("SUPERVISINAME1")
	Clasification = rsHeader("DOCUMENT_CLASSIFICATION")
	Clasificationname 	= rsHeader("fcname")
	ROWSTATE = rsHeader("ROWSTATE")
	VEHICLECODE = rsHeader("VEHICLECODE")
	VEHICLENAME = rsHeader("VEHICLENAME")
	POSTINGFLG = rsHeader("POSTINGFLG")
	
end if

IF POSTINGFLG = "0" THEN 
   POSTINGFLG = "TIDAK"
ELSE
   POSTINGFLG	= "YA"
END IF

'count detailupah
Dim SQL_Rows, NoOfRows,RSCount

SQL_Rows = "SELECT Count (*) from BKMUPAHDTL " & _ 
" where BKMNO='" & bkmno & "'"
set RSCount = connStr.Execute(SQL_Rows)
if not RSCount.eof then 
	NoOfRows = RSCount(0).value
end if
RSCount.close

'count detailbahan
Dim SQL_Rowsitem, NoOfRowsitem,RSCountitem

SQL_Rowsitem = "SELECT Count (*) from BKMVEHICLEDTL where BKMNO='" & bkmno & "'"
set RSCountitem = connStr.Execute(SQL_Rowsitem)
if not RSCountitem.eof then 
	NoOfRowsitem = RSCountitem(0).value
end if
RSCountitem.close

'proses dibawah ini digunakan untuk approved bkm dan create jurnal
dim strQueryStatus,modulFile,executeQry,modulFile1,executeQry1,bkmtype,sectionnameview,gangcodeview,vhcode
btnApprove 	= request("btnsave")
bkmno1 		= request("bkmno1")
bkmdate 	= request("bkmdate")
bkmtype 	= request("bkmtype")
sectionnameview 	= request("sectionname")
gangcodeview 	= request("gangcode")
vhcode 		= request("vhcode")

if btnApprove = "Replanned" then 
	fcbaview		= trim(request("fcbaview"))

	strQueryStatus = "UPDATE bkmhdr SET rowstate = 'Planned', postingflg='0',fcedit='" & Session("CodeName") & "',lastupdate=sysdate WHERE fcba = '" & fcbaview & "' AND bkmno = '" & bkmno1 & "'"
	connStr.execute(strQueryStatus)
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & strQueryStatus
	ExecuteAndLog  modulFile,executeQry
	
	 q = " INSERT INTO BKMREPLANNEDHISTORI ( BKMNO" & _
                            "                                           , CREATEDATE" & _
                            "                                           , BKMDATE" & _
                            "                                           , BKMTYPE" & _
                            "                                           , SECTIONNAME" & _
                            "                                           , GANG" & _
                            "                                           , VEHICLECODE" & _
                            "                                           , LASTUPDATE" & _
                            "                                           , FCBA" & _
                            "                                           , FCIP" & _
                            "                                           , FCENTRY, LASTTIME)"& _
                            " VALUES ('" & bkmno1 & "', " & _ 
                            " sysdate, " & _
                            " TO_DATE('" & bkmdate & "', 'dd/mm/yyyy'), " & _
                            " '" & bkmtype & "', " & _
                            " '" & sectionnameview & "', " & _
                            " '" & gangcodeview & "', " & _
                            " '" & vhcode & "', " & _
                            " sysdate, " & _
                            " '" & fcbaview & "', " & _
                            " '" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
                            " '" & Session("CodeName") & "', " & _
                            " TO_CHAR(SYSDATE, 'hh24:mi'))"
                   ''   response.write q
                    '  Response.end
                        connStr.execute(q)


	response.redirect("bkm_traksi_replan_view.asp?bkmno="&bkmno1&"&fcba="&fcbaview)


end if
%>

<html>
	<head>
		<title>FBPlants - View BKM TRAKSI</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
	<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>

<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: -15px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarhk{
		width: 45px;
	}
	.lebarabsen{
		width: 48px;
	}
	.lebarlokasi{
		width: 60px;
	}
	.lebartc{
		width: 70px;
	}
	.lebarrate{
		width: 60px;
	}
	.lebaractivity{
		width: 110px;
	}
	.lebarallow{
		width: 90px;
	}
	.lebarupah{
		width: 80px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakiri{
		text-align: left;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}
	.lebaritembarang{
		width: 400px;
	}
	.lebaractbarang{
		width: 300px;
	}
	.lebarvehicle{
		width: 150px;
	}
	.lebartime{
		width: 90px;
	}

	.lebarspkno{
		width: 250px;
	}
	.lebarovtime{
		width: 70px;
	}
	.warnaamount{
		background-color: #eee;
	}
	
</style>



<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-bkm"<% end if %>>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="">&nbsp;&nbsp;&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM TRAKSI VIEW</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_traksi_replan_view.asp">
	<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
	<input type="hidden" name="jumldatavh" id="jumldatavh" value="<%=NoOfRowsitem%>" />
	<input type="hidden" name="proses" id="proses" value="2" />
	<input type="hidden" name="bkmno1" id="bkmno1" value="<%=bkmno%>" />
	<input type="hidden" name="bkmdate" id="bkmdate" value="<%=DATECREATED%>" />
	<input type="hidden" name="bkmtype" id="bkmtype" value="<%=Clasification%>" />
	<input type="hidden" name="sectionname" id="sectionname" value="<%=SECTIONNAME%>" />
	<input type="hidden" name="gangcode" id="gangcode" value="<%=GANGCODE%>" />
	<input type="hidden" name="fcbaview" id="fcbaview" value="<%=fcbaview%>" />
	<input type="hidden" name="vhcode" id="vhcode" value="<%=VEHICLECODE%>" />
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >BKM No</td>
			<td align="left">
				<%=bkmno%>
				
			</td>


			<td align="left" >Section Name</td>
			<td align="left">
				<%=SECTIONNAME%>
				
			</td>
			<td align="left" >Mandor </td>
			<td align="left">
				<select name="supervisi" id="supervisi" class="lebarsection" >
					<option value="<%=SUPERVISI1%>"><%=SUPERVISINAME1%></option>
				</select>
			</td>
			
		</tr>
		<tr>
			<td align="left" >Tanggal</td>
			<td align="left">
				<%=DATECREATED%>
			</td>
			
			<td align="left" >Gang</td>
			<td align="left">
				<%=GANGCODE%>
			</td>
			<td align="left" >Kendaraan</td>
			<td align="left">
				<%=VEHICLECODE%>&nbsp;<%=VEHICLENAME%>
			</td>
		</tr>
		<tr>
			<td align="left" >Posting</td>
			<td align="left"><%=POSTINGFLG%>-<%=ROWSTATE%></td>
			

		</tr>
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='900' id='tableall'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH</li>
						<li class="TabbedPanelsTab">BKM VEHICLE/WORKSHOP</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm upah-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="600" class='data' id='table2'>
								<tr id="trheader2">
									<th>Employee</th>
									<th>TimeIn</th>
									<th>TimeOut</th>
									<th>Lokasi</th>
									<th>Absen</th>
									<th>Activity</th>
									<th>TC</th>
									<th>HK</th>
									<th>Unit</th>
									<th>Output</th>
									<th>Rate</th>
									<th>Allowed Code</th>
									<th>Rate Allowed</th>
									<th>OVTH</th>
									<th>TOVTime</th>
									<th colspan="2">&nbsp;</th>
								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,ttlhk,ttlunit,ttloutput,ttlrate,ttlrateall,ttlovth
								sqldtlupah = "SELECT u.BKMNO,u.fcba, u.chargeunit,u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
								" u.ALLOWANCE_CODE, nvl(u.ALLOWANCE_AMT,0)ALLOWANCE_AMT, u.ATTENDANCE,nvl(u.OVERTIME_HOURS,0)OVERTIME_HOURS,u.TYPE_OVERTIME,o.fcname type_ovname," & _ 
								" to_char(u.IN_FIN,'hh24:mi:ss')IN_FINDEPAN,to_char(u.OUT_FIN,'hh24:mi:ss')OUT_FINDEPAN, " & _
								" to_char(u.IN_FIN,'dd/mm/yyyy hh24:mi:ss')IN_FIN,to_char(u.OUT_FIN,'dd/mm/yyyy hh24:mi:ss')OUT_FIN,  " & _
								" u.LOCATIONCODE, nvl(u.OUTPUT,0)OUTPUT,nvl(u.rate,0)rate, nvl(u.UNIT,0)UNIT,al.FCCODE || ' - ' || al.fcname as ALLOWDEDNAME,u.OTHCONVERTION " & _
								" FROM BKMUPAHDTL u " & _ 
								" INNER JOIN employee e ON (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
								" INNER JOIN ACTMASTER a ON (u.ACTCODE=a.FCCODE) " & _
								" LEFT JOIN ALLOWDED al ON (al.FCCODE = u.ALLOWANCE_CODE and al.fcba=u.fcba) " & _ 
								" LEFT JOIN overtime_convertion o ON (o.type=u.TYPE_OVERTIME and o.HOURS=u.OVERTIME_HOURS and o.fcba=u.fcba)" & _
								" where u.BKMNO='" & bkmno & "' AND u.ATTENDANCE not in('KP') order by e.fcname, a.fcname,u.LOCATIONCODE,u.ALLOWANCE_CODE"
								set rsdtlupah = connStr.Execute(sqldtlupah)
							''	response.write sqldtlupah
								

																
								no = 0
								ttlhk = 0
								ttlunit = 0
								ttloutput = 0
								ttlrate = 0
								ttlrateall = 0
								ttlovth = 0
								totalamount = 0
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("fcname")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("IN_FINDEPAN")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("OUT_FINDEPAN")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("chargeunit")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("ATTENDANCE")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("activityname")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("LOCATIONCODE")%>
									</td>
									<td>
										<%=rsdtlupah("MANDAYS")%>
									</td>
									
									<td>
										<%=rsdtlupah("UNIT")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("OUTPUT")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("rate")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("ALLOWDEDNAME")%>
										
									</td>
																		
									<td class="ratakiri">
										<%=FormatNumber(rsdtlupah("ALLOWANCE_AMT"),0)%>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("OVERTIME_HOURS")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("type_ovname")%>									
									</td>
									
									
									<td></td>
									
								</tr>
								<%
									'dim ttlhk
									no = no +1
									ttlhk = cdbl(ttlhk) + cdbl(rsdtlupah("MANDAYS"))
									ttlunit = cdbl(ttlunit) + cdbl(rsdtlupah("unit"))
									ttloutput = cdbl(ttloutput) + cdbl(rsdtlupah("output"))
									ttlrate = cdbl(ttlrate) + cdbl(rsdtlupah("rate"))
									ttlrateall = cdbl(ttlrateall) + cdbl(rsdtlupah("ALLOWANCE_AMT"))
									ttlovth = cdbl(ttlovth) + cdbl(rsdtlupah("OVERTIME_HOURS"))
									
									
									rsdtlupah.MoveNext
									Wend
								%>
									<tr class="ratakanan">
										<td colspan="7" class="tebal">Total</td>
										<td class="tebal"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlhk,2)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlunit,2)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttloutput,2)%>"/></td>
										<td class="tebal ratakiri"><input class="ratatengah lebarrate" type="text" readonly  style="background-color:#eee;" name="ttlhk" id="ttlhk<%=no%>" value="<%=FormatNumber(ttlrate,2)%>"/></td>
										<td></td>
										<td class="tebal ratakiri"><input class="ratakanan lebarrate" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlrateall,0)%>"/></td>
										<td class="tebal ratatengah"><input class="ratatengah lebarhk" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(ttlovth,0)%>"/></td>
										<td></td>
										
									</tr>
							</table>
						</div>

					
						<!-- view bkm vehicle-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="1050" class='data' id='table3'>
									<tr id="trheader3">
										<th>Opr/Driver</th>
										<th>Helper</th>
										<th>Lokasi</th>
										<th>Time</th>
										<th>Activity</th>
										<th>TC</th>
										<th>Category Transport</th>
										<th>Trip</th>
										<th>Hours</th>
										<th>KM/HM</th>
										<th>Output</th>
										<th>Remark</th>
										<th colspan="2">&nbsp;</th>
									</tr>

									<%
									
										dim sqlvehicle,rsvehicle
										sqlvehicle = "Select bv.EMPLOYEECODE, e.FCNAME || '-' || e.FCCODE EMPLOYEENAME, " & _ 
										" bv.HELPERCODE,(SELECT hp.FCNAME FROM EMPLOYEE hp WHERE bv.HELPERCODE=hp.FCCODE and hp.fcba=bv.fcba) AS HELPERNAME, " & _
										" bv.CHARGEUNIT, bv.VEHICLECODE,v.FCCODE || ' - ' || v.FCNAME VEHICLENAME, bv.ACTCODE,c.FCNAME || ' - ' || c.FCCODE COANAME, " & _ 
										" bv.LOCATIONCODE,bv.CATEGORYTRANSPORT,t.FCCODE || ' - ' || t.FCNAME TRANSPORTKATNAME, " & _
										" bv.TIME, bv.HOURS, bv.UNIT, bv.OUTPUT, bv.TRIP,bv.remarks,bv.fcba " & _
										" From   BKMVEHICLEDTL bv " & _
										" INNER JOIN EMPLOYEE e on (e.FCCODE=bv.EMPLOYEECODE and e.fcba=bv.fcba)" & _
										" INNER JOIN ACTMASTER c on (bv.ACTCODE=c.FCCODE) " & _
										" LEFT JOIN VEHICLEACTIVITYCATEGORY t on (bv.CATEGORYTRANSPORT=t.FCCODE) " & _
										" INNER JOIN VEHICLE v on (v.FCCODE=bv.VEHICLECODE AND v.FCBA=bv.FCBA) " & _
										" WHERE bv.BKMNO='" & bkmno & "' order by bv.VEHICLECODE asc "
										'response.write sqlvehicle
										 set rsvehicle = connStr.Execute(sqlvehicle)
									%>

									<% 
										dim novh
										novh = 0
										While Not rsvehicle.EOF 
									%>
										<tr id='trv<%=novh%>'>
											<td class="ratakiri">
											<%=rsvehicle("EMPLOYEENAME")%>
											</td>
											<td class="ratakiri">
												<%=rsvehicle("HELPERNAME")%>
											</td>
											
											<td class="ratakiri">
												<%=rsvehicle("CHARGEUNIT")%>
											</td>
											<td class="ratatengah">
												<%=rsvehicle("TIME")%>
											</td>

											<td class="ratakiri">
												<%=rsvehicle("ACTCODE")%>&nbsp;<%=rsvehicle("COANAME")%>
											</td>

											<td class="ratatengah">
												<%=rsvehicle("LOCATIONCODE")%>
											</td>

											<td class="ratakiri">
												<%=rsvehicle("CATEGORYTRANSPORT")%>&nbsp;<%=rsvehicle("TRANSPORTKATNAME")%>
												
											</td>

											<td class="ratatengah">
												<%=rsvehicle("trip")%>
											</td>

											<td class="ratatengah">
												<%=rsvehicle("HOURS")%>
											</td>


											<td class="ratatengah">
												<%=rsvehicle("UNIT")%>
											</td>

											<td class="ratatengah">
												<%=rsvehicle("OUTPUT")%>
											</td>

											<td class="ratakiri">
												<%=rsvehicle("remarks")%>
											</td>

											
											<td></td>
										</tr>

											
										
									<%
										novh = novh + 1
										rsvehicle.MoveNext
										Wend
									
									%>
									
								</table>
							</div>
							
						</div>


						
					</div>
				</div>
			</td>
		</tr>
		<tr class="ratatengah">			
		<td colspan="9" align="center" >
		<% if ROWSTATE="Approved" Then %>
			<input type=Submit name="btnsave" id="btnsave" value="Replanned">&nbsp;&nbsp;
		<%end if%>
		    <input type=button onclick="window.close();" name="btnclose" value="CLOSE">
		</td>
		</tr>
	</table>
	<br>
	<table align=center border="0" class='data ratakiri' width='1100' id='tbldtl'>
		<tr>
			<th>Employee</th>
			<th>Lokasi</th>
			<th>Absen</th>
			<th>Activity</th>
			<th>TC</th>
			<th>HK</th>
			<th>Unit</th>
			<th>Output</th>
			<th>Rate</th>
			<th>Allowed Code</th>
			<th>Rate Allowed</th>
			<th>OVTH</th>
			<th>Remark</th>			
			<th>Amount</th>
			
		</tr>
		<%
			'get data detail upah
			dim rsupah,allcode,rateallow,hkview, dtlamt
			sqldtlupah = "SELECT u.BKMNO,u.fcba, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
			" p.ALLOWDEDCODE, nvl(u.ALLOWANCE_AMT,0)ALLOWANCE_AMT, u.ATTENDANCE,nvl(u.OVERTIME_HOURS,0)OVERTIME_HOURS,u.TYPE_OVERTIME,o.fcname type_ovname," & _ 
			" to_char(u.IN_FIN,'hh24:mi:ss')IN_FINDEPAN,to_char(u.OUT_FIN,'hh24:mi:ss')OUT_FINDEPAN, " & _
			" to_char(u.IN_FIN,'dd/mm/yyyy hh24:mi:ss')IN_FIN,to_char(u.OUT_FIN,'dd/mm/yyyy hh24:mi:ss')OUT_FIN,  " & _
			" u.LOCATIONCODE, nvl(u.OUTPUT,0)OUTPUT,nvl(u.rate,0)rate, nvl(u.UNIT,0)unit,p.ALLOWDEDCODE,al.FCCODE || ' - ' || al.fcname as ALLOWDEDNAME,nvl(p.AMOUNT,0)AMOUNT " & _
			" , u.remarks " & _ 
			" FROM BKMUPAHDTL u " & _ 
			" INNER JOIN employee e ON (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
			" INNER JOIN ACTMASTER a ON (u.ACTCODE=a.FCCODE) " & _
			" LEFT JOIN BKMPROSES  p ON (u.EMPLOYEECODE=p.EMPLOYEECODE AND p.BKMNO = u.BKMNO AND p.LINENOKEY = u.LINENOKEY AND u.fcba=p.fcba) " & _ 
			" LEFT JOIN ALLOWDED al ON (al.FCCODE = p.ALLOWDEDCODE and al.fcba=p.fcba) " & _ 
			" LEFT JOIN overtime_convertion o ON (o.type=u.TYPE_OVERTIME and o.HOURS=u.OVERTIME_HOURS and o.fcba=u.fcba)" & _
			" where u.BKMNO='" & bkmno & "' order by e.fcname, a.fcname,u.LOCATIONCODE,p.ALLOWDEDCODE"
			set rsupah = connStr.Execute(sqldtlupah)
			'response.write sqldtlupah
			totalamount = 0
			While Not rsupah.EOF

			allcode = rsupah("ALLOWDEDCODE")

			if allcode = "A01" THEN 
				rateallow = 0
				hkview    = rsupah("MANDAYS")
			ELSE 
				rateallow = rsupah("ALLOWANCE_AMT")
				hkview    = 0
			END IF 

			

		%>
		<tr>
			<td class="ratakiri"><%=rsupah("fcname")%></td>
			<td class="ratatengah"><%=rsupah("fcba")%></td>
			<td class="ratatengah"><%=rsupah("ATTENDANCE")%></td>
			<td class="ratakiri"><%=rsupah("activityname")%></td>
			<td class="ratatengah"><%=rsupah("LOCATIONCODE")%></td>
			<td class="ratatengah"><%=hkview%></td>
			<td class="ratatengah"><%=rsupah("UNIT")%></td>
			<td class="ratatengah"><%=rsupah("OUTPUT")%></td>
			<td class="ratatengah"><%=rsupah("RATE")%></td>
			<td class="ratakiri"><%=rsupah("ALLOWDEDNAME")%></td>
			<td class="ratakanan"><%=formatnumber(rateallow,0)%></td>
			<td class="ratatengah"><%=rsupah("OVERTIME_HOURS")%></td>
			<td class="ratakiri"><%=rsupah("remarks")%></td>
			<td class="ratakanan"><%=formatnumber(rsupah("AMOUNT"),0)%></td>
		</tr>
		<%
			totalamount = cdbl(totalamount) + cdbl(rsupah("AMOUNT"))
			rsupah.MoveNext
			Wend
		%>
		<tr>
			<td colspan="13" class="ratakanan tebal">Total</td>
			<td class="ratakanan tebal"><%=formatnumber(totalamount,0)%></td>
		</tr>
	</table>

</form>
</div>
</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

$(function() {	
	
	//$("#tglbkm").datepicker();
	//showSection();
	showVH();
	showSupervisi2();
	//addrowVH(1);
	
	
	getLoad();
	getLoadVH();
	//getjmlRow();

	var jumldatavh = $('#jumldatavh').val();

	if (jumldatavh=="0"){
		addrowVH(1);
	}
	
});



function numformat(val) {
	var nStr = val.toString();
	nStr = nStr.replace(/\,/g, "");
	var x = nStr.split('.');
	var x1 = x[0];
	var x2 = x.length > 1 ? '.' + x[1] : '';
	var rgx = /(\d+)(\d{3})/;
	while (rgx.test(x1)) {
		x1 = x1.replace(rgx, '$1' + '.' + '$2');
	}
	return x1 + x2;
}

//get lokasi upah
function showFCBA(selid) {
	var fcba = 'fcba';
	$('#lokasi').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp?viewdata=' + fcba,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get section
function showSection() {

	$('#sectionname').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsection.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get GANGCODE 
function showGang() {
	var sectionname = $("#sectionname").val();
	$('#gangcode').empty();
	$('#gangcode').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_gangcode.asp?sectionname=' + sectionname,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get supervisi 2
function showSupervisi2() {
	
	$('#supervisi').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get absen
function showAbsen(selid) {
	
	$('#absen').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get activity
function showActivity(selid) {
	
	$('#activity').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get tc
function showTC(selid) {
	
	$('#tcupah').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get allowedcode
function showAllowed(selid) {
	
	$('#allowedcode').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getallowed.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get employee
function showEmployee(selid) {
	var gangcode = $("#gangcode").val();
	$('#fccode').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_employee_edit.asp?gangcode=" + gangcode,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//getLoadUpah
function getLoad() {
		
	var bkmno = $("#bkmno").val();
	var upahedit = 2;
	$.post("bkm_employee_load.asp?upahedit=" + upahedit +"&bkmno=" + bkmno, function(data) {
		
		var json  = JSON.parse(data);
		var jumlview  = json.results.length;
		for (var i = 0; i < jumlview; i++) {

			showEmployee("fccode" + i);
			showFCBA("lokasi" + i);
			showAbsen("absen" + i);
			showActivity("activity" + i);
			showTC("tcupah" + i);
			showAllowed("allowedcode" + i);
			showTypeOv("typeovrtime" + i);
			
		}
	});
	
}


var jumlinv = $('#jumldata').val();
var StartIteration =  parseInt(jumlinv) + 1	
function addrowUpah(trid) {
	
	var lastrow = this.StartIteration++;
	

	if (lastrow < 2) {
		var row = $("#trheader2").after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#tru' + trid).after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowUpah(" + lastrow + ")' />";
	}

	var fccode = "<select name='fccode" + lastrow + "' id='fccode" + lastrow + "' class='ratakiri lebaractivity' onchange='getFinger(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + fccode + "</td>");

	var timein = "<span name='timein" + lastrow + "' id='timein" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeinx" + lastrow + "' id='timeinx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timein + "</td>");

	var timeout = "<span name='timeout" + lastrow + "' id='timeout" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeoutx" + lastrow + "' id='timeoutx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timeout + "</td>");

	var lokasi = "<select name='lokasi" + lastrow + "' id='lokasi" + lastrow + "' class='ratakiri lebarabsen'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah lebarabsen'>" + lokasi + "</td>");

	var absen = "<select name='absen" + lastrow + "' id='absen" + lastrow + "' class='ratakiri lebarabsen' onchange='getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah lebarabsen'>" + absen + "</td>");

	var activity = "<select name='activity" + lastrow + "' id='activity" + lastrow + "' class='ratakiri lebaractivity' onchange='getTC(" + lastrow + ");getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarabsen'>" + activity + "</td>");

	var blok = "<select type='text' name='tcupah" + lastrow + "' id='tcupah" + lastrow + "' value='0' class='ratatengah lebartc'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + blok + "</td>");

	var hk = "<input type='text' name='hk" + lastrow + "' id='hk" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarhk'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + hk + "</td>");

	var unit = "<input type='text' name='unit" + lastrow + "' id='unit" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + unit + "</td>");

	var output = "<input type='text' name='output" + lastrow + "' id='output" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + output + "</td>");

	var ratekb = "<input type='text' name='ratekb" + lastrow + "' id='ratekb" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarrate'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ratekb + "</td>");

	var allowedcode = "<select name='allowedcode" + lastrow + "' id='allowedcode" + lastrow + "' class='ratakiri lebarallow'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + allowedcode + "</td>");

	var ratepremi = "<input type='text' name='ratepremi" + lastrow + "' id='ratepremi" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratakanan lebarrate'>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + ratepremi + "</td>");

	var ovtimehours = "<input type='text' name='ovtimehours" + lastrow + "' id='ovtimehours" + lastrow + "' onkeyup='hanyaAngka(this);getOTHConvertion(" + lastrow + ")' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ovtimehours + "</td>");

	var typeovrtime = "<select type='text' name='typeovrtime" + lastrow + "' id='typeovrtime" + lastrow + "' onchange='getOTHConvertion(" + lastrow + ");' value='0' class='ratatengah lebarovtime'></select><input type='hidden' name='othconvertion" + lastrow + "' id='othconvertion" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + typeovrtime + "</td>");
	
	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowUpah(" + lastrow + ")' />";
	$('#tru' + lastrow).append("<td>" + tambahbutton + "</td>");

	$('#tru' + lastrow).append("<td>" + kurangbutton + "</td>");

	$('#jumldata').val(lastrow);

	showEmployee("fccode" + lastrow);
	showFCBA("lokasi" + lastrow);
	showAbsen("absen" + lastrow);
	showActivity("activity" + lastrow);
	showTC("tc" + lastrow);
	showAllowed("allowedcode" + lastrow);
	showTypeOv("typeovrtime" + lastrow);

	
}


function deleterowUpah(i) {
	Spry.Widget.Utils.destroyWidgets('fccode' + i);
	Spry.Widget.Utils.destroyWidgets('lokasi' + i);
	Spry.Widget.Utils.destroyWidgets('absen' + i);
	Spry.Widget.Utils.destroyWidgets('activity' + i);
	Spry.Widget.Utils.destroyWidgets('tcupah' + i);
	Spry.Widget.Utils.destroyWidgets('hk' + i);
	Spry.Widget.Utils.destroyWidgets('timein' + i);
	Spry.Widget.Utils.destroyWidgets('timeout' + i);
	Spry.Widget.Utils.destroyWidgets('timeinx' + i);
	Spry.Widget.Utils.destroyWidgets('timeoutx' + i);
	Spry.Widget.Utils.destroyWidgets('unit' + i);
	Spry.Widget.Utils.destroyWidgets('output' + i);
	Spry.Widget.Utils.destroyWidgets('ratekb' + i);
	Spry.Widget.Utils.destroyWidgets('allowedcode' + i);
	Spry.Widget.Utils.destroyWidgets('ratepremi' + i);
	Spry.Widget.Utils.destroyWidgets('ovtimehours' + i);
	Spry.Widget.Utils.destroyWidgets('typeovrtime' + i);
		
	 $('#tru' + i).remove();
	 //getdeljmlRow();
}

//show time from finger employee
function getFinger(i)
{
	var tglbkm = $('#tglbkm').val();
	var fccode = $('#fccode' + i).val();
	//alert(tglbkm);
	if (fccode != "") {
		$.post("bkm_employee_edit.asp?emp=" + 1 + "&tglbkm=" + tglbkm + "&fccode=" + fccode, function(data) {
			var json  = JSON.parse(data);
			//var jumlview  = json.results.length;
			var timein = json.results[0].timein;
			var timeout = json.results[0].timeout;
				$('#timein' + i).html(timein);
				$('#timeout' + i).html(timeout);
			
			
		});
		
	}

}


//get tc
function getTC(i)
{
	var activity = $('#activity' + i).val();
	var lokasi   = $('#lokasi' + i).val();

	$('#tcupah' + i).empty();
	$('#tcupah' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?activity=' + activity +'&lokasi=' + lokasi,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});

}


//function disabled input()
function getDisable(i)
{
	var absen = $('#absen' + i).val();
	if (absen=="KB"){
		//disbale
		$( '#hk' + i ).prop( "disabled", true );
		$( '#bunch' + i ).prop( "disabled", true );
		$( '#abw' + i ).prop( "disabled", true );
		$( '#kgbrondolan' + i ).prop( "disabled", true );
		$( '#allowedcode' + i ).prop( "disabled", true );
		
		//enable
		
		
	}else{
		$( '#hk' + i ).prop( "disabled", false );
		$( '#bunch' + i ).prop( "disabled", false );
		$( '#abw' + i ).prop( "disabled", false );
		$( '#kgbrondolan' + i ).prop( "disabled", false );
		$( '#unit' + i ).prop( "disabled", false );
		$( '#output' + i ).prop( "disabled", false );
		$( '#ratekb' + i ).prop( "disabled", false );
		$( '#allowedcode' + i ).prop( "disabled", false );
		
	}
	
}

//show tipe overtime
function showTypeOv(selid)
{
	
	//$('#typeovrtime' + i).empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_viewdata.asp?idview=' + 7,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});

}

//show time from finger employee
function getFinger(i)
{
	var tglbkm = $('#tglbkm').val();
	var fccode = $('#fccode' + i).val();
	//alert(tglbkm);
	if (fccode != "") {
		$.post("bkm_employee_edit.asp?emp=" + 1 + "&tglbkm=" + tglbkm + "&fccode=" + fccode, function(data) {
			var json  = JSON.parse(data);
			//var jumlview  = json.results.length;
			var timein = json.results[0].timein;
			var timeinx = json.results[0].timeinx;
			var timeout = json.results[0].timeout;
			var timeoutx = json.results[0].timeoutx;
				$('#timein' + i).html(timein);
				$('#timeout' + i).html(timeout);
				$('#timeinx' + i).val(timeinx);
				$('#timeoutx' + i).val(timeoutx);
			
			
		});
		
	}

}

//view data oth convertion
function getOTHConvertion(i)
{
	var ovtimehours = $("#ovtimehours" + i).val();
	var typeovrtime = $('#typeovrtime' + i).val();
	if (ovtimehours != "") {
		$.post("bkm_employee.asp?ovtimehours=" + ovtimehours + "&typeovrtime=" +  typeovrtime +  "&satuan=" + 7, function(data) {
			$('#othconvertion' + i).empty();
			var json  = JSON.parse(data);
			var othconvertion = json.results[0].convertion
			$('#othconvertion' + i).val(othconvertion);
						
		});
		
	}

	
}

function getIsibelakang(i)
{
	var ratepremidepan = $('#ratepremidepan' + i).val();
	if(ratepremidepan==""){
		$('#ratepremi' + i).val(0);
	}else{

		$('#ratepremi' + i).val(ratepremidepan);
	}
}

//show fcba vh
function showFCBAVH(selid) {
	var fcba = 'fcba';
	$('#lokasivh').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp?viewdata=' + fcba,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//show kendaraan
function showVH() {
	//$('#kendaraan').empty();
	var jmlitem = 2;
	$('#kendaraan').select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?jmlitem=" + jmlitem,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get tcvh
function showtcVH(selid) {
	
	$('#tcvh').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//getLoadvh
function getLoadVH() {
	
	var bkmno = $("#bkmno").val();
	var jmlitem = 3;
	$.post("bkm_show_item.asp?jmlitem=" + jmlitem +"&bkmno=" + bkmno, function(data) {
		var json  = JSON.parse(data);
		var jumlview  = json.results.length;
		
		for (var i = 0; i < jumlview; i++) {
			showEmployeeVH("employeecode" + i);
			showEmployeeVH("helpercode" + i);
			showFCBAVH("lokasivh" + i);
			showtcVH("tcvh" + i);
			//getTCVH("tcvh" + i);
			showActivityVH("activityivh" + i);
			showkategoriTR("kattransport" + i);
		}
	});
	
	
}



var jumldatavh = $('#jumldatavh').val();
//alert(jumldatavh);
if (jumldatavh=="0"){
		var StartIterationvh = 1
}else{
	var StartIterationvh =  parseInt(jumldatavh) + 1
}

		
function addrowVH(trid)
{

	var lstvh = this.StartIterationvh++;
	if (lstvh < 2) {
		var row = $("#trheader3").after("<tr id='trv" + lstvh + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trv' + trid).after("<tr id='trv" + lstvh + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowVH(" + lstvh + ")' />";
	}

	var employeecode = "<select name='employeecode" + lstvh + "' id='employeecode" + lstvh + "' class='lebarvehicle'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + employeecode + "</td>");

	var helpercode = "<select name='helpercode" + lstvh + "' id='helpercode" + lstvh + "' class='lebarvehicle'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + helpercode + "</td>");

	var lokasivh = "<select name='lokasivh" + lstvh + "' id='lokasivh" + lstvh + "' class='ratakiri lebarabsen'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri lebarabsen'>" + lokasivh + "</td>");

	var timestart = "<input type='text' name='timestart" + lstvh + "' id='timestart" + lstvh + "' value='' placeholder='hh:mm-hh:mm'  class='lebartime ratatengah'>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + timestart + "</td>");

	var activityivh = "<select name='activityivh" + lstvh + "' id='activityivh" + lstvh + "' class='lebaractivity' onchange='getTCVH(" + lstvh + ")'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + activityivh + "</td>");

	var tcvh = "<select name='tcvh" + lstvh + "' id='tcvh" + lstvh + "' class='lebartc ratatengah'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + tcvh + "</td>");

	var kattransport = "<select name='kattransport" + lstvh + "' id='kattransport" + lstvh + "' class='lebaractivity'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + kattransport + "</td>");

	var trip = "<input type='text' name='trip" + lstvh + "' id='trip" + lstvh + "' value='0' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trv' + lstvh).append("<td class='ratatengah'>" + trip + "</td>");

	var jamkerja = "<input type='text' name='jamkerja" + lstvh + "' id='jamkerja" + lstvh + "' value='0' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trv' + lstvh).append("<td class='ratatengah'>" + jamkerja + "</td>");

	var kmhm = "<input type='text' name='kmhm" + lstvh + "' id='kmhm" + lstvh + "' value='0' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trv' + lstvh).append("<td class='ratatengah'>" + kmhm + "</td>");

	var outputvh = "<input type='text' name='outputvh" + lstvh + "' id='outputvh" + lstvh + "' value='0' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trv' + lstvh).append("<td class='ratatengah'>" + outputvh + "</td>");

	var remarkvh = "<textarea name='remarkvh" + lstvh + "' id='remarkvh" + lstvh + "' rows='3' cols='40' class='lebaractivity'></textarea>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + remarkvh + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowVH(" + lstvh + ")' />";
	$('#trv' + lstvh).append("<td>" + tambahbutton + "</td>");

	$('#trv' + lstvh).append("<td>" + kurangbutton + "</td>");

	$('#jumldatavh').val(lstvh);

	showEmployeeVH("employeecode" + lstvh);
	showEmployeeVH("helpercode" + lstvh);
	showFCBAVH("lokasivh" + lstvh);
	showActivityVH("activityivh" + lstvh);
	showkategoriTR("kattransport" + lstvh);
		
}


function deleterowVH(i) {
	Spry.Widget.Utils.destroyWidgets('employeecode' + i);
	Spry.Widget.Utils.destroyWidgets('helpercode' + i);
	Spry.Widget.Utils.destroyWidgets('lokasi' + i);
	Spry.Widget.Utils.destroyWidgets('timestart' + i);
	Spry.Widget.Utils.destroyWidgets('activityivh' + i);
	Spry.Widget.Utils.destroyWidgets('tcvh' + i);
	Spry.Widget.Utils.destroyWidgets('kattransport' + i);
	Spry.Widget.Utils.destroyWidgets('trip' + i);
	Spry.Widget.Utils.destroyWidgets('jamkerja' + i);
	Spry.Widget.Utils.destroyWidgets('kmhm' + i);
	Spry.Widget.Utils.destroyWidgets('outputvh' + i);
	Spry.Widget.Utils.destroyWidgets('remarkvh' + i);
	 $('#trv' + i).remove();
	 
}


//get employee
function showEmployeeVH(selid) {
	$('#employeecode').empty();
	var tglbkm = $('#tglbkm').val();
	var bkmno = $('#bkmno').val();
	var emp = 3;
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_employee_edit.asp?emp=" + emp + "&tglbkm=" + tglbkm + "&bkmno=" + bkmno,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get activity kendaraan
function showActivityVH(selid) {
	
	$('#activityivh').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get kategori transport
function showkategoriTR(selid) {
	
	$('#kattransport').empty();
	var jmlitem = 4;
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_show_item.asp?jmlitem=' + jmlitem,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get tcvh1
function showtcVH(selid) {
	//alert(selid);
	
	 $('#tcvh').empty();
	
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//GET TCVH 
function getTCVH(i)
{
	var activityivh = $('#activityivh' + i).val();
	var lokasi   = $('#lokasivh' + i).val();
	$('#tcvh' + i).empty();
	$('#tcvh' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?activity=' + activityivh + '&lokasi=' + lokasi,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});

}

</script>
