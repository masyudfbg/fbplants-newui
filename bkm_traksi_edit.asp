<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno,fcbahdr
bkmno		= trim(request("bkmno"))
businessunit= trim(session("fcba"))

'response.write bkmno

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if


' get data header
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,SECTIONNAME,GANGCODE,ROWSTATEID,SUPERVISI1,SUPERVISINAME1,Clasification,Clasificationname,ROWSTATE,POSTINGFLG,VEHICLECODE,VEHICLENAME
sqlheader = "select h.BKMNO,to_char(h.BKMDATE,'mm/dd/yyyy')BKMDATE,h.GANGCODE,h.SUPERVISI1,(SELECT s1.FCNAME FROM EMPLOYEE  s1 WHERE h.SUPERVISI1=s1.FCCODE and s1.fcba=h.fcba ) AS SUPERVISINAME1, " & _
" h.SUPERVISI2,(SELECT s2.FCNAME FROM EMPLOYEE s2 WHERE h.SUPERVISI2=s2.FCCODE and s2.fcba=h.fcba) AS SUPERVISINAME2,h.SUPERVISI3, " & _
" (SELECT s3.FCNAME FROM EMPLOYEE s3 WHERE h.SUPERVISI3=s3.FCCODE and s3.fcba=h.fcba) AS SUPERVISINAME3,h.SUPERVISI4, " & _
" (SELECT s4.FCNAME FROM EMPLOYEE s4 WHERE h.SUPERVISI4=s4.FCCODE and s4.fcba=h.fcba) AS SUPERVISINAME4," & _ 
" h.SECTIONCODE,h.DOCUMENT_CLASSIFICATION,c.fcname, h.ROWSTATE,h.POSTINGFLG,h.VEHICLECODE,v.FCCODE || '-' || v.FCNAME VEHICLENAME,h.fcba " & _
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
	POSTINGFLG = rsHeader("POSTINGFLG")
	ROWSTATE = rsHeader("ROWSTATE")
	
	VEHICLECODE = rsHeader("VEHICLECODE")
	VEHICLENAME = rsHeader("VEHICLENAME")
	fcbahdr = rsHeader("fcba")
	
	
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


%>

<html>
	<head>
		<title>FBPlants - BKM TRAKSI</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>

		<link href="/js/spry/widgets/selectvalidation/SpryValidationSelect.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/selectvalidation/SpryValidationSelect.js" type="text/javascript"></script>
	<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261015a"></script>
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM TRAKSI EDIT</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_traksi_prosesedit.asp">
	<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
	<input type="hidden" name="jumldatavh" id="jumldatavh" value="<%=NoOfRowsitem%>" />
	<input type="hidden" name="fcbahdr" id="fcbahdr" value="<%=fcbahdr%>" />
	<input type="hidden" name="proses" id="proses" value="2" />
	
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >BKM No</td>
			<td align="left">
				<input type="text" name="bkmno" readonly id="bkmno" size="30" value="<%=bkmno%>" title="BKM No">
				
			</td>


			<td align="left" >Section Name</td>
			<td align="left">
				<input type="text" name="sectionname" readonly id="sectionname" size="20" value="<%=SECTIONNAME%>">
				
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
				<input type="text" class="" readonly maxlength="15" name="tglbkm" readonly id="tglbkm" size="15" value="<%=DATECREATED%>" title="Tanggal BKM" onchange="get_Allowentry();">
			</td>
			
			<td align="left" >Gang</td>
			<td align="left">
				<input type="text" name="gangcode"  readonly id="gangcode" size="15" value="<%=GANGCODE%>">
			</td>
			<td align="left" >Kendaraan</td>
			<td align="left">
				<select name="kendaraan" id="kendaraan" class="lebarsection" >
					<option value=<%=VEHICLECODE%>><%=VEHICLENAME%></option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" >Posting</td>
			<td align="left"><%=POSTINGFLG%>-<%=ROWSTATE%> <input type="hidden" value="<%=ROWSTATE%>" id="statusbkm"/></td>
			

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
								sqldtlupah = "SELECT u.BKMNO,u.chargeunit fcba, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
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
										<select class="lebaractivity " name="fccode<%=no%>" id="fccode<%=no%>" onchange="getFinger(<%=no%>)" onClick="getFinger(<%=no%>)">
											<option value=<%=rsdtlupah("EMPLOYEECODE")%>><%=rsdtlupah("fcname")%></option>
										</select>
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("IN_FINDEPAN")%>
										<span id="timein<%=no%>"></span>
										<input type="hidden" class="lebarabsen" name="timeinx<%=no%>" id="timeinx<%=no%>" value="<%=rsdtlupah("IN_FIN")%>">	
									</td>
									<td class="ratatengah">
										<%=rsdtlupah("OUT_FINDEPAN")%>
										<span id="timeout<%=no%>"></span>
										<input type="hidden" class="lebarabsen" name="timeoutx<%=no%>" id="timeoutx<%=no%>" value="<%=rsdtlupah("OUT_FIN")%>">	
									</td>
									<td class="ratatengah">
										<select class="lebarabsen " name="lokasi<%=no%>" id="lokasi<%=no%>" >
											<option value=<%=rsdtlupah("fcba")%>><%=rsdtlupah("fcba")%></option>
										</select>
									</td>
									<td class="ratatengah">
										<select class="lebarabsen " name="absen<%=no%>" id="absen<%=no%>" onchange="getDisable(<%=no%>)">
											<option value=<%=rsdtlupah("ATTENDANCE")%>><%=rsdtlupah("ATTENDANCE")%></option>
										</select>
									</td>
									<td class="ratakiri">
										<select class="lebaractivity ratakiri" name="activity<%=no%>" id="activity<%=no%>" onchange="getTC(<%=no%>);">
											<option value=<%=rsdtlupah("ACTCODE")%>><%=rsdtlupah("activityname")%></option>
										</select>
									</td>
									<td class="ratatengah">
										<select class="lebartc ratakiri" name="tcupah<%=no%>" id="tcupah<%=no%>">
											<option value=<%=rsdtlupah("LOCATIONCODE")%>><%=rsdtlupah("LOCATIONCODE")%></option>
										</select>
									</td>
									<td>
										<input class="ratatengah lebarhk" type="text" name="hk<%=no%>" id="hk<%=no%>" value="<%=rsdtlupah("MANDAYS")%>"/>
									</td>
									
									<td>
										<input class="ratatengah lebarabsen" type="text" name="unit<%=no%>" id="unit<%=no%>" value="<%=rsdtlupah("UNIT")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratatengah lebarabsen" type="text" name="output<%=no%>" id="output<%=no%>" value="<%=rsdtlupah("OUTPUT")%>"/>
									</td>
									
									<td class="ratakiri">
										<input class="ratatengah lebarrate" type="text" name="ratekb<%=no%>" id="ratekb<%=no%>" value="<%=rsdtlupah("rate")%>"/>
									</td>
									<td class="ratakiri">
										<select class="lebarallow ratakiri" name="allowedcode<%=no%>" id="allowedcode<%=no%>">
											<option value=<%=rsdtlupah("ALLOWANCE_CODE")%>><%=rsdtlupah("ALLOWDEDNAME")%></option>
										</select>
									</td>
																		
									<td class="ratakiri">
										<input class="ratakanan lebarrate" type="text" onkeyup="getIsibelakang(<%=no%>)"  name="ratepremidepan<%=no%>" id="ratepremidepan<%=no%>" value="<%=FormatNumber(rsdtlupah("ALLOWANCE_AMT"),0)%>"/>
										<input class="ratakanan lebarrate" type="hidden"   name="ratepremi<%=no%>" id="ratepremi<%=no%>" value="<%=rsdtlupah("ALLOWANCE_AMT")%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarhk" type="text" onkeyup="getOTHConvertion(<%=no%>);"  name="ovtimehours<%=no%>" id="ovtimehours<%=no%>" value="<%=rsdtlupah("OVERTIME_HOURS")%>"/>
									</td>
									<td class="ratakiri">
										<select class="lebarovtime " name="typeovrtime<%=no%>" id="typeovrtime<%=no%>" onchange="getOTHConvertion(<%=no%>);">
											<option value=<%=rsdtlupah("TYPE_OVERTIME")%>><%=rsdtlupah("type_ovname")%></option>
										</select>
										<input class="ratatengah lebarhk" type="hidden"   name="othconvertion<%=no%>" id="othconvertion<%=no%>" value="<%=rsdtlupah("othconvertion")%>"/>
									</td>
									
									<td>
										<img src='css/login/images/add.jpg' onclick="addrowUpah('<%=no%>')" />
									</td>
									<td>
										<% if no >= 0 then %>
											<img src='css/login/images/del.png' onclick="deleterowUpah('<%=no%>')" />
										<% else 
											response.write "&nbsp;"
										end if %>
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
										" bv.fcba, bv.VEHICLECODE,v.FCCODE || ' - ' || v.FCNAME VEHICLENAME, bv.ACTCODE,c.FCNAME || ' - ' || c.FCCODE COANAME, " & _ 
										" bv.LOCATIONCODE,bv.CATEGORYTRANSPORT,t.FCCODE || ' - ' || t.FCNAME TRANSPORTKATNAME, " & _
										" bv.TIME, bv.HOURS, bv.UNIT, bv.OUTPUT, bv.TRIP,bv.remarks,bv.chargeunit " & _
										" From   BKMVEHICLEDTL bv " & _
										" LEFT JOIN EMPLOYEE e on (e.FCCODE=bv.EMPLOYEECODE and e.fcba=bv.fcba)" & _
										" LEFT JOIN JOBMASTER c on (bv.ACTCODE=c.FCCODE) " & _
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
												<select class="lebarvehicle " name="employeecode<%=novh%>" id="employeecode<%=novh%>">
													<option value=<%=rsvehicle("EMPLOYEECODE")%>><%=rsvehicle("EMPLOYEENAME")%></option>
												</select>
											</td>
											<td class="ratakiri">
												<select class="lebarvehicle " name="helpercode<%=novh%>" id="helpercode<%=novh%>">
													<option value=<%=rsvehicle("HELPERCODE")%>><%=rsvehicle("HELPERNAME")%></option>
												</select>
											</td>
											
											<td class="ratakiri">
												<select class="lebarabsen " name="lokasivh<%=novh%>" onchange="getTCVH(<%=novh%>)" id="lokasivh<%=novh%>">
													<option value=<%=rsvehicle("chargeunit")%>><%=rsvehicle("chargeunit")%></option>
												</select>
											</td>
											<td class="ratatengah">
												<input class="ratatengah lebartime timepicker" autocomplete='off' type="text" name="timestart<%=novh%>" id="timestart<%=novh%>" value="<%=rsvehicle("TIME")%>"/>
											</td>

											<td class="ratakiri">
												<select class="lebaractivity ratakiri" name="activityivh<%=novh%>" onchange="getTCVH(<%=novh%>)" id="activityivh<%=novh%>">
													<option value=<%=rsvehicle("ACTCODE")%>><%=rsvehicle("COANAME")%></option>
												</select>
											</td>

											<td class="ratatengah">
												<select class="lebartc ratakiri" name="tcvh<%=novh%>" id="tcvh<%=novh%>">
													<option value=<%=rsvehicle("LOCATIONCODE")%>><%=rsvehicle("LOCATIONCODE")%></option>
												</select>
											</td>

											<td class="ratakiri">
												<select class="lebaractivity ratakiri" name="kattransport<%=novh%>" id="kattransport<%=novh%>">
													<option value=<%=rsvehicle("CATEGORYTRANSPORT")%>><%=rsvehicle("TRANSPORTKATNAME")%></option>
												</select>
											</td>

											<td class="ratatengah">
												<input class="ratatengah lebarabsen" type="text" name="trip<%=novh%>" id="trip<%=novh%>" value="<%=rsvehicle("trip")%>"/>
											</td>

											<td class="ratatengah">
												<input class="ratatengah lebarabsen" type="text" name="jamkerja<%=novh%>" id="jamkerja<%=novh%>" value="<%=rsvehicle("HOURS")%>"/>
											</td>


											<td class="ratatengah">
												<input class="ratatengah lebarabsen" type="text" name="kmhm<%=novh%>" id="kmhm<%=novh%>" value="<%=rsvehicle("UNIT")%>"/>
											</td>

											<td class="ratatengah">
												<input class="ratatengah lebarabsen" type="text" name="outputvh<%=novh%>" id="outputvh<%=novh%>" value="<%=rsvehicle("OUTPUT")%>"/>
											</td>

											<td class="ratakiri">
												<textarea class="ratakiri lebaractivity" row="3" cols=40 name="remarkvh<%=novh%>" id="remarkvh<%=novh%>"><%=rsvehicle("remarks")%></textarea>
											</td>

											<td>
												<img src='css/login/images/add.jpg' onclick="addrowVH('<%=novh%>')" />
											</td>
											<td>
												<% if novh >= 0 then %>
													<img src='css/login/images/del.png' onclick="deleterowVH('<%=novh%>')" />
												<% else 
													response.write "&nbsp;"
												end if %>
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
			<%
			if  POSTINGFLG = "TIDAK" then%>
			<td  colspan="11"><input type=Submit name="btnsave" id="btnsimpan" value="SAVE" onclick="prosesSubmit()">&nbsp;&nbsp;<input type=button onclick="window.close();" name="btncancel" value="CLOSE">&nbsp;&nbsp;<input type=submit name="btnsave" id="btndelete" value="DELETE"></td>
			<%End if%>
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
			sqldtlupah = "SELECT u.BKMNO,u.chargeunit fcba, u.EMPLOYEECODE,e.fcname || '-' ||  e.fccode as fcname, u.ACTCODE,a.fcname || '-' ||  u.ACTCODE as activityname, NVL(u.MANDAYS,0)MANDAYS, " & _
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
			<td class="ratakanan"><%=formatnumber(rateallow,2)%></td>
			<td class="ratatengah"><%=rsupah("OVERTIME_HOURS")%></td>
			<td class="ratakiri"><%=rsupah("remarks")%></td>
			<td class="ratakanan"><%=formatnumber(rsupah("AMOUNT"),2)%></td>
		</tr>
		<%
			totalamount = cdbl(totalamount) + cdbl(rsupah("AMOUNT"))
			rsupah.MoveNext
			Wend
		%>
		<tr>
			<td colspan="13" class="ratakanan tebal">Total</td>
			<td class="ratakanan tebal"><%=formatnumber(totalamount,2)%></td>
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

	get_Allowentry();
	
});

function prosesSubmit() {
	
	if($('#fcbadata').val() == "ARM"){
		document.frmEntry.submit();
	}else{
		var errors = [];

		// === VALIDASI HEADER ===
		if (!$.trim($('#tglbkm').val()))    errors.push("• Tanggal BKM harus dipilih.");
		if (!$('#sectionname').val())        errors.push("• Section Name harus dipilih.");
		if (!$('#kendaraan').val())          errors.push("• Kendaraan harus dipilih.");

		if (errors.length > 0) {
			alert("Inputan anda belum lengkap:\n\n" + errors.join("\n"));
			return;
		}

		// === VALIDASI TAB BKM UPAH ===
		var jumldata = parseInt($('#jumldata').val()) || 0;
		if (jumldata === 0) {
			alert("Inputan anda belum lengkap:\n\n• Tab BKM UPAH belum diisi.");
			return;
		}

		var upahErrors = [];
		var skipValidasi = ["P1", "MK", "CT", "ML", "KH", "UN", "S1", "KB"];
		// var skipValidasi = ["KJ", "WS", "WH", "DL", "KH"];

		for (var i = 1; i <= jumldata; i++) {
			if ($('#tru' + i).length === 0) continue;

			if (!$('#fccode'   + i).val()) upahErrors.push("Baris " + i + " (Upah): Employee harus diisi.");
			if (!$('#lokasi'   + i).val()) upahErrors.push("Baris " + i + " (Upah): Lokasi harus diisi.");
			if (!$('#absen'    + i).val()) upahErrors.push("Baris " + i + " (Upah): Absen harus diisi.");
			if (!$('#activity' + i).val()) upahErrors.push("Baris " + i + " (Upah): Activity harus diisi.");
			if (!$('#tcupah'   + i).val()) upahErrors.push("Baris " + i + " (Upah): TC harus diisi.");

			var absen        = $('#absen'        + i).val();
			var presentgrade = $('#presentgrade' + i).val();

			if (!skipValidasi.includes(absen)) {
				var hk = parseFloat($.trim($('#hk' + i).val())
					.replace(/\./g,"")
					.replace(",","."));

				if (isNaN(hk) || hk <= 0)
					upahErrors.push("Baris " + i + " (Upah): HK harus lebih dari 0.");
			}

			// Validasi tambahan: jika Status KHT, HK wajib diisi > 0
			if (presentgrade === "KHT") {
				var hkKHT = parseFloat($.trim($('#hk' + i).val())
					.replace(/\./g,"")
					.replace(",","."));

				if (isNaN(hkKHT) || hkKHT <= 0)
					upahErrors.push("Baris " + i + " (Upah): Status KHT, HK harus lebih dari 0.");
			}

			var ovtime = parseFloat($.trim($('#ovtimehours' + i).val()).replace(/\./g,"").replace(",","."));
			if (!isNaN(ovtime) && ovtime > 0 && !$('#typeovrtime' + i).val()) {
				upahErrors.push("Baris " + i + " (Upah): Tipe Overtime harus dipilih.");
			}
		}

		if (upahErrors.length > 0) {
			alert("Inputan anda belum lengkap:\n\n" + upahErrors.join("\n"));
			return;
		}

		// === CEK KONDISI SEMUA BARIS UPAH: skipValidasi + KHT ===
		// Jika semua baris memenuhi kondisi ini, BKM VEHICLE boleh kosong
		var semuaKHT = true;
		for (var k = 1; k <= jumldata; k++) {
			if ($('#tru' + k).length === 0) continue;

			var absenK        = $('#absen'        + k).val();
			var presentgradeK = $('#presentgrade' + k).val();

			if (!(skipValidasi.includes(absenK) && presentgradeK === "KHT")) {
				semuaKHT = false;
				break;
			}
		}

		// === VALIDASI TAB BKM VEHICLE ===
		if (!semuaKHT) {
			var jumldatavh = parseInt($('#jumldatavh').val()) || 0;
			if (jumldatavh === 0) {
				alert("Inputan anda belum lengkap:\n\n• Tab BKM VEHICLE/WORKSHOP belum diisi.");
				return;
			}

			var vhErrors = [];
			for (var j = 0; j <= jumldatavh; j++) {
				if ($('#trv' + j).length === 0) continue;
				
				// Tambahkan cek: skip baris yang semua field kosong (baris baru belum diisi)
				var empVal = $('#employeecode' + j).val();
				var lokVal = $('#lokasivh' + j).val();
				var actVal = $('#activityivh' + j).val();
				
				// Jika semua field utama kosong, anggap baris ini belum dipakai, skip
				if (!empVal && !lokVal && !actVal) continue;
				
				if (!empVal) vhErrors.push("Baris " + j + " (Vehicle): Operator/Driver harus diisi.");
				if (!lokVal) vhErrors.push("Baris " + j + " (Vehicle): Lokasi harus diisi.");
				if (!$('#activityivh' + j).val()) vhErrors.push("Baris " + j + " (Vehicle): Activity harus diisi.");
				if (!$('#tcvh' + j).val()) vhErrors.push("Baris " + j + " (Vehicle): TC harus diisi.");

				var timestart = $.trim($('#timestart' + j).val());
				if (!timestart) {
					vhErrors.push("Baris " + j + " (Vehicle): Time harus diisi.");
				} else if (!/^\d{2}:\d{2}-\d{2}:\d{2}$/.test(timestart)) {
					vhErrors.push("Baris " + j + " (Vehicle): Format Time harus hh:mm-hh:mm.");
				}

				var jamkerja = parseFloat($.trim($('#jamkerja' + j).val()).replace(/\./g,"").replace(",","."));
				if (isNaN(jamkerja) || jamkerja <= 0) vhErrors.push("Baris " + j + " (Vehicle): Hours harus lebih dari 0.");
			}

			if (vhErrors.length > 0) {
				alert("Inputan anda belum lengkap:\n\n" + vhErrors.join("\n"));
				return;
			}
		}

		// === SEMUA LOLOS, BARU SUBMIT ===
		document.frmEntry.submit();
    }
}

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
			url : 'bkm_getactivity.asp?panen=' + 2,
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

function get_Approved()
{
	var gangcode = $("#rowstate_").val();
	if (gangcode != "Approved") {
				$('#btnsimpan').show();
				$('#btndelete').show();
			}else{
				alert("BKM Sudah Approved !!!");
				$('#btnsimpan').hide();
				$('#btndelete').hide();
			} 
}

function get_Allowentry()
{
	var gangcode = $("#gangcode").val();
	var tglbkm = $('#tglbkm').val();
	var statusbkm = $("#statusbkm").val();
	
	if (tglbkm != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			console.log(closingStatus +' - '+ statusbkm);
			if(closingStatus=="FALSE" && statusbkm == "Planned"){
				addrowVH(1);
				$('#btnsimpan').show();
				$('#btndelete').show();
			} else if (closingStatus == "TRUE" && statusbkm == "Review Accounting" || statusbkm == "Planned") {
				addrowVH(1);
				$('#btnsimpan').show();
				$('#btndelete').show();
			} else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$('#btnsimpan').hide();
				$('#btndelete').hide();
				
				//break;
			}
			
			
						
		});
		
	} 
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
			getTC(i);
			showAllowed("allowedcode" + i);
			showTypeOv("typeovrtime" + i);
			
		}
	});
	
}


var jumlinv = $('#jumldata').val();
var StartIteration =  parseInt(jumlinv) + 1	
function addrowUpah(trid) {
	var fcba = '<%= Session("fcba") %>';
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

	var lokasi = "<select name='lokasi" + lastrow + "' id='lokasi" + lastrow + "' class='ratakiri lebarabsen'><option value='" + fcba + "'>" + fcba + "</option></select>";
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

	new Spry.Widget.ValidationSelect("fccode" + lastrow);
	new Spry.Widget.ValidationSelect("lokasi" + lastrow);
	new Spry.Widget.ValidationSelect("absen" + lastrow);
	new Spry.Widget.ValidationSelect("activity" + lastrow);
	new Spry.Widget.ValidationSelect("tcupah" + lastrow);
	new Spry.Widget.ValidationTextField("hk" + lastrow);
	
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
	var currentVal  = $('#tcupah' + i).val();
	var currentText = $('#tcupah' + i + ' option:selected').text();

	$('#tcupah' + i).empty();
	if (currentVal) {
		$('#tcupah' + i).append(new Option(currentText, currentVal, true, true));
	}
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
			getTCVH(i);
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

	var lokasivh = "<select name='lokasivh" + lstvh + "' id='lokasivh" + lstvh + "' class='ratakiri lebarabsen' onchange='getTCVH(" + lstvh + ")'></select>";
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
	Spry.Widget.Utils.destroyWidgets('lokasivh' + i);
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
	var veh = $("#kendaraan").val();
	$('#activityivh').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 4 + '&veh=' + veh,
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
	var tglbkm = $('#tglbkm').val();
	var currentVal  = $('#tcvh' + i).val();
	var currentText = $('#tcvh' + i + ' option:selected').text();

	$('#tcvh' + i).empty();
	if (currentVal) {
		$('#tcvh' + i).append(new Option(currentText, currentVal, true, true));
	}
	$('#tcvh' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?activity=' + activityivh + '&lokasi=' + lokasi + '&tglbkm=' + tglbkm,
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
