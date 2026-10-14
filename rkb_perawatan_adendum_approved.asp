<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - Monthly Work Plan</title><link rel="StyleSheet" href="dtree.css" type="text/css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

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
</head>
<style>
	

	.lebarsection{
		width: 160px;
	}
	.lebarnormahk{
		width: 100px;
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

	.lebarlokasi{
		width: 110px;
	}
	.lebarstatus{
		width: 70px;
	}

	.lebarhaactual{
		width: 70px;
	}

	.lebartc{
		width: 50px;
	}

	.lebaractbarang{
		width: 180px;
	}
</style>
<%
	dim rkbpost,norkb
	dim sqlhdr,rshdr 
	dim RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE,ROWSTATE, DOCUMENT_CLASSIFICATION, TOTUPAH,CREATEDATE

	dim classification,simpan,fcba,countrkb,sqlsimpandtl
	DIM statuspost,periodethn,periodebln,section,tglinput,ttlamountupah
	dim sqlnomor,rsnomor,tglprev,currentnomor,currentnomorfix,chargeunit
	DIM actcode,blok,ttlhk,rate,normamandays,luasdikerjakan,rotasi,totalamounthk
	DIM nQtdFields,nCount,linenokey



	rkbpost = request("rkbno")
	simpan = request("simpan")

	sqlhdr = "Select RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE, TO_CHAR(CREATEDATE,'DD/MM/YYYY')CREATEDATE, " & _
	" LASTUPDATE, FCENTRY, FCEDIT, FCIP, FCBA, " & _
	" POSTINGFLG, ROWSTATE, DOCUMENT_CLASSIFICATION, TOTUPAH, TOTBAHAN, " & _
	" TOTVRA, TOTWIP " & _
	" From RKBHDR WHERE RKBNO='" & rkbpost & "'"
	'Response.Write sqlhdr
	'connStr.execute(sqlhdr)
	set rshdr = connStr.Execute(sqlhdr)
	if not rshdr.eof then
		
		PERIOD_YEAR	= rshdr("PERIOD_YEAR") 
		PERIOD_MONTH= rshdr("PERIOD_MONTH")  
		SECTIONCODE	= rshdr("SECTIONCODE")
		ROWSTATE	= rshdr("ROWSTATE")
		DOCUMENT_CLASSIFICATION	= rshdr("DOCUMENT_CLASSIFICATION")
		TOTUPAH		= rshdr("TOTUPAH")
		CREATEDATE	= rshdr("CREATEDATE")
	end if

	

	if simpan="SAVEDATA" then

			RKBNO 			= request("norkbview")
			fcba 			= session("fcba")
			'Response.Write fcba

			'looping di tabel detail
			nQtdFields 	= request.form("jmlrawat")
			nCount 		= 1
			linenokey  	= 1

			'Response.Write nQtdFields
			'Response.End
			
			do while cint(nCount) <= cint(nQtdFields) 
				actcode 		= request.form("activityrwt" & nCount)
			'	Response.Write actcode
			'	Response.End
				blok 			= request.form("tcrwt" & nCount)
				ttlhk 			= request.form("ttlhk" & nCount)
				rate 			= request.form("rate" & nCount)
				normamandays	= request.form("mandayform" & nCount)
				luasdikerjakan	= request.form("luasdikerjakan" & nCount)
				rotasi			= request.form("rotasi" & nCount)
				totalamounthk	= request.form("amountbck" & nCount)
				haactual 		= request.form("haactual" & nCount)

					
				sqlsimpandtl = "insert into RKBUPAHDTL(RKBNO, ACTCODE, LOCATIONCODE, MANDAYS, RATE, " & _
				" NORMAMANDAYS, VOLUME,  " & _
				" ROTASI,TOTALAMOUNT_HK,LINENOKEY, " & _
				" CREATEDATE, LASTUPDATE, FCENTRY, FCEDIT, FCIP,FCBA,HAACTUAL,ROWSTATE) " & _
				" values('" & RKBNO & "','" & actcode & "','" & blok & "','" & ttlhk & "', " & _ 
				" '" & rate & "','" & normamandays & "','" & luasdikerjakan & "','" & rotasi & "','" & totalamounthk & "','" & linenokey & "', " & _ 
				" TO_DATE('" & tglinput & "', 'dd/mm/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "', " & _ 
				" '" & fcba & "','" & haactual & "','Planned')"
				'connstr.BeginTrans
				'Response.Write sqlsimpandtl
				connStr.execute(sqlsimpandtl)

			
				nCount 	= nCount + 1
				linenokey 	= linenokey + 1
			loop

		response.redirect("rkb_perawatan_adendum.asp?rkbno="&RKBNO)
	End IF

%>

<body>
<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
 <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 	<div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">  

	<input type="hidden" name="fcbatxt" readonly id="fcbatxt" value="<%=fcba%>">

	<table align=center border="0" class='data posisikiri' width='100%' id='tableall'>			
		<tr>
			<td  align="left">
				<font face="Verdana" size="1" align="LEFT">No. RKB</font>
			</td>
			<td  align="left">
				<input type="text" name="norkb" readonly id="norkb" size="30" value="<%=rkbpost%>" title="No RKB" style='background-color:#eee;'>
			</td>
			<td align="left" >Periode</td>
			<td align="left">
			    <select name="tahun" id="tahun" style="width:100px">
				<%
					dim sqltahun7,rstahun7
					sqltahun7 = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun desc"  
					Set rstahun7 = Server.CreateObject("ADODB.Recordset")
					rstahun7.CursorType = 3
					rstahun7.CursorLocation = 3
					rstahun7.Open sqltahun7, connstr
					While Not rstahun7.EOF 
						if cint(period_year)=cint(rstahun7("tahun")) then %>
						<option value="<%=rstahun7("tahun")%>" Selected ><%=rstahun7("tahun")%></option>
						<%else%>
						<option value="<%=rstahun7("tahun")%>"><%=rstahun7("tahun")%></option>
						<%end if%>					
					<%rstahun7.movenext
					wend 
					rstahun7.close 
					

				%>	
				</select>
				<select name="bulan" id="bulan" style="width:120px">
				<%
					dim sqlbulan7,rsbulan7
					sqlbulan7 = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
					Set rsbulan7 = Server.CreateObject("ADODB.Recordset")
					rsbulan7.CursorType = 3
					rsbulan7.CursorLocation = 3
					rsbulan7.Open sqlbulan7, connstr
					While Not rsbulan7.EOF 
						if cint(period_month)=cint(rsbulan7("bulan")) then %>
						<option value="<%=rsbulan7("bulan")%>" Selected><%=rsbulan7("description")%></option>
						<%else%>
						<option value="<%=rsbulan7("bulan")%>"><%=rsbulan7("description")%></option>
						<%end if%>
					<%rsbulan7.movenext
					wend 
					rsbulan7.close 
					

				%>	
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" >Classification</td>
			<td align="left">
				<input type="text" name="classifiction" readonly id="classifiction" size="30" value="<%=DOCUMENT_CLASSIFICATION%>" title="Classification">
				
			</td>
			<td align="left" >Section</td>
			<td align="left" colspan="3">
				<select name="sectionname" id="sectionname" class="lebarsection" readonly>
						<option value="<%=SECTIONCODE%>"><%=SECTIONCODE%></option>
				</select>
			</td>
					
		</tr>
		<tr>
			<td align="left">Status</td>
			<td align="left">
				<%=ROWSTATE%>
			</td>
			
		</tr>

	</table>
	
	<br>
	
	<table align=center border="0" class='data posisikiri' width='100%' id='tableall'>
		
		<tr>
			<td>
			   	<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">RKB Perawatan</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm wip-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table0'>
								<tr id="trheader0">
									<th>Aktivitas</th>
									<th>Blok</th>
									<th>Ha Actual</th>
									<th>Rate</th>
									<th>Mandays Norm</th>
									<th>Luas Dikerjakan</th>
									<th>Total HK</th>
									<th>Rotasi</th>
									<th>Amount</th>
									<th>Status</th>
									
								</tr>


								<%
									DIM ttlluas,ttlsemuahk,ttlrotasi,ttlsemuaamount
									DIM sqldtl,rsdtl 
									sqldtl = "Select R.RKBNO, R.ACTCODE,J.FCNAME,R.LOCATIONCODE, NVL(R.MANDAYS,0)MANDAYS, NVL(R.RATE,0)RATE, " & _ 
									" NVL(R.NORMAMANDAYS,0)NORMAMANDAYS, NVL(R.VOLUME,0)VOLUME, NVL(R.ROTASI,0)ROTASI,NVL(R.TOTALAMOUNT_HK,0)TOTALAMOUNT_HK, R.LINENOKEY,R.FCBA,R.HAACTUAL,R.ROWSTATE " & _ 
									" From   RKBUPAHDTL R " & _ 
									" LEFT JOIN JOBMASTER J ON R.ACTCODE=J.FCCODE " & _ 
									" WHERE R.RKBNO='" & rkbpost & "' AND R.ROWSTATE='Approved' AND R.ADDENDUM=0"
									set rsdtl = connStr.Execute(sqldtl)
									'Response.Write sqldtl
																		
									no = 0
									ttlluas = 0
									ttlsemuahk = 0
									ttlrotasi = 0
									ttlsemuaamount = 0
									While Not rsdtl.EOF %>
														

								<tr id="trb0<%=no%>">
									
									<td class="ratakiri">
										<select class="lebaractivity " name="activityrwtv<%=no%>" id="activityrwtv<%=no%>" class='lebaractbarang' onchange="getTCrwt(<%=no%>)">
											<option value=<%=rsdtl("ACTCODE")%>><%=rsdtl("fcname")%></option>
										</select>
									</td>

									<td class="ratatengah">
										<select class='lebartc ratatengah' name="tcrwtv<%=no%>" id="tcrwtv<%=no%>" >
											<option value=<%=rsdtl("LOCATIONCODE")%>><%=rsdtl("LOCATIONCODE")%></option>
										</select>
									</td>

									<td>
										<input class="ratatengah lebarhaactual" type="text" name="haactualv<%=no%>" id="haactualv<%=no%>" value="<%=formatnumber(rsdtl("HAACTUAL"))%>"/>
									</td>

									<td>
										<input class='lebarhaactual ratatengah' type="text" name="ratev<%=no%>" id="ratev<%=no%>" value="<%=rsdtl("RATE")%>"/>
									</td>
									
									<td>
										<input class="ratatengah lebarnormahk" type="text" name="mandayformv<%=no%>" id="mandayformv<%=no%>" value="<%=formatnumber(rsdtl("NORMAMANDAYS"),4)%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text"  name="luasdikerjakanv<%=no%>" id="luasdikerjakanv<%=no%>" value="<%=rsdtl("VOLUME")%>"/>
									</td>
									
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text" name="ttlhkv<%=no%>" id="ttlhkv<%=no%>" value="<%=formatnumber(rsdtl("MANDAYS"),4)%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebartc" type="text" name="rotasiv<%=no%>" id="rotasiv<%=no%>" value="<%=rsdtl("ROTASI")%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebarlokasi" type="text" name="amountv<%=no%>" id="amountv<%=no%>" value="<%=rsdtl("TOTALAMOUNT_HK")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebarstatus" type="text" name="rowstatev<%=no%>" id="rowstatev<%=no%>" value="<%=rsdtl("ROWSTATE")%>"/>
									</td>
																
									
									<td></td>
									
								</tr>
								<%
									ttlluas = cdbl(ttlluas) + cdbl(rsdtl("VOLUME"))
									ttlsemuahk = cdbl(ttlsemuahk) + cdbl(rsdtl("MANDAYS"))
									ttlrotasi = cdbl(ttlrotasi) + cdbl(rsdtl("ROTASI"))
									ttlsemuaamount = cdbl(ttlsemuaamount) + cdbl(rsdtl("TOTALAMOUNT_HK"))
									no = no +1 
									rsdtl.MoveNext
									Wend
								
								%>
								<tr class="ratatengah">
										<td colspan="5" class="tebal">Total</td>
										<td class="tebal">
											<input class="ratatengah lebarhaactual" type="text" readonly  style="background-color:#eee;" name="amountluasdikerjakanfrm" id="amountluasdikerjakanfrm" value="<%=formatnumber(ttlluas,0)%>"/>
											<input class="ratatengah lebarhaactual" type="hidden" readonly  style="background-color:#eee;" name="amountluasdikerjakan" id="amountluasdikerjakan" value="<%=ttlluas%>"/>
										
										</td>
										<td class="tebal">
											<input class="ratakanan lebarhaactual" type="text" readonly  style="background-color:#eee;" name="amountttlhkfrm" id="amountttlhkfrm" value="<%=formatnumber(ttlsemuahk,2)%>"/>
											<input class="ratakanan lebarhaactual" type="hidden" readonly  style="background-color:#eee;" name="amountttlhk" id="amountttlhk" value="<%=ttlsemuahk%>"/>
										</td>
										<td class="tebal">
											<input class="ratakanan lebarhaactual" type="text" readonly  style="background-color:#eee;" name="amountrotasifrm" id="amountrotasifrm" value="<%=formatnumber(ttlrotasi,0)%>"/>
											<input class="ratakanan lebarhaactual" type="hidden" readonly  style="background-color:#eee;" name="amountrotasi" id="amountrotasi" value="<%=ttlrotasi%>"/>
										</td>
										<td class="tebal">
											<input class="ratakanan lebarlokasi" type="text" readonly  style="background-color:#eee;" name="amountrkbrpfrm" id="amountrkbrpfrm" value="<%=formatnumber(ttlsemuaamount,0)%>"/>
											<input class="ratakanan lebarlokasi" type="hidden" readonly  style="background-color:#eee;" name="amountrkbrp" id="amountrkbrp" value="<%=ttlsemuaamount%>"/>
										</td>
								</tr>
								
							</table> <br>
							
						</div>

					</div>
				</div>
			</td>		
		</tr>
			
	</table><br>
<form name="frmEntry" method="post" action="rkb_perawatan_adendum.asp">
	<input type="hidden" name="jmlrawat" id="jmlrawat" value="0" />
	<input type="hidden" name="norkbview" readonly id="norkbview" size="30" value="<%=rkbpost%>" title="No RKB" style='background-color:#eee;'>
	<table align=center border="0" class='data posisikiri' width='100%' id='table1'>
		<tr>
			<td>
					<div class="TabbedPanels" id="tp2" align='center'>
					
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">Addendum RKB Perawatan</li>
						
					</ul>
					<div class="TabbedPanelsContentGroup">

						<!-- view bkm wip-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table1'>
								<tr id="trheader1">
									<th>Aktivitas</th>
									<th>Blok</th>
									<th>Ha Actual</th>
									<th>Rate</th>
									<th>Mandays Norm</th>
									<th>Luas Dikerjakan</th>
									<th>Total HK</th>
									<th>Rotasi</th>
									<th>Amount</th>
									<th>Status</th>
								</tr>

								<%
									'DIM ttlluas,ttlsemuahk,ttlrotasi,ttlsemuaamount
									DIM sqldtl1,rsdtl1 
									sqldtl1 = "Select R.RKBNO, R.ACTCODE,J.FCNAME,R.LOCATIONCODE, NVL(R.MANDAYS,0)MANDAYS, NVL(R.RATE,0)RATE, " & _ 
									" NVL(R.NORMAMANDAYS,0)NORMAMANDAYS, NVL(R.VOLUME,0)VOLUME, NVL(R.ROTASI,0)ROTASI,NVL(R.TOTALAMOUNT_HK,0)TOTALAMOUNT_HK, R.LINENOKEY,R.FCBA,R.HAACTUAL,R.ROWSTATE " & _ 
									" From   RKBUPAHDTL R " & _ 
									" LEFT JOIN JOBMASTER J ON R.ACTCODE=J.FCCODE " & _ 
									" WHERE R.RKBNO='" & rkbpost & "' AND R.ADDENDUM=1"
									set rsdtl1 = connStr.Execute(sqldtl1)
									'Response.Write sqldtl
																		
									noadd = 0
									ttlluas = 0
									ttlsemuahk = 0
									ttlrotasi = 0
									ttlsemuaamount = 0
									While Not rsdtl1.EOF %>
														

								<tr id="trb0<%=noadd%>">
									
									<td class="ratakiri">
										<select class="lebaractivity " name="activityrwtv<%=noadd%>" id="activityrwtv<%=noadd%>" class='lebaractbarang' onchange="getTCrwt(<%=noadd%>)">
											<option value=<%=rsdtl1("ACTCODE")%>><%=rsdtl1("fcname")%></option>
										</select>
									</td>

									<td class="ratatengah">
										<select class='lebartc ratatengah' name="tcrwtv<%=noadd%>" id="tcrwtv<%=noadd%>" >
											<option value=<%=rsdtl1("LOCATIONCODE")%>><%=rsdtl1("LOCATIONCODE")%></option>
										</select>
									</td>

									<td>
										<input class="ratatengah lebarhaactual" type="text" name="haactualv<%=noadd%>" id="haactualv<%=noadd%>" value="<%=formatnumber(rsdtl1("HAACTUAL"))%>"/>
									</td>

									<td>
										<input class='lebarhaactual ratatengah' type="text" name="ratev<%=noadd%>" id="ratev<%=noadd%>" value="<%=rsdtl1("RATE")%>"/>
									</td>
									
									<td>
										<input class="ratatengah lebarnormahk" type="text" name="mandayformv<%=noadd%>" id="mandayformv<%=noadd%>" value="<%=formatnumber(rsdtl1("NORMAMANDAYS"),4)%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text"  name="luasdikerjakanv<%=noadd%>" id="luasdikerjakanv<%=noadd%>" value="<%=rsdtl1("VOLUME")%>"/>
									</td>
									
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text" name="ttlhkv<%=noadd%>" id="ttlhkv<%=noadd%>" value="<%=formatnumber(rsdtl1("MANDAYS"),4)%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebartc" type="text" name="rotasiv<%=noadd%>" id="rotasiv<%=noadd%>" value="<%=rsdtl1("ROTASI")%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebarlokasi" type="text" name="amountv<%=noadd%>" id="amountv<%=noadd%>" value="<%=rsdtl1("TOTALAMOUNT_HK")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebarstatus" type="text" name="rowstatev<%=noadd%>" id="rowstatev<%=noadd%>" value="<%=rsdtl1("ROWSTATE")%>"/>
									</td>
																
									
									<td></td>
									
								</tr>
								<%
									ttlluas = cdbl(ttlluas) + cdbl(rsdtl1("VOLUME"))
									ttlsemuahk = cdbl(ttlsemuahk) + cdbl(rsdtl1("MANDAYS"))
									ttlrotasi = cdbl(ttlrotasi) + cdbl(rsdtl1("ROTASI"))
									ttlsemuaamount = cdbl(ttlsemuaamount) + cdbl(rsdtl1("TOTALAMOUNT_HK"))
									noadd = noadd +1 
									rsdtl1.MoveNext
									Wend
								
								%>
								
							
							</table>
							
						</div>

					</div>
				</div>
			</td>
		</tr>
		
	</table>
	<br><br>
	<table align=center border="1" width="900" class='data' id='table2'>
								<tr>
									<th colspan="3">Summary</th>
								</tr>
								<tr id="trheader2">
									<th>Aktivitas</th>
									<th>Blok</th>
									<th>Luas Dikerjakan</th>
									
								</tr>
								
								<%
									dim sqlsum,rssum
									sqlsum = " Select R.RKBNO, R.ACTCODE || ' - ' || J.FCNAME JOBNAME,R.LOCATIONCODE, SUM(NVL(R.VOLUME,0))VOLUME " & _ 
									" From   RKBUPAHDTL R  " & _ 
									" LEFT JOIN JOBMASTER J ON R.ACTCODE=J.FCCODE  " & _ 
									" WHERE R.RKBNO='" & rkbpost & "' " & _ 
									" GROUP BY R.RKBNO, R.ACTCODE,J.FCNAME,R.LOCATIONCODE"
									set rssum = connStr.Execute(sqlsum)
									'Response.Write sqldtl
																		
									ttlluassummary = 0
									While Not rssum.EOF %>
										<tr style="font-size: 15px;">
											<td align="left"><%=rssum("JOBNAME")%></td>
											<td align="center" width="60px"><%=rssum("LOCATIONCODE")%></td>
											<td align="right" width="90px"><%=formatnumber(rssum("VOLUME"),0)%></td>
									
										</tr>
									<%
										ttlluassummary = cdbl(ttlluassummary) + cdbl(rssum("VOLUME"))
										rssum.MoveNext
										Wend
									%>
									<tr>
										<td colspan="2" class="tebal" style="font-size: 15px;">Total</td>
										<td class="tebal" style="font-size: 15px;"><%=formatnumber(ttlluassummary,0)%></td>
									</tr>
							</table>

</form>


 </div>
 </div>
 </body>
 </head>
 <link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script type="text/javascript">
function hanyaAngka(evt) {
	evt.value = evt.value.replace(/[^0-9,.-]/g,"");
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

$(function() {	
	$("#tglinput").datepicker();
	showSection();
	
});

function getNorkb()
{
	var sectionname = $("#sectionname").val();
	var tahun 		= $("#tahun").val();
	var bulan 		= $("#bulan").val();
	var fcbatxt 	= $("#fcbatxt").val();
	var norkb		= "RKB/" + fcbatxt + "/" + sectionname + "/" + tahun + bulan
	$("#norkb").val(norkb);
}


 //get section
function showSection() {
	
	//$('#sectionname').empty();
	$('#sectionname').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsection.asp?section=' + 1,
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

//show fcba
function getFCBA()
{
	var company = $("#companyar").val();
	showFCBA(company);
}

function showFCBA(company) {
	$('#fcba').empty();
	$('#fcba').select2({
		ajax: {
			delay: 1000,
			url : 'getListfcbaallajax.asp?fccompanycode=' + company,
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


var StartIterationitem=1;
function addrowPerawatan(trid)
{
	var lstitem = this.StartIterationitem++;
	if (lstitem < 2) {
		var row = $("#trheader1").after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb' + trid).after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowPerawatan(" + lstitem + ");' />";
	}

	var activityrwt = "<select name='activityrwt" + lstitem + "' id='activityrwt" + lstitem + "' class='lebaractbarang' onchange='getTCrwt(" + lstitem + ");getRate(" + lstitem + ");getNormamandays(" + lstitem + ")'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + activityrwt + "</td>");

	var tcrwt = "<select name='tcrwt" + lstitem + "' id='tcrwt" + lstitem + "' class='lebartc ratatengah' onchange='getHaactual(" + lstitem + ")'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + tcrwt + "</td>");

	var haactual = "<input type='text' name='haactual" + lstitem + "' id='haactual" + lstitem + "' style='background-color:#eee;' value='' readonly class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + haactual + "</td>");

	var rate = "<input type='text' name='rate" + lstitem + "' id='rate" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' style='background-color:#eee;' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + rate + "</td>");
	
	var mandayform = "<select name='mandayform" + lstitem + "' id='mandayform" + lstitem + "' class='lebarnormahk ratatengah'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + mandayform + "</td>");
	
	var luasdikerjakan = "<input type='text' name='luasdikerjakan" + lstitem + "' id='luasdikerjakan" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getTotalhk(" + lstitem + ");' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + luasdikerjakan + "</td>");

	var ttlhk = "<input type='text' name='ttlhk" + lstitem + "' id='ttlhk" + lstitem + "' value='0' onkeyup='hanyaAngka(this);' style='background-color:#eee;'  readonly class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + ttlhk + "</td>");
	
	var rotasi = "<input type='text' name='rotasi" + lstitem + "' id='rotasi" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmount(" + lstitem + ");' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + rotasi + "</td>");
	
	var amount = "<input type='text' name='amount" + lstitem + "' id='amount" + lstitem + "' style='background-color:#eee;' value='0' onkeyup='hanyaAngka(this)' class='lebarlokasi ratakanan'><input type='hidden' name='amountbck" + lstitem + "' id='amountbck" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + amount + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowPerawatan(" + lstitem + ");' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jmlrawat').val(lstitem);

	showActivityrwt("activityrwt" + lstitem);

	
}

//hapus row rkbperawatan
function deleterowPerawatan(i) {
	Spry.Widget.Utils.destroyWidgets('activityrwt' + i);
	Spry.Widget.Utils.destroyWidgets('tcrwt' + i);
	Spry.Widget.Utils.destroyWidgets('haactual' + i);
	Spry.Widget.Utils.destroyWidgets('rate' + i);
	Spry.Widget.Utils.destroyWidgets('mandayform' + i);
	Spry.Widget.Utils.destroyWidgets('luasdikerjakan' + i);
	Spry.Widget.Utils.destroyWidgets('rotasi' + i);
	Spry.Widget.Utils.destroyWidgets('amount' + i);
	 $('#trb' + i).remove();
	// getjmlRowbhn();
}

//get activity item
function showActivityrwt(selid) {
	
	$('#activityrwt').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 5,
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

//tampilkan blok
function getTCrwt(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	$('#tcrwt' + i).empty();
	$('#tcrwt' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?activity=' + activityrwt + "&tcby=" + 1,
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

function getHaactual(i)
{
	var tcrwt = $('#tcrwt' + i).val();
	if (tcrwt != "") {
		$.post("bkm_getactivity.asp?tcrwt=" + tcrwt + "&panen=" +  6, function(data) {
			$('#haactual' + i).empty();
			var json  = JSON.parse(data);
			var luas_gross = json.results[0].luas_gross
			$('#haactual' + i).val(luas_gross);
						
		});
		
	}
}

function getRate(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	if (activityrwt != "") {
		$.post("bkm_getactivity.asp?panen=" +  7, function(data) {
			$('#haactual' + i).empty();
			var json  = JSON.parse(data);
			var rate_1 = json.results[0].rate_1
			$('#rate' + i).val(rate_1);
						
		});
		
	}
}

//tampilkan norma mandays
function getNormamandays(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	$('#mandayform' + i).empty();
	$('#mandayform' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?activity=' + activityrwt + "&panen=" + 8,
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

function getTotalhk(i)
{
	var mandayform 		= parseFloat($("#mandayform" + i).val());
	var luasdikerjakan 	= parseFloat($("#luasdikerjakan" + i).val());
	var ttlhkrp			= mandayform*luasdikerjakan;
	$("#ttlhk" + i).val(ttlhkrp);
	
}

function getAmount(i)
{ 
	var rate 			= parseFloat($("#rate" + i).val());
	var ttlhk 			= parseFloat($("#ttlhk" + i).val());
	var rotasi 			= parseFloat($("#rotasi" + i).val());
	var amountttl 		= (rate*ttlhk)*rotasi;
	amountttl			= Math.round(amountttl);
	$("#amount" + i).val(numformat(amountttl));
	$("#amountbck" + i).val((amountttl));
}

function getPuttotal() {
		var jmlluasdikerjakan	= 0;
		var jmlttlhk			= 0;
		var jmlttlrotasi		= 0;
		var jmlttlamount		= 0;
		var jmlrawat			= parseInt($('#jmlrawat').val());
		for(var i=0; i<jmlrawat; i++) {
			// jmlluasdikerjakan	    += parseInt($('#luasdikerjakan'+i).val());
			// jmlttlhk	    		+= parseInt($('#ttlhk'+i).val());
			// jmlttlrotasi	    	+= parseInt($('#rotasi'+i).val());
			jmlttlamount	    	+= parseInt($('#amount'+i).val());
		  	alert(jmlrawat);
			
		}	
		
		// $('#amountluasdikerjakanfrm').val((jmlluasdikerjakan));
		// $('#amountttlhkfrm').val((jmlttlhk));
		// $('#amountrotasifrm').val((jmlttlrotasi));
		$('#amountrkbrpfrm').val((jmlttlamount));
		
	}
 </script>