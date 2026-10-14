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

	.lebarhaactual{
		width: 70px;
	}

	.lebartc{
		width: 50px;
	}

	.lebaractbarang{
		width: 180px;
	}
	.lebarstatus{
		width: 70px;
	}
</style>
<%
	dim rkbpost,norkb
	dim sqlhdr,rshdr 
	dim RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE,ROWSTATE, DOCUMENT_CLASSIFICATION, TOTUPAH,CREATEDATE

	dim classification,simpan,fcba,countrkb,sqlsimpandtl
	DIM statuspost,periodethn,periodebln,section,tglinput,ttlamountupah
	dim sqlnomor,rsnomor,tglprev,currentnomor,currentnomorfix,chargeunit
	DIM actcodex,blok,ttlhk,rate,normamandays,luasdikerjakan,rotasi,totalamounthk
	DIM nQtdFields,nCount,linenokey,norma_hk



	rkbpost = request("rkbno")
	simpan = request("simpan")
	rowstateid 		= Request("rowstateid")

	sqlhdr = "Select RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE, TO_CHAR(CREATEDATE,'DD/MM/YYYY')CREATEDATE, " & _
	" LASTUPDATE, FCENTRY, FCEDIT, FCIP, FCBA, " & _
	" POSTINGFLG, ROWSTATE, DOCUMENT_CLASSIFICATION, TOTUPAH, TOTBAHAN, " & _
	" TOTVRA, TOTWIP " & _
	" From RKBHDR WHERE RKBNO='" & rkbpost & "'"
	set rshdr = connStr.Execute(sqlhdr)
	if not rshdr.eof then
		
		PERIOD_YEAR	= rshdr("PERIOD_YEAR") 
		PERIOD_MONTH= rshdr("PERIOD_MONTH")  
		SECTIONCODE	= rshdr("SECTIONCODE")
		ROWSTATE	= rshdr("ROWSTATE")
		DOCUMENT_CLASSIFICATION	= rshdr("DOCUMENT_CLASSIFICATION")
		TOTUPAH		= rshdr("TOTUPAH")
		CREATEDATE	= rshdr("CREATEDATE")
		fcba		= rshdr("fcba")
	end if

	Dim SQL_Rowsrwt, NoOfRowsrwt,RSCountrwt

	SQL_Rowsrwt = "SELECT Count (*) from RKBUPAHDTL where RKBNO='" & rkbpost & "' AND ADDENDUM=1"
	set RSCountrwt = connStr.Execute(SQL_Rowsrwt)
	if not RSCountrwt.eof then 
		NoOfRowsrwt = RSCountrwt(0).value
	end if
	RSCountrwt.close


	if simpan="SAVE" then


					
			RKBNO 			= request("norkbview")
			fcbatxt 		= Request("fcbatxt")

			nQtdFields 	= request.form("jmlrawat")
			nCount 		= 1
			linenokey  	= 1

		

			do while cint(nCount) <= cint(nQtdFields) 

				tahuntanam 		= request.form("tahuntanam" & nCount)
				actcode 		= request.form("activityrwt" & nCount)
				blok 			= request.form("tcrwt" & nCount)
				rate 			= request.form("rate" & nCount)
				bjrbudget		= request.form("bjrbudget" & nCount)
				kgbudget		= request.form("kgbudget" & nCount)
				sb				= request.form("sb" & nCount)
				lb				= request.form("lb" & nCount)

				rateborongan	= request.form("rateborongan" & nCount)
				totalamountborongan	= request.form("amountboronganbck" & nCount)
				norma_hk		= request.form("norma_hk" & nCount)
				rotasi			= request.form("rotasi" & nCount)

				if rateborongan= "" then
					rateborongan=0
				end if
				if totalamountborongan= "" then
					totalamountborongan=0
				end if

				if rateborongan > 0 then 
					rate = 0
					bjrbudget = 0
					sb = 0
					lb = 0
					norma_hk = 0
					rotasi = 0
				end if

			'	Response.Write tahuntanam
			'	Response.end

				if actcode<>"" then
				
					sqlsimpandtl = "MERGE INTO RKBUPAHDTL  R " & _ 
				" USING (SELECT '" & rkbno & "' RKBNO, '" & actcode & "' ACTCODE, '" & blok & "' LOCATIONCODE, '" & rate & "' RATE, " & _
				"	'" & bjrbudget & "' ABW_BUDGET,'" & kgbudget & "' KG_BUDGET, " & _
				"	'" & sb & "' RATE_SB, '" & lb & "' RATE_LB, '" & linenokey & "' LINENOKEY, " & _
				"	'" & fcbatxt & "' FCBA,'" & tahuntanam & "' TAHUNTANAM,1 ADDENDUM,'" & rateborongan & "' RATEBORONGAN, " & _ 
				"   '" & totalamountborongan & "' TOTALAMOUNT_BORONGAN, '" & norma_hk & "' NORMAMANDAYS,'" & rotasi & "' rotasi " & _
				"	FROM DUAL) RK " & _
				" ON (R.RKBNO = RK.RKBNO AND R.ACTCODE=RK.ACTCODE AND R.LOCATIONCODE=RK.LOCATIONCODE AND R.FCBA=RK.FCBA AND R.ADDENDUM=RK.ADDENDUM) " & _
				" WHEN MATCHED THEN " & _
				"	UPDATE " & _
				"	SET R.RATE_SB=RK.RATE_SB,R.RATE_LB = RK.RATE_LB,R.ABW_BUDGET=RK.ABW_BUDGET,R.KG_BUDGET=RK.KG_BUDGET," & _
				"   LASTUPDATE=SYSDATE,R.TAHUNTANAM=RK.TAHUNTANAM,R.RATEBORONGAN=RK.RATEBORONGAN, " & _ 
				"   R.TOTALAMOUNT_BORONGAN=RK.TOTALAMOUNT_BORONGAN, NORMAMANDAYS=RK.NORMAMANDAYS,ROTASI=RK.ROTASI  " & _ 
				" WHEN NOT MATCHED THEN " & _
				"	INSERT (RKBNO, ACTCODE, LOCATIONCODE, RATE, " & _
					" ABW_BUDGET, KG_BUDGET,  " & _
					" RATE_SB,RATE_LB,LINENOKEY, " & _
					" CREATEDATE, LASTUPDATE, FCENTRY, FCEDIT, FCIP,FCBA,ROWSTATE,ADDENDUM,TAHUNTANAM,RATEBORONGAN,TOTALAMOUNT_BORONGAN,NORMAMANDAYS,ROTASI) " & _
				" VALUES ('" & rkbno & "', '" & actcode & "','" & blok & "','" & rate & "', " & _ 
				" '" & bjrbudget & "','" & kgbudget & "', " & _ 
				" '" & sb & "','" & lb & "','" & linenokey & "',SYSDATE,SYSDATE,'" & Session("CodeName") & "', " & _ 
				" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "', " & _ 
				" '" & fcbatxt & "','Planned',1,'" & tahuntanam & "','" & rateborongan & "','" & totalamountborongan & "','" & norma_hk & "','" & rotasi & "')" 
				ExecuteAndLog "rkb_panen_new_addendum.asp", sqlsimpandtl
				END IF
			
				nCount 	= nCount + 1
				linenokey 	= linenokey + 1
			loop

		response.redirect("rkb_panen_new_adendum.asp?rkbno="&RKBNO&"&rowstateid="&rowstateid)

	End IF
	if simpan="APPROVED" then
		dim sqlupdate
		norkb 			= Request("norkbview")
		

		sqlupdate = "update RKBUPAHDTL set rowstate='Approved' where rkbno='" & norkb & "'"
		ExecuteAndLog "rkb_panen_new_addendum.asp", sqlupdate
		response.redirect("rkb_panen_new_adendum.asp?rkbno="&norkb&"&rowstateid="&rowstateid)
	END IF

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
				<input type="hidden" name="rowstateid" readonly id="rowstateid" size="30" value="<%=ROWSTATE%>" title="ROWSTATE">
			</td>
			
		</tr>

	</table>
	
	<br>
	
	<table align=center border="0" class='data posisikiri' width='100%' id='tableall'>
		
		<tr>
			<td>
			   	<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">RKB Panen</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm wip-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table0'>
								<tr id="trheader0">
									<th>Tahun Tanam</th>
									<th>Aktivitas</th>
									<th>Blok</th>
									<th>Rate</th>
									<th>BJR Budget</th>
									<th>KG/JJG Budget</th>
									<th>SB(Rp.)</th>
									<th>LB(Rp.)</th>
									<th>Status</th>
									<th>Rate Borongan</th>
									<th>Amount Borongan</th>
								</tr>


								<%
									DIM ttlluas,ttlsemuahk,ttlrotasi,ttlsemuaamount
									DIM sqldtl,rsdtl 
									sqldtl = "Select R.TAHUNTANAM, R.RKBNO, R.ACTCODE,J.FCNAME,R.LOCATIONCODE, NVL(R.RATE,0)RATE, " & _ 
									" NVL(R.ABW_BUDGET,0)ABW_BUDGET, NVL(R.KG_BUDGET,0)KG_BUDGET, NVL(R.RATE_SB,0)RATE_SB,NVL(R.RATE_LB,0)RATE_LB,R.ROWSTATE, " & _ 
									" NVL(R.RATEBORONGAN,0)RATEBORONGAN,NVL(R.TOTALAMOUNT_BORONGAN,0)TOTALAMOUNT_BORONGAN,NVL(R.NORMAMANDAYS,0)NORMAMANDAYS,NVL(R.ROTASI,0)ROTASI From RKBUPAHDTL R " & _ 
									" LEFT JOIN JOBMASTER J ON R.ACTCODE=J.FCCODE " & _ 
									" WHERE R.RKBNO='" & rkbpost & "' AND R.ROWSTATE='Approved' AND R.ADDENDUM=0"
									set rsdtl = connStr.Execute(sqldtl)
									'Response.Write sqldtl
																		
									no = 1
									ttlluas = 0
									ttlsemuahk = 0
									ttlrotasi = 0
									ttlsemuaamount = 0
									While Not rsdtl.EOF %>
														

								<tr id="trb0<%=no%>">
									<td class="ratakiri">
											<select class="lebarhaactual " name="tahuntanamv<%=no%>" id="tahuntanamv<%=no%>" >
											<option value=<%=rsdtl("TAHUNTANAM")%>><%=rsdtl("TAHUNTANAM")%></option>
											</select>
										</td>
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
										<input class='lebarhaactual ratatengah' type="text" name="ratev<%=no%>" id="ratev<%=no%>" value="<%=formatnumber(rsdtl("RATE"),0)%>"/>
									</td>
									
									<td>
										<input class="ratatengah lebarnormahk" type="text" name="mandayformv<%=no%>" id="mandayformv<%=no%>" value="<%=formatnumber(rsdtl("ABW_BUDGET"),2)%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text"  name="luasdikerjakanv<%=no%>" id="luasdikerjakanv<%=no%>" value="<%=formatnumber(rsdtl("KG_BUDGET"),2)%>"/>
									</td>
									
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text" name="ttlhkv<%=no%>" id="ttlhkv<%=no%>" value="<%=formatnumber(rsdtl("RATE_SB"),2)%>"/>
									</td>


									<td class="ratakiri">
										<input class="ratatengah lebarlokasi" type="text" name="amountv<%=no%>" id="amountv<%=no%>" value="<%=formatnumber(rsdtl("RATE_LB"),2)%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebarlokasi" type="text" name="rowstatev<%=no%>" id="rowstatev<%=no%>" value="<%=rsdtl("ROWSTATE")%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text" name="rateboronganv<%=no%>" id="rateboronganv<%=no%>" value="<%=rsdtl("RATEBORONGAN")%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratakanan lebarlokasi" readonly type="text" name="amountboronganv<%=no%>" id="amountboronganv<%=no%>" value="<%=formatnumber(rsdtl("TOTALAMOUNT_BORONGAN"),0)%>"/>
										<input class="ratakanan lebarlokasi" readonly type="hidden" name="amountboronganbckv<%=no%>" id="amountboronganbckv<%=no%>" value="<%=rsdtl("TOTALAMOUNT_BORONGAN")%>"/>
									</td>
																
									
									<td></td>
									
								</tr>
								<%
									
									no = no +1 
									rsdtl.MoveNext
									Wend
								
								%>
								
								
							</table> <br>
							
						</div>

					</div>
				</div>
			</td>		
		</tr>
			
	</table><br>
<form name="frmEntry" method="post" action="rkb_panen_new_adendum.asp">
	<input type="hidden" name="jmlrawat" id="jmlrawat" value="<%=NoOfRowsrwt%>" />
	<input type="hidden" name="fcbatxt" readonly id="fcbatxt" value="<%=fcba%>">
	<input type="hidden" name="norkbview" readonly id="norkbview" size="30" value="<%=rkbpost%>" title="No RKB" style='background-color:#eee;'>
	<table align=center border="0" class='data posisikiri' width='100%' id='table1'>
		<tr>
			<td>
					<div class="TabbedPanels" id="tp2" align='center'>
					
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">Addendum RKB Panen</li>
						
					</ul>
					<div class="TabbedPanelsContentGroup">

						<!-- view bkm wip-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table1'>
								<tr id="trheader1">
									<th>Tahun Tanam</th>
									<th>Aktivitas</th>
									<th>Blok</th>
									<th>Rate</th>
									<th>BJR Budget</th>
									<th>KG/JJG Budget</th>
									<th>SB(Rp.)</th>
									<th>LB(Rp.)</th>
									<th>Rate Borongan</th>
									<th>Amount Borongan</th>
									<th colspan="2"></th>
								</tr>

								<%
										'DIM ttlluas,ttlsemuahk,ttlrotasi,ttlsemuaamount
										DIM sqldtl1,rsdtl1 
										sqldtl1 = "Select R.TAHUNTANAM, R.RKBNO, R.ACTCODE,J.FCNAME,R.LOCATIONCODE, NVL(R.RATE,0)RATE, " & _ 
									" NVL(R.ABW_BUDGET,0)ABW_BUDGET, NVL(R.KG_BUDGET,0)KG_BUDGET, NVL(R.RATE_SB,0)RATE_SB,NVL(R.RATE_LB,0)RATE_LB, " & _ 
									" NVL(R.RATEBORONGAN,0)RATEBORONGAN,NVL(R.TOTALAMOUNT_BORONGAN,0)TOTALAMOUNT_BORONGAN,NVL(R.NORMAMANDAYS,0)NORMAMANDAYS,NVL(R.ROTASI,0)ROTASI From RKBUPAHDTL R " & _ 
									" LEFT JOIN JOBMASTER J ON R.ACTCODE=J.FCCODE " & _ 
									" WHERE R.RKBNO='" & rkbpost & "' AND R.ADDENDUM=1"
										set rsdtl1 = connStr.Execute(sqldtl1)
									'	Response.Write sqldtl
																			
										noadd = 1
										ttlluas = 0
										ttlsemuahk = 0
										ttlrotasi = 0
										ttlsemuaamount = 0
										While Not rsdtl1.EOF %>
															

									<tr id="trb<%=noadd%>">
										
										<td class="ratakiri">
											<select class="lebarhaactual " name="tahuntanam<%=noadd%>" id="tahuntanam<%=noadd%>" >
											<option value=<%=rsdtl1("TAHUNTANAM")%>><%=rsdtl1("TAHUNTANAM")%></option>
											</select>
										</td>

										<td class="ratatengah">
											<select class="lebarsection " name="activityrwt<%=noadd%>" id="activityrwt<%=noadd%>" class='lebaractbarang' onchange="getTCrwt(<%=noadd%>)">
											<option value=<%=rsdtl1("ACTCODE")%>><%=rsdtl1("fcname")%></option>
											</select>
										</td>

										<td>
											<select class='lebartc ratatengah' name="tcrwt<%=noadd%>" id="tcrwt<%=noadd%>" >
											<option value=<%=rsdtl1("LOCATIONCODE")%>><%=rsdtl1("LOCATIONCODE")%></option>
										</select>
										</td>

										<td>
										<input class='lebarhaactual ratatengah' readonly type="text" style='background-color:#eee;' name="ratedpn<%=noadd%>" id="ratedpn<%=noadd%>" value="<%=formatnumber(rsdtl1("RATE"),0)%>"/>
										<input class='lebarhaactual ratatengah' readonly type="hidden" style='background-color:#eee;' name="rate<%=noadd%>" id="rate<%=noadd%>" value="<%=rsdtl1("RATE")%>"/>
										<input  type="hidden" style='background-color:#eee;' name="norma_hk<%=noadd%>" id="norma_hk<%=noadd%>" value="<%=formatnumber(rsdtl1("normamandays"))%>"/>
										<input  type="hidden" style='background-color:#eee;' name="rotasi<%=noadd%>" id="rotasi<%=noadd%>" value="<%=formatnumber(rsdtl1("rotasi"))%>"/>
									</td>
									
									<td>
										
										<input class="ratatengah lebarnormahk"  type="text"  name="bjrbudget<%=noadd%>" id="bjrbudget<%=noadd%>" value="<%=formatnumber(rsdtl1("ABW_BUDGET"),2)%>"/>
									</td>
									<td class="ratatengah">
										
										<input class="ratatengah lebarhaactual" type="text" onkeyup="hanyaAngka(this);getTotalhk(<%=noadd%>);" name="kgbudget<%=noadd%>" id="kgbudget<%=noadd%>" value="<%=rsdtl1("KG_BUDGET")%>"/>
									</td>
									
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual"  type="text" onkeyup="getSB(<%=noadd%>);" name="sbdepan<%=noadd%>" id="sbdepan<%=noadd%>" value="<%=formatnumber(rsdtl1("RATE_SB"),0)%>"/>
										<input class="ratatengah lebarhaactual"  type="hidden" name="sb<%=noadd%>" id="sb<%=noadd%>" value="<%=rsdtl1("RATE_SB")%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebartc" type="text" name="lb<%=noadd%>" id="lb<%=noadd%>" value="<%=rsdtl1("RATE_LB")%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text" name="rateborongan<%=noadd%>" id="rateborongan<%=noadd%>" value="<%=rsdtl1("RATEBORONGAN")%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratakanan lebarlokasi" readonly type="text" name="amountborongan<%=noadd%>" id="amountborongan<%=noadd%>" value="<%=formatnumber(rsdtl1("TOTALAMOUNT_BORONGAN"),0)%>"/>
										<input class="ratakanan lebarlokasi" readonly type="hidden" name="amountboronganbck<%=noadd%>" id="amountboronganbck<%=noadd%>" value="<%=rsdtl1("TOTALAMOUNT_BORONGAN")%>"/>
									</td>
																	
										<td>
											<img src='css/login/images/add.jpg' onclick="addrowPanen('<%=noadd%>')" />
										</td>
										<td></td>
										
									</tr>
									<%
									
										noadd = noadd + 1 
										rsdtl1.MoveNext
										Wend
									
									%>
									
								
								
								<tr>
									<td  colspan="10" class="ratatengah"><input type=Submit name="simpan" value="SAVE">&nbsp;&nbsp;<input type=Submit name="simpan" value="APPROVED"></td>
								</tr>
							</table>
							
						</div>

					</div>
				</div>
			</td>
		</tr>
		
	</table>
</form>
	<br><br>
	<table align=center border="1" width="900" class='data' id='table2'>
		<tr>
			<th colspan="3">Summary</th>
		</tr>
		<tr id="trheader2">
			<th>Aktivitas</th>
			<th>Blok</th>
			<th>Kg Budget</th>
			
		</tr>
		
		<%
			dim sqlsum,rssum
			sqlsum = " Select R.RKBNO, R.ACTCODE || ' - ' || J.FCNAME JOBNAME,R.LOCATIONCODE, SUM(NVL(R.KG_BUDGET,0))KG_BUDGET " & _ 
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
					<td align="right" width="90px"><%=formatnumber(rssum("KG_BUDGET"),2)%></td>
			
				</tr>
			<%
				ttlluassummary = cdbl(ttlluassummary) + cdbl(rssum("KG_BUDGET"))
				rssum.MoveNext
				Wend
			%>
			<tr>
				<td colspan="2" class="tebal" style="font-size: 15px;">Total</td>
				<td class="tebal" style="font-size: 15px;"><%=formatnumber(ttlluassummary,2)%></td>
			</tr>
	</table>

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
	addrowPanen();
	
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

function getSB(i)
{
	var sbdepan = $("#sbdepan" + i).val();
	$("#sb" + i).val(sbdepan);
}


//var StartIterationitem=1;
var jumlinv = $('#jmlrawat').val();
var StartIterationitem =  parseInt(jumlinv) + 1
function addrowPanen(trid)
{
	var lstitem = this.StartIterationitem++;
	if (lstitem < 2) {
		var row = $("#trheader1").after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb' + trid).after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowPanen(" + lstitem + ");' />";
	}

	var tahuntanam = "<select name='tahuntanam" + lstitem + "' id='tahuntanam" + lstitem + "' class='lebarhaactual' onchange='getNorm(" + lstitem + ")'></select><input type='hidden' name='norma_hk" + lstitem + "' id='norma_hk" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + tahuntanam + "</td>");

	var activityrwt = "<select name='activityrwt" + lstitem + "' id='activityrwt" + lstitem + "' class='lebarsection' onchange='getTCrwt(" + lstitem + ");getRate(" + lstitem + ");'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + activityrwt + "</td>");

	var tcrwt = "<select name='tcrwt" + lstitem + "' id='tcrwt" + lstitem + "' class='lebartc ratatengah'></select><input type='hidden' name='rotasi" + lstitem + "' id='rotasi" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + tcrwt + "</td>");

	var rate = "<input type='text' name='rate" + lstitem + "' id='rate" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' style='background-color:#eee;' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakanan'>" + rate + "</td>");
	
	var bjrbudget = "<input type='text' name='bjrbudget" + lstitem + "' id='bjrbudget" + lstitem + "' value='0' onkeyup='hanyaAngka(this);' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakanan'>" + bjrbudget + "</td>");
	
	var kgbudget = "<input type='text' name='kgbudget" + lstitem + "' id='kgbudget" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getTotalhk(" + lstitem + ");getAmountborongan(" + lstitem + ");' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + kgbudget + "</td>");

	var sb = "<input type='text' name='sb" + lstitem + "' id='sb" + lstitem + "' value='0' onkeyup='hanyaAngka(this);' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + sb + "</td>");
	
	var lb = "<input type='text' name='lb" + lstitem + "' id='lb" + lstitem + "' value='0' onkeyup='hanyaAngka(this);' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + lb + "</td>");

	var rateborongan = "<input type='text' name='rateborongan" + lstitem + "' id='rateborongan" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmountborongan(" + lstitem + ");' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + rateborongan + "</td>");

	var amountborongan = "<input type='text' name='amountborongan" + lstitem + "' id='amountborongan" + lstitem + "' style='background-color:#eee;' value='0' onkeyup='hanyaAngka(this)' class='lebarlokasi ratakanan'><input type='hidden' name='amountboronganbck" + lstitem + "' id='amountboronganbck" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + amountborongan + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowPanen(" + lstitem + ");' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jmlrawat').val(lstitem);

	showActivityrwt("activityrwt" + lstitem);
	showTahuntanam("tahuntanam" + lstitem);

	
}

//hapus row rkbperawatan
function deleterowPanen(i) {
	Spry.Widget.Utils.destroyWidgets('tahuntanam' + i);
	Spry.Widget.Utils.destroyWidgets('activityrwt' + i);
	Spry.Widget.Utils.destroyWidgets('tcrwt' + i);
	Spry.Widget.Utils.destroyWidgets('rate' + i);
	Spry.Widget.Utils.destroyWidgets('bjrbudget' + i);
	Spry.Widget.Utils.destroyWidgets('kgbudget' + i);
	Spry.Widget.Utils.destroyWidgets('sb' + i);
	Spry.Widget.Utils.destroyWidgets('lb' + i);
	Spry.Widget.Utils.destroyWidgets('rateborongan' + i);
	Spry.Widget.Utils.destroyWidgets('amountborongan' + i);
	Spry.Widget.Utils.destroyWidgets('norma_hk' + i);
	Spry.Widget.Utils.destroyWidgets('rotasi' + i);
	 $('#trb' + i).remove();
	// getjmlRowbhn();
}

//get tahun tanam item
function showTahuntanam(selid) {
	
	$('#tahuntanam').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 12,
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

//get activity item
function showActivityrwt(selid) {
	
	$('#activityrwt').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 11,
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
	var tahun = $('#tahun').val();
	var bulan = $('#bulan').val();
	var sectionname = $('#sectionname').val();
	var activityrwt = $('#activityrwt' + i).val();
	$('#tcrwt' + i).empty();
	$('#tcrwt' + i).select2({
		ajax: {
			delay: 1000,
		//	url : 'bkm_getblok.asp?activity=' + activityrwt + "&tcby=" + 1,
			url : 'bkm_getblok_rkb.asp?activity=' + activityrwt + "&tcby=" + 1 + "&tahun=" + tahun + "&bulan=" + bulan + "&section=" + sectionname,
			
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


function getRate(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	if (activityrwt != "") {
		$.post("bkm_getactivity.asp?panen=" +  7, function(data) {
			var json  = JSON.parse(data);
			var rate_1 = json.results[0].rate_1
			$('#rate' + i).val(rate_1);
						
		});
		
	}
}

function getNorm(i)
{
	var tahuntanam = $('#tahuntanam' + i).val();
	if (tahuntanam != "") {
		$.post("bkm_getactivity.asp?panen=" +  20 + "&tahuntanam=" + tahuntanam, function(data) {
			//$('#rate' + i).empty();
			var json  = JSON.parse(data);
			var norma_hk = json.results[0].norma_hk
			var rotasi = json.results[0].rotasi
			$('#rotasi' + i).val(rotasi);
			$('#norma_hk' + i).val(norma_hk);			
		});
		
	}
}

function getTotalhk(i)
{
	var mandayform 		= parseFloat($("#mandayform" + i).val());
	var luasdikerjakan 	= parseFloat($("#luasdikerjakan" + i).val());
	var ttlhkrp			= parseFloat(mandayform)*parseFloat(luasdikerjakan);
	$("#ttlhk" + i).val(ttlhkrp);
	//alert(ttlhkrp);
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


function getAmountborongan(i)
{ 

	var rateborongan 	= parseFloat($("#rateborongan" + i).val());
	var kgbudget 		= parseFloat($("#kgbudget" + i).val());
	var amountttl 		= (rateborongan*kgbudget);
	amountttl			= Math.round(amountttl);
	$("#amountborongan" + i).val(numformat(amountttl));
	$("#amountboronganbck" + i).val((amountttl));
	//alert(rate);
}


 </script>