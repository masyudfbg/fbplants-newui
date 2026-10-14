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
		.lebaractbarang{
		width: 180px;
	}

	.lebaractbarang1{
		width: 200px;
	}

	.lebaritembarang{
		width: 200px;
	}
	
	.lebaruom{
		width: 50px;
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
	DIM nQtdFields,nCount,linenokey,rateborongan,totalamountborongan



	rkbpost = request("rkbno")
	simpan = request("simpan")
	rowstateid 		= Request("rowstateid")

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
			fcba 			= session("fcba")

			'looping di tabel detail
			nQtdFields 	= request.form("jmlrawat")
			nCount 		= 1
			linenokey  	= 1

		

			do while cint(nCount) <= cint(nQtdFields) 

				'ttlhk 			= replace('" & AMOUNT & "',',','')

				actcodex 		= request.form("activityrwt" & nCount)
				blok 			= request.form("tcrwt" & nCount)
				ttlhk 			= replace(request.form("ttlhk" & nCount),",","")
				rate 			= request.form("rate" & nCount)
				normamandays	= request.form("mandayform" & nCount)
				luasdikerjakan	= request.form("luasdikerjakan" & nCount)
				rotasi			= request.form("rotasi" & nCount)
				totalamounthk	= request.form("amountbck" & nCount)
				haactual 		= request.form("haactual" & nCount)
				rateborongan	= request.form("rateborongan" & nCount)
				totalamountborongan	= request.form("amountboronganbck" & nCount)

				
				if ttlhk= "" then
					ttlhk=0
				end if
				if haactual= "" then
					haactual=0
				end if
				if rate= "" then
					rate=0
				end if
				if normamandays= "" then
					normamandays=0
				end if
				if luasdikerjakan= "" then
					luasdikerjakan=0
				end if
				if rotasi= "" then
					rotasi=0
				end if
				if totalamounthk= "" then
					totalamounthk=0
				end if
				if rateborongan= "" then
					rateborongan=0
				end if
				if totalamountborongan= "" then
					totalamountborongan=0
				end if

				if rateborongan > 0 then 
					ttlhk = 0
				end if


				if actcodex<>"" then

				sqlsimpandtl = "MERGE INTO RKBUPAHDTL  R " & _ 
				" USING (SELECT '" & RKBNO & "' RKBNO, '" & actcodex & "' ACTCODE, '" & blok & "' LOCATIONCODE, '" & ttlhk & "' MANDAYS, '" & rate & "' RATE, " & _
				"	'" & normamandays & "' NORMAMANDAYS,'" & luasdikerjakan & "' VOLUME, " & _
				"	'" & rotasi & "' ROTASI, '" & totalamounthk & "' TOTALAMOUNT_HK, '" & linenokey & "' LINENOKEY, " & _
				"	'" & fcba & "' FCBA, '" & haactual & "' HAACTUAL,'" & rateborongan & "' RATEBORONGAN,'" & totalamountborongan & "' TOTALAMOUNT_BORONGAN,1 ADDENDUM " & _
				"	FROM DUAL) RK " & _
				" ON (R.RKBNO = RK.RKBNO AND R.ACTCODE=RK.ACTCODE AND R.LOCATIONCODE=RK.LOCATIONCODE AND R.FCBA=RK.FCBA AND R.ADDENDUM=RK.ADDENDUM) " & _
				" WHEN MATCHED THEN " & _
				"	UPDATE " & _
				"	SET R.NORMAMANDAYS=RK.NORMAMANDAYS,R.VOLUME = RK.VOLUME,R.MANDAYS=RK.MANDAYS,R.ROTASI=RK.ROTASI,R.TOTALAMOUNT_HK=RK.TOTALAMOUNT_HK, " & _
				"   LASTUPDATE=SYSDATE,R.RATEBORONGAN=RK.RATEBORONGAN,R.TOTALAMOUNT_BORONGAN=RK.TOTALAMOUNT_BORONGAN " & _
				" WHEN NOT MATCHED THEN " & _
				"	INSERT (RKBNO, ACTCODE, LOCATIONCODE, MANDAYS, RATE,  " & _
				" NORMAMANDAYS, VOLUME,ROTASI,   " & _
				" TOTALAMOUNT_HK, LINENOKEY,  " & _
				"  CREATEDATE,LASTUPDATE, FCENTRY, FCEDIT, FCIP,  " & _ 
				" FCBA, HAACTUAL,ROWSTATE,ADDENDUM,RATEBORONGAN,TOTALAMOUNT_BORONGAN) " & _
				" VALUES ('" & RKBNO & "', '" & actcodex & "','" & blok & "','" & ttlhk & "','" & rate & "', " & _ 
				" '" & normamandays & "','" & luasdikerjakan & "', " & _ 
				" '" & rotasi & "','" & totalamounthk & "','" & linenokey & "',SYSDATE,SYSDATE,'" & Session("CodeName") & "', " & _ 
				" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & haactual & "','Planned',1,'" & rateborongan & "','" & totalamountborongan & "')" 
				ExecuteAndLog "rkb_perawatan_adendum.asp", sqlsimpandtl
				end if
			
				nCount 	= nCount + 1
				linenokey 	= linenokey + 1
			loop

		response.redirect("rkb_perawatan_adendum.asp?rkbno="&RKBNO&"&rowstateid="&rowstateid)

	End IF

	if simpan="APPROVED" then
		dim sqlupdate
		norkb 			= Request("norkbview")
		'Response.Write norkb 
		'Response.end
		'sqlupdate = "update rkbhdr set rowstate='Approved' where rkbno='" & norkb & "'"
		'connStr.execute(sqlupdate)

		sqlupdate = "update RKBUPAHDTL set rowstate='Approved' where rkbno='" & norkb & "'"
		ExecuteAndLog "rkb_perawatan_adendum.asp", sqlupdate
		response.redirect("rkb_perawatan_adendum.asp?rkbno="&norkb&"&rowstateid="&rowstateid)
	END IF

	'batas rkb upah

	Dim SQL_Rowsitem, NoOfRowsitem,RSCountitem

	SQL_Rowsitem = "SELECT Count (*) from RKBBAHANDTL where RKBNO='" & rkbpost & "'"
	set RSCountitem = connStr.Execute(SQL_Rowsitem)
	if not RSCountitem.eof then 
		NoOfRowsitem = RSCountitem(0).value
	end if
	RSCountitem.close

	' simpan bahan
	if simpan = "SAVE BAHAN" then
		dim sqlbahan,nQtdFieldsbahan,nCountbahan,linenokeybahan
		RKBNO 				= request("norkbview")
		fcba 				= session("fcba")
		nQtdFieldsbahan 	= request.form("jumldataitem")
		nCountbahan 		= 1
		linenokeybahan 		= 1	

		do while cint(nCountbahan) <= cint(nQtdFieldsbahan)

			dim itembarang 
			itembarang = request.form("itemcode" & nCountbahan)
			if itembarang<>"" then 

				sqlbahan = " insert into RKBBAHANDTL (RKBNO,ACTCODE,LOCATIONCODE,STOCKCODE, " & _ 
				" UOM,QTY,CREATEDATE,LASTUPDATE,FCENTRY, " & _
				" FCEDIT,FCIP,FCBA,LINENOKEY,RATESTOCK,TOTALAMOUNT,ACTIVEMATERIALNORM) " & _
				" values ('" & rkbno & "','" & request.form("activityitem" & nCountbahan) & "','" & request.form("tcitem" & nCountbahan) & "', " & _ 
				" '" & request.form("itemcode" & nCountbahan) & "','" & request.form("uom" & nCountbahan) & "', " & _ 
				" '" & request.form("qtyitem" & nCountbahan) & "',SYSDATE,SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
				" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & linenokeybahan & "','" & request.form("rateitem" & nCountbahan) & "', " & _ 
				" '" & request.form("amountitem" & nCountbahan) & "','" & request.form("normamm" & nCountbahan) & "')"
				ExecuteAndLog "rkb_perawatan_adendum.asp", sqlbahan

			end if

			nCountbahan 	= nCountbahan + 1
			linenokeybahan 	= linenokeybahan + 1

		loop
		response.redirect("rkb_perawatan_adendum.asp?rkbno="&RKBNO&"&rowstateid="&rowstateid)
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
						<li class="TabbedPanelsTab">RKB Perawatan</li>
						<li class="TabbedPanelsTab">RKB Bahan</li>
						
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
									<th>Rate Borongan</th>
									<th>Amount Borongan</th>
									<th colspan="2"></th>
								</tr>


								<%
									DIM ttlluas,ttlsemuahk,ttlrotasi,ttlsemuaamount
									DIM sqldtl,rsdtl 
									sqldtl = "Select R.RKBNO, R.ACTCODE,J.FCNAME,R.LOCATIONCODE, NVL(R.MANDAYS,0)MANDAYS, NVL(R.RATE,0)RATE, " & _ 
									" NVL(R.NORMAMANDAYS,0)NORMAMANDAYS, NVL(R.VOLUME,0)VOLUME, NVL(R.ROTASI,0)ROTASI,NVL(R.TOTALAMOUNT_HK,0)TOTALAMOUNT_HK, R.LINENOKEY,R.FCBA,NVL(R.HAACTUAL,0)HAACTUAL,R.ROWSTATE, " & _ 
									" NVL(R.RATEBORONGAN,0)RATEBORONGAN,NVL(R.TOTALAMOUNT_BORONGAN,0)TOTALAMOUNT_BORONGAN From   RKBUPAHDTL R " & _ 
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
										<input class="ratatengah lebarhaactual" type="text" name="haactualv<%=no%>" id="haactualv<%=no%>" value="<%=formatnumber(rsdtl("HAACTUAL"),2)%>"/>
									</td>

									<td>
										<input class='lebarhaactual ratatengah' type="text" name="ratev<%=no%>" id="ratev<%=no%>" value="<%=formatnumber(rsdtl("RATE"),0)%>"/>
									</td>
									
									<td>
										<input class="ratatengah lebarnormahk" type="text" name="mandayformv<%=no%>" id="mandayformv<%=no%>" value="<%=formatnumber(rsdtl("NORMAMANDAYS"),4)%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text"  name="luasdikerjakanv<%=no%>" id="luasdikerjakanv<%=no%>" value="<%=formatnumber(rsdtl("VOLUME"),2)%>"/>
									</td>
									
									<td class="ratakiri">
										<input class="ratatengah lebarhaactual" type="text" name="ttlhkv<%=no%>" id="ttlhkv<%=no%>" value="<%=formatnumber(rsdtl("MANDAYS"),4)%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebartc" type="text" name="rotasiv<%=no%>" id="rotasiv<%=no%>" value="<%=rsdtl("ROTASI")%>"/>
									</td>

									<td class="ratakiri">
										<input class="ratatengah lebarlokasi" type="text" name="amountv<%=no%>" id="amountv<%=no%>" value="<%=formatnumber(rsdtl("TOTALAMOUNT_HK"),0)%>"/>
									</td>
									<td class="ratakiri">
										<input class="ratakiri lebarlokasi" type="text" name="rowstatev<%=no%>" id="rowstatev<%=no%>" value="<%=rsdtl("ROWSTATE")%>"/>
									</td>
																
									<td class="ratakiri">
										<input class="ratatengah lebartc" type="text" name="rateboronganv<%=no%>" id="rateboronganv<%=no%>" value="<%=rsdtl("RATEBORONGAN")%>"/>
										</td>

										<td class="ratakiri">
											<input class="ratakanan lebarlokasi" readonly type="text" name="amountboronganv<%=no%>" id="amountboronganv<%=no%>" value="<%=formatnumber(rsdtl("TOTALAMOUNT_BORONGAN"),0)%>"/>
											<input class="ratakanan lebarlokasi" readonly type="hidden" name="amountboronganbckv<%=no%>" id="amountboronganbckv<%=no%>" value="<%=rsdtl("TOTALAMOUNT_BORONGAN")%>"/>
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

						<div class="TabbedPanelsContent">
								<!-- ============================================================================================================================= -->
								<div id="datadetailbhn">
									<table align=center border="0"  width="850" class='data' id='table1'>
										<tr id="trheader2">
											<th>Activity</th>
											<th>TC</th>
											<th>Item</th>
											<th>Uom</th>
											<th>Norm</th>
											<th>Qty</th>
											<th>Rate</th>
											<th>Amount</th>
										
										</tr>

							
										<%
											DIM ttlqtyitem,ttlamountitem,nobr
											DIM sqlitemdtl,rsitemdtl 
											sqlitemdtl = "Select R.RKBNO, R.ACTCODE,J.FCCODE || ' - ' || J.FCNAME ACTNAME, R.LOCATIONCODE, R.STOCKCODE,A.FCCODE || ' - ' || A.FCNAME ACTMMNAME, R.QTY, " & _ 
											" R.UOM, R.ACTIVEMATERIALNORM, R.HAPLAN, R.ROTASI, R.RATESTOCK, " & _
											" R.TOTALAMOUNT, R.LINENOKEY " & _
											" From   RKBBAHANDTL R " & _
											" LEFT JOIN JOBMASTER J ON R.ACTCODE=J.FCCODE " & _
											" LEFT JOIN ACTIVEMATERIAL A ON R.STOCKCODE=A.FCCODE " & _
											" WHERE R.RKBNO='" & rkbpost & "'"
											' response.write sqlitemdtl
											set rsitemdtl = connStr.Execute(sqlitemdtl)

											nobr = 0
											ttlqtyitem = 0
											ttlamountitem = 0
											While Not rsitemdtl.EOF %>

											<tr id='trbbh<%=nobr%>'>
																						
												<td class="ratakiri">
													<select class="lebaractbarang ratakiri" onchange="getTCitem(<%=nobr%>);showItem(<%=nobr%>)" name="activityitembh<%=nobr%>" id="activityitembh<%=nobr%>">
														<option value=<%=rsitemdtl("ACTCODE")%>><%=rsitemdtl("ACTNAME")%></option>
													</select>
												</td>
												<td class="ratatengah">
													<select class="lebaruom ratakiri" name="tcitembh<%=nobr%>" id="tcitembh<%=nobr%>">
														<option value=<%=rsitemdtl("LOCATIONCODE")%>><%=rsitemdtl("LOCATIONCODE")%></option>
													</select>
												</td>
												<td class="ratakiri">
													<select class="lebaritembarang " name="itemcodebh<%=nobr%>" id="itemcodebh<%=nobr%>">
														<option value=<%=rsitemdtl("STOCKCODE")%>><%=rsitemdtl("ACTMMNAME")%></option>
													</select>
												</td>
												<td class="ratatengah">
													<input class="ratatengah lebaruom" type="text" name="uombh<%=nobr%>" id="uombh<%=nobr%>" value="<%=rsitemdtl("UOM")%>"/>
												</td>
												<td class="ratatengah">
													<input class="ratatengah lebaruom" type="text" name="normammbh<%=nobr%>" id="normammbh<%=nobr%>" value="<%=formatnumber(rsitemdtl("ACTIVEMATERIALNORM"))%>"/>
												</td>
												<td class="ratatengah">
													<input class="ratatengah lebaruom" type="text" name="qtyitembh<%=nobr%>" id="qtyitembh<%=nobr%>" value="<%=formatnumber(rsitemdtl("qty"))%>"/>
												</td>
												<td class="ratatengah">
													<input class="ratatengah lebarnormahk" type="text" name="rateitembh<%=nobr%>" id="rateitembh<%=nobr%>" value="<%=rsitemdtl("RATESTOCK")%>"/>
												</td>
												<td class="ratatengah">
													<input class="ratatengah lebartc" type="text" name="amountitemfrmbh<%=nobr%>" id="amountitemfrmbh<%=nobr%>" value="<%=formatnumber(rsitemdtl("TOTALAMOUNT"),0)%>"/>
													<input class="ratatengah lebartc" type="hidden" name="amountitembh<%=nobr%>" id="amountitembh<%=nobr%>" value="<%=(rsitemdtl("TOTALAMOUNT"))%>"/>
												</td>

									
										</tr>

										<%
											ttlqtyitem = cdbl(ttlqtyitem) + cdbl(rsitemdtl("qty"))
											ttlamountitem = cdbl(ttlamountitem) + cdbl(rsitemdtl("TOTALAMOUNT"))
										
											nobr = nobr + 1 
											rsitemdtl.MoveNext
											Wend
										
										%>

										
									</table>
								</div>
								<!-- ============================================================================================================================= -->
						</div>

					</div>
				</div>
			</td>		
		</tr>
			
	</table><br>
<form name="frmEntry" method="post" action="rkb_perawatan_adendum.asp">
	<input type="hidden" name="jmlrawat" id="jmlrawat" value="<%=NoOfRowsrwt%>" />
	<input type="hidden" name="norkbview" readonly id="norkbview" size="30" value="<%=rkbpost%>" title="No RKB" style='background-color:#eee;'>
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="0" />
	<table align=center border="0" class='data posisikiri' width='100%' id='table1'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp2" align='center'>
					
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">Addendum RKB Perawatan</li>
						<li class="TabbedPanelsTab">Addendum RKB Bahan</li>
						
					</ul>
					<div class="TabbedPanelsContentGroup">

						<!-- view bkm bahan-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='tableprwt'>
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
									<th>Rate Borongan</th>
									<th>Amount Borongan</th>
									<th>Status</th>
									<th colspan="2"></th>
								</tr>

								<%
										'DIM ttlluas,ttlsemuahk,ttlrotasi,ttlsemuaamount
										DIM sqldtl1,rsdtl1 
										sqldtl1 = "Select R.RKBNO, R.ACTCODE,J.FCNAME,R.LOCATIONCODE, NVL(R.MANDAYS,0)MANDAYSADDENDUM, NVL(R.RATE,0)RATE, " & _ 
										" NVL(R.NORMAMANDAYS,0)NORMAMANDAYS, ROUND(NVL(R.VOLUME,0))VOLUME, NVL(R.ROTASI,0)ROTASI,ROUND(NVL(R.TOTALAMOUNT_HK,0))TOTALAMOUNT_HK, R.LINENOKEY,R.FCBA,NVL(R.HAACTUAL,0)HAACTUAL,R.ROWSTATE, " & _ 
										" NVL(R.RATEBORONGAN,0)RATEBORONGAN,NVL(R.TOTALAMOUNT_BORONGAN,0)TOTALAMOUNT_BORONGAN From  RKBUPAHDTL R " & _ 
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
											<select class="lebaractivity " name="activityrwt<%=noadd%>" id="activityrwtv<%=noadd%>" class='lebaractbarang' onchange="getTCrwt(<%=noadd%>);getNormamandays(<%=noadd%>)">
												<option value=<%=rsdtl1("ACTCODE")%>><%=rsdtl1("fcname")%></option>
											</select>
										</td>

										<td class="ratatengah">
											<select class='lebartc ratatengah' name="tcrwt<%=noadd%>" id="tcrwtv<%=noadd%>" >
												<option value=<%=rsdtl1("LOCATIONCODE")%>><%=rsdtl1("LOCATIONCODE")%></option>
											</select>
										</td>

										<td>
											<input class="ratatengah lebarhaactual" type="text" name="haactual<%=noadd%>" id="haactualdpnv<%=noadd%>" value="<%=formatnumber(rsdtl1("HAACTUAL"),0)%>"/>
											
										</td>

										<td>
											<input class='lebarhaactual ratatengah' type="text" name="ratedpn<%=noadd%>" id="ratedpn<%=noadd%>" value="<%=formatnumber(rsdtl1("RATE"),0)%>"/>
											<input class='lebarhaactual ratatengah' type="hidden" name="rate<%=noadd%>" id="rate<%=noadd%>" value="<%=rsdtl1("RATE")%>"/>
										</td>
										
										<td>
											<input class="ratatengah lebarnormahk" type="text" name="mandayform<%=noadd%>" id="mandayform<%=noadd%>" value="<%=formatnumber(rsdtl1("NORMAMANDAYS"),2)%>"/>
										
										</td>
										<td class="ratakiri">
											<input class="ratatengah lebarhaactual" type="text" onkeyup="hanyaAngka(this);getTotalhk(<%=noadd%>);getAmountborongan(<%=noadd%>);"  name="luasdikerjakan<%=noadd%>" id="luasdikerjakan<%=noadd%>" value="<%=rsdtl1("VOLUME")%>"/>
											
										</td>
										
										<td class="ratakiri">
											<input class="ratatengah lebarhaactual" type="text"  onkeyup='hanyaAngka(this);getAmount(<%=noadd%>);' name="ttlhk<%=noadd%>" id="ttlhk<%=noadd%>" value="<%=FormatNumber(rsdtl1("MANDAYSADDENDUM"))%>"/>
										
										</td>

										<td class="ratakiri">
											<input class="ratatengah lebartc" type="text" onkeyup='hanyaAngka(this);getAmount(<%=noadd%>);' name="rotasi<%=noadd%>" id="rotasi<%=noadd%>" value="<%=rsdtl1("ROTASI")%>"/>
										</td>

										<td class="ratakiri">
											<input class="ratakanan lebarlokasi" type="text" name="amountdpn<%=noadd%>" id="amount<%=noadd%>" value="<%=formatnumber(rsdtl1("TOTALAMOUNT_HK"),0)%>"/>
											<input class="ratatengah lebarlokasi" type="hidden" name="amountbck<%=noadd%>" id="amountbck<%=noadd%>" value="<%=rsdtl1("TOTALAMOUNT_HK")%>"/>
										</td>
										<td class="ratakiri">
										<input class="ratatengah lebartc" type="text" name="rateborongan<%=noadd%>" id="rateborongan<%=noadd%>" value="<%=rsdtl1("RATEBORONGAN")%>"/>
										</td>

										<td class="ratakiri">
											<input class="ratakanan lebarlokasi" readonly type="text" name="amountborongan<%=noadd%>" id="amountborongan<%=noadd%>" value="<%=formatnumber(rsdtl1("TOTALAMOUNT_BORONGAN"),0)%>"/>
											<input class="ratakanan lebarlokasi" readonly type="hidden" name="amountboronganbck<%=noadd%>" id="amountboronganbck<%=noadd%>" value="<%=rsdtl1("TOTALAMOUNT_BORONGAN")%>"/>
										</td>
								
										<td class="ratakiri">
											<input class="ratakiri lebarstatus" type="text" name="rowstate<%=noadd%>" id="rowstatev<%=noadd%>" value="<%=rsdtl1("ROWSTATE")%>"/>
										</td>
																	
										<td>
											<img src='css/login/images/add.jpg' onclick="addrowPerawatan('<%=noadd%>')" />
										</td>
										<td></td>
										
									</tr>
									<%
										ttlluas = cdbl(ttlluas) + cdbl(rsdtl1("VOLUME"))
										ttlsemuahk = cdbl(ttlsemuahk) + cdbl(rsdtl1("MANDAYSADDENDUM"))
										ttlrotasi = cdbl(ttlrotasi) + cdbl(rsdtl1("ROTASI"))
										ttlsemuaamount = cdbl(ttlsemuaamount) + cdbl(rsdtl1("TOTALAMOUNT_HK"))
										noadd = noadd + 1 
										rsdtl1.MoveNext
										Wend
									
									%>
									
								
								
								<tr>
									<td  colspan="10" class="ratatengah"><input type=Submit name="simpan" value="SAVE">&nbsp;&nbsp;<input type=Submit name="simpan" value="APPROVED"></td>
								</tr>
							</table>
							
						</div>

						<div class="TabbedPanelsContent">
								<!-- ============================================================================================================================= -->
								<div id="addbhn">
									<table align=center border="0"  width="850" class='data' id='tableaddbhn'>
										<tr id="trheader3">
											<th>Activity</th>
											<th>TC</th>
											<th>Item</th>
											<th>Uom</th>
											<th>Norm</th>
											<th>Qty</th>
											<th>Rate</th>
											<th>Amount</th>
											<th colspan="2">&nbsp;</th>
										</tr>

										<tr id='trbadd'>
												<tbody id="bodytablebrg">
												</tbody>
												
										</tr>

										<tr>
											<td colspan='8' align=center><input Type="Submit" id="SIMPAN_BAHAN" name="simpan" value="SAVE BAHAN"  />
											&nbsp;&nbsp;</td>
											
										</tr>
										
									</table>
								</div>
								<!-- ============================================================================================================================= -->
						</div>

					</div>
				</div>
			</td>
		</tr>
		
	</table>
</form>
	<br><br>
	<table align=center border="0" width="900" class='data' id='table2'>
		
		<tr>
			<td>
				<div class="TabbedPanels" id="tp3" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">Summary RKB Perawatan</li>
						<li class="TabbedPanelsTab">Summary RKB Bahan</li>
						
					</ul>
					<div class="TabbedPanelsContentGroup">
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table2'>
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
											<td align="right" width="90px"><%=formatnumber(rssum("VOLUME"),2)%></td>
									
										</tr>
									<%
										ttlluassummary = cdbl(ttlluassummary) + cdbl(rssum("VOLUME"))
										rssum.MoveNext
										Wend
									%>
									<tr>
										<td colspan="2" class="tebal" style="font-size: 15px;">Total Perawatan</td>
										<td class="tebal" style="font-size: 15px;"><%=formatnumber(ttlluassummary,2)%></td>
									</tr>
							</table>
						</div>

						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table3'>
								<tr id="trheader2">
									<th>Aktivitas</th>
									<th>Blok</th>
									<th>Bahan</th>
									<th>Qty</th>
								</tr>
								
								<%
									
									sqlsum = " Select R.RKBNO, R.ACTCODE || ' - ' || J.FCNAME JOBNAME,R.LOCATIONCODE,A.FCCODE || ' - ' || A.FCNAME ACTMMNAME, SUM(NVL(R.QTY,0))QTY " & _ 
									" From   RKBBAHANDTL R  " & _ 
									" LEFT JOIN JOBMASTER J ON R.ACTCODE=J.FCCODE  " & _ 
									" LEFT JOIN ACTIVEMATERIAL A ON R.STOCKCODE=A.FCCODE " & _
									" WHERE R.RKBNO='" & rkbpost & "' " & _ 
									" GROUP BY R.RKBNO, R.ACTCODE,J.FCNAME,R.LOCATIONCODE,A.FCCODE || ' - ' || A.FCNAME"
									set rssum = connStr.Execute(sqlsum)
									'Response.Write sqldtl
																		
									ttlluassummary = 0
									While Not rssum.EOF %>
										<tr style="font-size: 15px;">
											<td align="left" width="220px"><%=rssum("JOBNAME")%></td>
											<td align="center" width="40px"><%=rssum("LOCATIONCODE")%></td>
											<td align="left" width="220px"><%=rssum("ACTMMNAME")%></td>
											<td align="right" width="40px"><%=formatnumber(rssum("QTY"),2)%></td>
									
										</tr>
									<%
										ttlluassummary = cdbl(ttlluassummary) + cdbl(rssum("QTY"))
										rssum.MoveNext
										Wend
									%>
									<tr>
										<td colspan="3" class="tebal" style="font-size: 15px;">Total Bahan</td>
										<td class="tebal" style="font-size: 15px;"><%=formatnumber(ttlluassummary,2)%></td>
									</tr>
							</table>
						</div>
					</div>
				</div>
			</td>
			
		</tr>


 </div>
 </div>
 </body>
 </head>
 <link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script type="text/javascript">
new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});
new Spry.Widget.TabbedPanels("tp2", {
		defaultTab: 0
	});
new Spry.Widget.TabbedPanels("tp3", {
		defaultTab: 0
	});


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
	addrowPerawatan();
	addrowBahan();
	
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


//batas script rkb perawatan
var jumlrwt = $('#jmlrawat').val();	
if (jumlrwt=="0"){
		var StartIterationrwt = 1;
		addrowPerawatan(1);
}else{
	var StartIterationrwt =  parseInt(jumlrwt) + 1
}
function addrowPerawatan(trid)
{
	var lstitem = this.StartIterationrwt++;
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

	var rateborongan = "<input type='text' name='rateborongan" + lstitem + "' id='rateborongan" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmountborongan(" + lstitem + ");' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + rateborongan + "</td>");

	var amountborongan = "<input type='text' name='amountborongan" + lstitem + "' id='amountborongan" + lstitem + "' style='background-color:#eee;' value='0' onkeyup='hanyaAngka(this)' class='lebarlokasi ratakanan'><input type='hidden' name='amountboronganbck" + lstitem + "' id='amountboronganbck" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + amountborongan + "</td>");

	var status = "";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + status + "</td>");

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
	//Spry.Widget.Utils.destroyWidgets('status' + i);
	Spry.Widget.Utils.destroyWidgets('amount' + i);
	Spry.Widget.Utils.destroyWidgets('rateborongan' + i);
	Spry.Widget.Utils.destroyWidgets('amountborongan' + i);
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
	var tahun = $('#tahun').val();
	var bulan = $('#bulan').val();
	var sectionname = $('#sectionname').val();
	var activityrwt = $('#activityrwt' + i).val();
	$('#tcrwt' + i).empty();
	$('#tcrwt' + i).select2({
		ajax: {
			delay: 1000,
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
	var rateborongan 			= parseFloat($("#rateborongan" + i).val());
	var luasdikerjakan 			= parseFloat($("#luasdikerjakan" + i).val());
	var amountttl 		= (rateborongan*luasdikerjakan);
	amountttl			= Math.round(amountttl);
	$("#amountborongan" + i).val(numformat(amountttl));
	$("#amountboronganbck" + i).val((amountttl));
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

// BATAS RKB UPAH ADDENDUM
//batas script rkb bahan
var jumlitem = parseInt($('#jumldataitem').val());	
if (jumlitem=="0"){
		var StartIterationitem = 1;
		addrowBahan(1);
}else{
	var StartIterationitem =  parseInt(jumlitem) + 1
}
function addrowBahan(trid)
{

	var lstitem = this.StartIterationitem++;
	
	if (lstitem < 2) {
		var row = $("#trheader3").after("<tr id='trbadd" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trbadd' + trid).after("<tr id='trbadd" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowBahan(" + lstitem + ");' />";
	}

	var activityitem = "<select name='activityitem" + lstitem + "' id='activityitem" + lstitem + "' class='lebaractbarang' onchange='getTCitem(" + lstitem + ");showItem(" + lstitem + ")'></select>";
	$('#trbadd' + lstitem).append("<td class='ratakiri'>" + activityitem + "</td>");

	var tcitem = "<select name='tcitem" + lstitem + "' id='tcitem" + lstitem + "' class='lebaruom ratatengah'></select>";
	$('#trbadd' + lstitem).append("<td class='ratakiri'>" + tcitem + "</td>");

	var itemcode = "<select name='itemcode" + lstitem + "' id='itemcode" + lstitem + "' onchange='getSatuan(" + lstitem + ")' class='lebaritembarang'></select>";
	$('#trbadd' + lstitem).append("<td class='ratakiri'>" + itemcode + "</td>");

	var uom = "<input type='text' name='uom" + lstitem + "' id='uom" + lstitem + "' value='' readonly class='lebaruom ratatengah'>";
	$('#trbadd' + lstitem).append("<td class='ratakiri'>" + uom + "</td>");

	var normamm = "<input type='text' name='normamm" + lstitem + "' id='normamm" + lstitem + "' value='0' class='lebaruom ratatengah'>";
	$('#trbadd' + lstitem).append("<td class='ratakiri'>" + normamm + "</td>");

	var qtyitem = "<input type='text' name='qtyitem" + lstitem + "' id='qtyitem" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmountitem(" + lstitem + ")' class='lebaruom ratatengah'>";
	$('#trbadd' + lstitem).append("<td class='ratakiri'>" + qtyitem + "</td>");

	var rateitem = "<input type='text' name='rateitem" + lstitem + "' id='rateitem" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmountitem(" + lstitem + ")' class='lebarnormahk ratatengah'>";
	$('#trbadd' + lstitem).append("<td class='ratakiri'>" + rateitem + "</td>");

	var amountitem = "<input type='text' name='amountitemfrm" + lstitem + "' id='amountitemfrm" + lstitem + "' readonly value='0' onkeyup='hanyaAngka(this)' class='lebarnormahk ratatengah'><input type='hidden' name='amountitem" + lstitem + "' id='amountitem" + lstitem + "' readonly value='0' onkeyup='hanyaAngka(this)' class='lebarnormahk ratatengah'>";
	$('#trbadd' + lstitem).append("<td class='ratakiri'>" + amountitem + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowBahan(" + lstitem + ");' />";
	$('#trbadd' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trbadd' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jumldataitem').val(lstitem);


	showActivityitem("activityitem" + lstitem);
	new Spry.Widget.ValidationTextField("normamm" + lstitem);
	new Spry.Widget.ValidationTextField("qtyitem" + lstitem);
	new Spry.Widget.ValidationTextField("rateitem" + lstitem);
	
}

//hapus row bahan
function deleterowBahan(i) {
	Spry.Widget.Utils.destroyWidgets('itemcode' + i);
	Spry.Widget.Utils.destroyWidgets('uom' + i);
	Spry.Widget.Utils.destroyWidgets('activityitem' + i);
	Spry.Widget.Utils.destroyWidgets('tcitem' + i);
	Spry.Widget.Utils.destroyWidgets('normamm' + i);
	Spry.Widget.Utils.destroyWidgets('qtyitem' + i);
	Spry.Widget.Utils.destroyWidgets('rateitem' + i);
	Spry.Widget.Utils.destroyWidgets('amountitem' + i);
	Spry.Widget.Utils.destroyWidgets('amountitemfrm' + i);
	 $('#trbadd' + i).remove();
	// getjmlRowbhn();
}

//get activity item
function showActivityitem(selid) {
	$('#activityitem').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen='+14,
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
function getTCitem(i)
{
	var tahun = $('#tahun').val();
	var bulan = $('#bulan').val();
	var sectionname = $('#sectionname').val();
	var activityitem = $('#activityitem' + i).val();
	$('#tcitem' + i).empty();
	$('#tcitem' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok_rkb.asp?activity=' + activityitem + "&tcby=" + 1 + "&tahun=" + tahun + "&bulan=" + bulan + "&section=" + sectionname,
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

//tampilkan item
function showItem(i) {
	var activityitem = $('#activityitem' + i).val();
	var fcbatxt 	= $("#fcbatxt").val();
	$('#itemcode' + i).empty();
	$('#itemcode' + i).select2({
		ajax: {
			delay: 1000,
			url : "bkm_getactivity.asp?panen=" + 15 + "&activity=" + activityitem + "&fcba=" + fcbatxt,
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

//ambil data satuan
function getSatuan(i)
{
	var activityitem = $('#activityitem' + i).val();
	var itemcode = $('#itemcode' + i).val();
	if (itemcode != "") {
		$.post("bkm_employee.asp?itemcode=" + itemcode +"&satuan=" + 8 + "&jobcode=" + activityitem, function(data) {
			$('#uom' + i).empty();
			var json  		= JSON.parse(data);
			var uom 		= json.results[0].uomactivematerial
			var actmatnorm 	= json.results[0].actmatnorm
			$('#uom' + i).val(uom);
			$('#normamm' + i).val(actmatnorm);
						
		});
		
	}
}

//get amount ittem
function getAmountitem(i)
{
	var qtyitem 		= parseFloat($('#qtyitem' + i).val());
	var rateitem 		= parseFloat($('#rateitem' + i).val());
	if(qtyitem==""){
		qtyitem=0;
	}
	if(rateitem==""){
		rateitem=0;
	}
	var amountitem  	= qtyitem*rateitem;
	var amountitemfrm	= numformat(amountitem);
	$('#amountitemfrm' + i).val(amountitemfrm);
	$('#amountitem' + i).val(amountitem);
}

//get load activity bahan dan item
//getLoaditem
function getLoaditem() {
	
	var jumldataitem = $("#jumldataitem").val();
	for (var i = 0; i < jumldataitem; i++) {
		showActivityitem("activityitem" + i);
		//showItem("activityitem" + i);
	}

}
 </script>