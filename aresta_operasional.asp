<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno 
bkmno		= trim(request("bkmno"))
businessunit= trim(request("fcba"))

'response.write bkmno

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if


' get data header
dim blokhdr,sectionhdr,thnhdr,blnhdr,statushdr,hahdr,r1,own,sakun 
blokhdr = trim(request("fccode"))
sectionhdr = trim(request("division"))
thnhdr = trim(request("tahun"))
blnhdr = trim(request("bulan"))
statushdr = trim(request("status"))
hahdr = trim(request("ha"))
own = UCase(trim(request("own")))
r1 = trim(request("r1"))
sakun       = Request("sakun")




'count detailupah
' Dim SQL_Rows, NoOfRows,RSCount

' SQL_Rows = "SELECT Count (*) from BKMUPAHDTL " & _ 
' " where BKMNO='" & bkmno & "'"
' set RSCount = connStr.Execute(SQL_Rows)
' if not RSCount.eof then 
	' NoOfRows = RSCount(0).value
' end if
' RSCount.close

' 'count detailbahan
' Dim SQL_Rowsitem, NoOfRowsitem,RSCountitem

' SQL_Rowsitem = "SELECT Count (*) from BKMBAHANDTL where BKMNO='" & bkmno & "'"
' set RSCountitem = connStr.Execute(SQL_Rowsitem)
' if not RSCountitem.eof then 
	' NoOfRowsitem = RSCountitem(0).value
' end if
' RSCountitem.close


%>

<html>
	<head>
		<title>FBPlants - Aresta Operasional</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
	</head>

<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiript{
		position: relative;
		left: 35px;
	}
	.posisikiri{
		position: relative;
		left: -15px;
	}
	.posisikiritblatas{
		position: relative;
		left: -5px;
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
		width: 200px;
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



<body>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:center;">
		<strong class="posisikiript">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : ARESTA OPERASIONAL</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_operasional_prosesedit.asp">
	<input type="hidden" name="proses" id="proses" value="2" />
	<table align=center border="0" class='data posisikiri' width='90%' id='tableall'>			
		<tr>
			<td align="left" >Blok</td>
			<td align="left">
				<input type="text" name="blok" readonly id="blok" size="10" value="<%=blokhdr%>" title="BKM No">
			</td>


			<td align="left" >Section Name</td>
			<td align="left">
				<input type="text" name="sectionname" readonly id="sectionname" size="15" value="<%=sectionhdr%>">
				
			</td>
			<td align="left" >Status</td>
			<td align="left">
				<input type="text" name="status" readonly id="status" size="15" value="<%=statushdr%>-<%=own%>">
				
			</td>
			
		</tr>
		<tr>
			<td align="left" >Tahun</td>
			<td align="left">
				<input type="text" class="" readonly maxlength="15" name="tahun" readonly id="tahun" size="10" value="<%=thnhdr%>" title="Tanggal BKM">
			</td>
			
			<td align="left" >Bulan</td>
			<td align="left">
				<input type="text" name="bulan"  readonly id="bulan" size="15" value="<%=blnhdr%>">
			</td>
			
			<td align="left" >Ha</td>
			<td align="left">
				<input type="text" name="bulan"  readonly id="bulan" size="15" value="<%=hahdr%>">
			</td>
			
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='1100' id='tableall'>
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">LC&PERAWATAN</li>
						<li class="TabbedPanelsTab">PEMUPUKAN</li>
						<li class="TabbedPanelsTab">PANEN</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view perawatan-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="1100" class='data' id='table2'>
								<tr id="trheader2">
									<th>Tanggal</th>
									<th>Pekerjaan</th>
									<% if trim(r1) = "1" then%>
									<th>Blok</th>
									<%end if%>
									<th>HK</th>
									<th>Unit</th>
									<th>UOM</th>
									<th>OutPut(Ha)</th>
									<th>Amount</th>
									<th colspan="2">&nbsp;</th>
								</tr>

								<% 
								
								
								'get data detail perawatan
								dim sqldtlper,rsdtlper,no,ttlhk,ttlfisik,ttlamount,ttlha
								if trim(r1) = "2" then
								sqldtlper = "SELECT to_char(h.bkmdate,'DD/MM/YYYY') bkmdate,u.ACTCODE,u.ACTCODE || '-' ||  a.fcname as activityname,a.uom_unit, " & _   
								" sum(nvl(u.UNIT,0))UNIT, sum( nvl(u.OUTPUT,0)) OUTPUT,sum(nvl(u.MANDAYS,0))MANDAYS,sum(nvl(p.AMOUNT ,0))AMOUNT  " & _ 
								" FROM BKMUPAHDTL u  " & _ 
								" INNER JOIN BKMHDR h ON (h.bkmno=u.bkmno and h.fcba=u.fcba)  " & _ 
								" INNER JOIN (SELECT fcba,bkmno,employeecode,Sum(amount) amount FROM BKMPROSES GROUP BY fcba,bkmno,employeecode)  p on (p.bkmno=h.bkmno and p.fcba=h.fcba AND u.employeecode=p.employeecode) " & _
								" INNER JOIN jobmaster a ON (u.ACTCODE=a.FCCODE)  " & _ 
								" where h.postingflg='1' AND  TO_CHAR(h.BKMDATE,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and h.fcba='" & businessunit & "' " & _ 
								" AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
								" AND a.jobgroupcode IN (" & _
								" 		'1320100'," & _
								" 		'1320200'," & _
								" 		'1320300'," & _
								" 		'1320400'," & _
								" 		'1320500'," & _
								" 		'1320599'," & _
								" 		'1320600'," & _
								" 		'1320700'," & _
								" 		'1320900'," & _
								" 		'1321000'," & _
								" 		'1321100'," & _
								" 		'1321200'," & _
								" 		'1321300'," & _
								" 		'1321400'," & _
								" 		'6500100'," & _
								" 		'6500200'," & _
								" 		'6500210'," & _
								" 		'6500300'," & _
								" 		'6500400'," & _
								" 		'6500500'," & _
								" 		'6500600'," & _
								" 		'6500700'," & _
								" 		'6500800'," & _
								" 		'6500900'," & _
								" 		'6501000')" & _
								" AND attendance <> 'KP' and u.LOCATIONCODE='" & blokhdr & "' " & _ 
								" AND a.fccode like '" & sakun & "%'" & _
								" GROUP BY h.bkmdate,u.ACTCODE,u.ACTCODE || '-' || a.fcname,a.uom_unit " & _ 
								" HAVING sum(nvl(p.AMOUNT ,0)) <> 0 " & _
								" UNION ALL "  & _
								" SELECT to_char(u.wipdate,'DD/MM/YYYY') bkmdate,u.jobcode ACTCODE,u.jobcode || '-' ||  trim(a.fcname) || ' (WIP)' as activityname,a.uom_unit, " & _   
								" sum(nvl(u.qty_final,0))UNIT, sum( nvl(u.OUTPUT_final,0)) OUTPUT,sum(0)MANDAYS,sum(nvl(u.qty_final,0)*nvl(u.rate_final,0))AMOUNT  " & _ 
								" FROM workinprogressfinal_new u  " & _ 
								" INNER JOIN jobmaster a ON (u.jobcode=a.FCCODE)  " & _ 
								" where TO_CHAR(u.wipdate,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and u.chargeunit='" & businessunit & "' " & _ 
								" AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
								" AND a.jobgroupcode IN (" & _
								" 		'1320100'," & _
								" 		'1320200'," & _
								" 		'1320300'," & _
								" 		'1320400'," & _
								" 		'1320500'," & _
								" 		'1320599'," & _
								" 		'1320600'," & _
								" 		'1320700'," & _
								" 		'1320900'," & _
								" 		'1321000'," & _
								" 		'1321100'," & _
								" 		'1321200'," & _
								" 		'1321300'," & _
								" 		'1321400'," & _
								" 		'6500100'," & _
								" 		'6500200'," & _
								" 		'6500210'," & _
								" 		'6500300'," & _
								" 		'6500400'," & _
								" 		'6500500'," & _
								" 		'6500600'," & _
								" 		'6500700'," & _
								" 		'6500800'," & _
								" 		'6500900'," & _
								" 		'6501000')" & _
								" AND u.LOCATIONCODE='" & blokhdr & "' " & _ 
								" AND a.fccode like '" & sakun & "%'" & _
								" GROUP BY u.wipdate,u.jobcode,u.jobcode || '-' ||trim(a.fcname)||' (WIP)',u.LOCATIONCODE,a.uom_unit " & _ 
								" HAVING sum(nvl(u.qty_final ,0)*nvl(u.rate_final ,0)) <> 0 " & _
								" ORDER BY bkmdate,ACTCODE"
								else
								sqldtlper = "SELECT to_char(h.bkmdate,'DD/MM/YYYY') bkmdate,u.ACTCODE,u.ACTCODE || '-' ||  a.fcname as activityname,a.uom_unit, " & _   
								" sum(nvl(u.UNIT,0))UNIT, sum( nvl(u.OUTPUT,0)) OUTPUT,sum(nvl(u.MANDAYS,0))MANDAYS,sum(nvl(p.AMOUNT ,0))AMOUNT  " & _ 
								" ,u.LOCATIONCODE " & _
								" FROM BKMUPAHDTL u  " & _ 
								" INNER JOIN BKMHDR h ON (h.bkmno=u.bkmno and h.fcba=u.fcba)  " & _ 
								" INNER JOIN (SELECT fcba,bkmno,employeecode,Sum(amount) amount FROM BKMPROSES GROUP BY fcba,bkmno,employeecode)  p on (p.bkmno=h.bkmno and p.fcba=h.fcba AND u.employeecode=p.employeecode) " & _
								" INNER JOIN jobmaster a ON (u.ACTCODE=a.FCCODE)  " & _ 
								" INNER JOIN Field_hist_tab f on (u.LOCATIONCODE=f.fccode and u.chargeunit=f.fcba and f.year=to_number(TO_CHAR(h.BKMDATE,'YYYY')) and f.month=to_number(TO_CHAR(h.BKMDATE,'MM'))) " & _ 
								" where h.postingflg='1' AND  TO_CHAR(h.BKMDATE,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and u.chargeunit='" & businessunit & "' " & _ 
								" AND attendance <> 'KP' AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
								" AND upper(f.ownership)='" & own & "' " & _
								" AND a.jobgroupcode IN (" & _
								" 		'1320100'," & _
								" 		'1320200'," & _
								" 		'1320300'," & _
								" 		'1320400'," & _
								" 		'1320500'," & _
								" 		'1320599'," & _
								" 		'1320600'," & _
								" 		'1320700'," & _
								" 		'1320900'," & _
								" 		'1321000'," & _
								" 		'1321100'," & _
								" 		'1321200'," & _
								" 		'1321300'," & _
								" 		'1321400'," & _
								" 		'6500100'," & _
								" 		'6500200'," & _
								" 		'6500210'," & _
								" 		'6500300'," & _
								" 		'6500400'," & _
								" 		'6500500'," & _
								" 		'6500600'," & _
								" 		'6500700'," & _
								" 		'6500800'," & _
								" 		'6500900'," & _
								" 		'6501000')" & _
								" AND a.fccode like '" & sakun & "%'" & _
								" GROUP BY h.bkmdate,u.ACTCODE,u.ACTCODE || '-' || a.fcname,u.LOCATIONCODE,a.uom_unit " & _ 
								" HAVING sum(nvl(p.AMOUNT ,0)) <> 0 " & _
								" UNION ALL "  & _
								" SELECT to_char(u.wipdate,'DD/MM/YYYY') bkmdate,u.jobcode,u.jobcode || '-' || trim(a.fcname)||' (WIP)' as activityname,a.uom_unit, " & _   
								" sum(nvl(u.qty_final,0))UNIT, sum( nvl(u.OUTPUT_final,0)) OUTPUT,sum(0)MANDAYS,sum(nvl(u.qty_final,0)*nvl(u.rate_final,0))AMOUNT  " & _ 
								" , u.LOCATIONCODE " & _
								" FROM workinprogressfinal_new u  " & _ 
								" INNER JOIN jobmaster a ON (u.jobcode=a.FCCODE)  " & _ 
								" INNER JOIN Field_hist_tab f on (u.LOCATIONCODE=f.fccode and u.chargeunit=f.fcba and f.year=to_number(TO_CHAR(u.wipdate,'YYYY')) and f.month=to_number(TO_CHAR(u.wipdate,'MM'))) " & _
								" where TO_CHAR(u.wipdate,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and u.chargeunit='" & businessunit & "' " & _ 
								" AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
								" AND upper(f.ownership)='" & own & "' " & _
								" AND a.jobgroupcode IN (" & _
								" 		'1320100'," & _
								" 		'1320200'," & _
								" 		'1320300'," & _
								" 		'1320400'," & _
								" 		'1320500'," & _
								" 		'1320599'," & _
								" 		'1320600'," & _
								" 		'1320700'," & _
								" 		'1320900'," & _
								" 		'1321000'," & _
								" 		'1321100'," & _
								" 		'1321200'," & _
								" 		'1321300'," & _
								" 		'1321400'," & _
								" 		'6500100'," & _
								" 		'6500200'," & _
								" 		'6500210'," & _
								" 		'6500300'," & _
								" 		'6500400'," & _
								" 		'6500500'," & _
								" 		'6500600'," & _
								" 		'6500700'," & _
								" 		'6500800'," & _
								" 		'6500900'," & _
								" 		'6501000')" & _
								" AND a.fccode like '" & sakun & "%'" & _
								" GROUP BY u.wipdate,u.jobcode,u.jobcode || '-' ||trim(a.fcname)||' (WIP)',u.LOCATIONCODE,a.uom_unit " & _ 
								" HAVING sum(nvl(u.qty_final ,0)*nvl(u.rate_final ,0)) <> 0 " & _
								" ORDER BY ACTCODE,LOCATIONCODE,bkmdate"
								end if
								'response.write sqldtlper
								set rsdtlper = connStr.Execute(sqldtlper)
														
								no = 0
								ttlhk=0
								ttlfisik=0
								ttlha=0
								ttlamount=0
								While Not rsdtlper.EOF %>
								<tr id="tru<%=no%>">
									<td class="ratakiri">
										<%=rsdtlper("bkmdate")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlper("activityname")%>
									</td>
									<% if trim(r1) = "1" then %>
									<td class="ratakiri">
										<%=rsdtlper("LOCATIONCODE")%>
									</td>
									<%end if%>
									<td class="ratatengah">
										<%=rsdtlper("MANDAYS")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlper("UNIT")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlper("uom_unit")%>
									</td>
									<td class="ratatengah">
										<%=rsdtlper("OUTPUT")%>
									</td>
									<td class="ratakanan">
										<%=FORMATNUMBER(rsdtlper("AMOUNT"),0)%>
									</td>
									
									<td></td>
									
								</tr>
								<%
									ttlhk		= CDBL(ttlhk) + cdbl(rsdtlper("MANDAYS"))
									ttlfisik	= CDBL(ttlfisik) + cdbl(rsdtlper("UNIT"))
									ttlha	    = CDBL(ttlha) + cdbl(rsdtlper("OUTPUT"))
									ttlamount	= CDBL(ttlamount) + cdbl(rsdtlper("AMOUNT"))
									
									rsdtlper.MoveNext
									Wend
								%>
									<tr class="ratakanan">
									<% if trim(r1) = "1" then%>
										<td colspan="3" class="tebal">Total</td>
									<%else%>
										<td colspan="2" class="tebal">Total</td>
									<%end if%>
										<td class="tebal ratatengah"><%=FormatNumber(ttlhk,2)%></td>
										<td class="tebal ratatengah"><%=FormatNumber(ttlfisik,2)%></td>
										<td class="tebal ratatengah"></td>
										<td class="tebal ratatengah"><%=FormatNumber(ttlha,2)%></td>
										<td class="tebal ratakanan"><%=FormatNumber(ttlamount,0)%></td>
									
										
									</tr>
									<tr class="ratatengah">
										
									</tr>
							</table>
						</div>

						<!-- view pemupukan-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="1100" class='data' id='table1'>
									<tr id="trheader1">
										<th>Tanggal</th>
										<th>Pekerjaan</th>
										<th>UOM</th>
										<th>HK</th>
										<th>OutPut(Ha)</th>
										<th>Amount</th>
										<th colspan="2">&nbsp;</th>
									</tr>

									<%
									
										'get data detail pemupukan
										dim sqldtlpemu,rsdtlpemu,nopemu,ttlhkpemu,ttlfisikpu,ttlamountpu
										
										
								if trim(r1) = "2" then
								sqldtlpemu = "SELECT to_char(h.bkmdate,'DD/MM/YYYY') bkmdate,u.ACTCODE,u.ACTCODE || '-' ||  a.fcname as activityname,a.uom_unit, " & _   
										" sum(nvl(u.UNIT,0))UNIT, sum( nvl(u.OUTPUT,0)) OUTPUT,sum(nvl(u.MANDAYS,0))MANDAYS,sum(nvl(p.AMOUNT ,0))AMOUNT  " & _ 
										" FROM BKMUPAHDTL u  " & _ 
										" INNER JOIN BKMHDR h ON (h.bkmno=u.bkmno and h.fcba=u.fcba)  " & _ 
										" INNER JOIN (SELECT fcba,bkmno,employeecode,Sum(amount) amount FROM BKMPROSES GROUP BY fcba,bkmno,employeecode)  p on (p.bkmno=h.bkmno and p.fcba=h.fcba AND u.employeecode=p.employeecode) " & _
										" INNER JOIN jobmaster a ON (u.ACTCODE=a.FCCODE)  " & _ 
										" where h.postingflg='1' AND TO_CHAR(h.BKMDATE,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and h.fcba='" & businessunit & "' " & _ 
										" AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
										" AND a.jobgroupcode IN ('1320800','6500300') " & _
										" and attendance <> 'KP' and u.LOCATIONCODE='" & blokhdr & "' " & _ 
										" GROUP BY h.bkmdate,u.ACTCODE,u.ACTCODE || '-' || a.fcname,a.uom_unit " & _ 
										" HAVING sum(nvl(p.AMOUNT ,0)) <> 0 " & _
										" UNION ALL "  & _
										" SELECT to_char(u.wipdate,'DD/MM/YYYY') bkmdate,u.jobcode ACTCODE,u.jobcode || '-' ||  trim(a.fcname) || ' (WIP)' as activityname,a.uom_unit, " & _   
										" sum(nvl(u.qty_final,0))UNIT, sum( nvl(u.OUTPUT_final,0)) OUTPUT,sum(0)MANDAYS,sum(nvl(u.qty_final,0)*nvl(u.rate_final,0))AMOUNT  " & _ 
										" FROM workinprogressfinal_new u  " & _ 
										" INNER JOIN jobmaster a ON (u.jobcode=a.FCCODE)  " & _ 
										" where TO_CHAR(u.wipdate,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and u.chargeunit='" & businessunit & "' " & _ 
										" AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
										" AND a.jobgroupcode IN ('1320800','6500300') " & _
										" and u.LOCATIONCODE='" & blokhdr & "' " & _ 
										" GROUP BY u.wipdate,u.jobcode,u.jobcode || '-' ||trim(a.fcname)||' (WIP)',u.LOCATIONCODE,a.uom_unit " & _ 
										" HAVING sum(nvl(u.qty_final ,0)*nvl(u.rate_final ,0)) <> 0 " & _
										" ORDER BY bkmdate,ACTCODE"
								else
								sqldtlpemu = "SELECT to_char(h.bkmdate,'DD/MM/YYYY') bkmdate,u.ACTCODE,u.ACTCODE || '-' ||  a.fcname as activityname,a.uom_unit, " & _   
										" sum(nvl(u.UNIT,0))UNIT, sum( nvl(u.OUTPUT,0)) OUTPUT,sum(nvl(u.MANDAYS,0))MANDAYS,sum(nvl(p.AMOUNT ,0))AMOUNT  " & _ 
										" ,u.LOCATIONCODE " & _
										" FROM BKMUPAHDTL u  " & _ 
										" INNER JOIN BKMHDR h ON (h.bkmno=u.bkmno and h.fcba=u.fcba)  " & _ 
										" INNER JOIN (SELECT fcba,bkmno,employeecode,Sum(amount) amount FROM BKMPROSES GROUP BY fcba,bkmno,employeecode)  p on (p.bkmno=h.bkmno and p.fcba=h.fcba AND u.employeecode=p.employeecode) " & _
										" INNER JOIN jobmaster a ON (u.ACTCODE=a.FCCODE)  " & _ 
										" INNER JOIN Field_hist_tab f on (u.LOCATIONCODE=f.fccode and u.chargeunit=f.fcba and f.year=to_number(TO_CHAR(h.BKMDATE,'YYYY')) and f.month=to_number(TO_CHAR(h.BKMDATE,'MM'))) " & _ 
										" where h.postingflg='1' AND TO_CHAR(h.BKMDATE,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and u.chargeunit='" & businessunit & "' " & _ 
										" AND attendance <> 'KP' AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
										" AND upper(f.ownership)='" & own & "' " & _
										" AND a.jobgroupcode IN ('1320800','6500300') " & _
										" GROUP BY h.bkmdate,u.ACTCODE,u.ACTCODE || '-' || a.fcname,u.LOCATIONCODE,a.uom_unit " & _ 
										" HAVING sum(nvl(p.AMOUNT ,0)) <> 0 " & _
										" UNION ALL "  & _
										" SELECT to_char(u.wipdate,'DD/MM/YYYY') bkmdate,u.jobcode,u.jobcode || '-' || trim(a.fcname)||' (WIP)' as activityname,a.uom_unit, " & _   
										" sum(nvl(u.qty_final,0))UNIT, sum( nvl(u.OUTPUT_final,0)) OUTPUT,sum(0)MANDAYS,sum(nvl(u.qty_final,0)*nvl(u.rate_final,0))AMOUNT  " & _ 
										" , u.LOCATIONCODE " & _
										" FROM workinprogressfinal_new u  " & _ 
										" INNER JOIN jobmaster a ON (u.jobcode=a.FCCODE)  " & _ 
										" INNER JOIN Field_hist_tab f on (u.LOCATIONCODE=f.fccode and u.chargeunit=f.fcba and f.year=to_number(TO_CHAR(u.wipdate,'YYYY')) and f.month=to_number(TO_CHAR(u.wipdate,'MM'))) " & _
										" where TO_CHAR(u.wipdate,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and u.chargeunit='" & businessunit & "' " & _ 
										" AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
										" AND upper(f.ownership)='" & own & "' " & _
										" AND a.jobgroupcode IN ('1320800','6500300') " & _
										" GROUP BY u.wipdate,u.jobcode,u.jobcode || '-' ||trim(a.fcname)||' (WIP)',u.LOCATIONCODE,a.uom_unit " & _ 
										" HAVING sum(nvl(u.qty_final ,0)*nvl(u.rate_final ,0)) <> 0 " & _
										" ORDER BY ACTCODE,LOCATIONCODE,bkmdate"
								end if
								
								'response.write sqldtlpemu
										set rsdtlpemu = connStr.Execute(sqldtlpemu)
																
										nopemu = 0
										ttlhkpemu=0
										ttlfisikpu=0
										ttlamountpu=0
										While Not rsdtlpemu.EOF %>
										<tr id='trb<%=nopemu%>'>
										    <td class="ratakiri">
												<%=rsdtlpemu("bkmdate")%>
											</td>
											<td class="ratakiri">
												<%=rsdtlpemu("activityname")%>
											</td>
											<td class="ratatengah">
												<%=rsdtlpemu("uom_unit")%>
											</td>
											<td class="ratatengah">
												<%=rsdtlpemu("MANDAYS")%>
											</td>
											<td class="ratatengah">
												<%=rsdtlpemu("OUTPUT")%>
											</td>
											<td class="ratakanan">
												<%=FORMATNUMBER(rsdtlpemu("AMOUNT"),0)%>
											</td>
											<td></td>
										</tr>

											
										
									<%
										ttlhkpemu		= CDBL(ttlhkpemu) + cdbl(rsdtlpemu("MANDAYS"))
										ttlfisikpu		= CDBL(ttlfisikpu) + cdbl(rsdtlpemu("OUTPUT"))
										ttlamountpu	    = CDBL(ttlamountpu) + cdbl(rsdtlpemu("AMOUNT"))
										nopemu = nopemu + 1
										rsdtlpemu.MoveNext
										Wend
									
									%>

									<tr class="ratakanan">
										<td colspan="3" class="tebal">Total</td>
										<td class="tebal ratatengah"><%=FormatNumber(ttlhkpemu,2)%></td>
										<td class="tebal ratatengah"><%=FormatNumber(ttlfisikpu,2)%></td>
										<td class="tebal ratakanan"><%=FormatNumber(ttlamountpu,0)%></td>
									
										
									</tr>
									
									
								</table>
							</div>
							
						</div>

						<!-- view panen-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="1100" class='data' id='table1'>
									<tr id="trheader1">
	                                    <th>Tanggal</th>
										<th>Pekerjaan</th>
										<th>UOM</th>
										<th>HK</th>
										<th>OutPut(Ha)</th>
										<th>Jjg</th>
										<th>Kg</th>
										<th>Amount</th>
										<th colspan="2">&nbsp;</th>
									</tr>

									<%
									
										'get data detail panen
										dim sqldtlpan,rsdtlpan,nopan,ttlhkpan,ttlfisikpan,ttlamountpan,ttlfisikha,ttlfisikkg
										
									if trim(r1) = "2" then
								sqldtlpan = "SELECT to_char(h.bkmdate,'DD/MM/YYYY') bkmdate,u.ACTCODE,u.ACTCODE || '-' ||  a.fcname as activityname,a.uom_unit, " & _   
										" sum(nvl(u.UNIT,0))UNIT, sum( nvl(u.OUTPUT,0)) OUTPUT,sum(nvl(u.MANDAYS,0))MANDAYS, sum(nvl(u.bunch_weight,0))bunch_weight, sum(nvl(p.AMOUNT ,0))AMOUNT  " & _ 
										" FROM BKMUPAHDTL u  " & _ 
										" INNER JOIN BKMHDR h ON (h.bkmno=u.bkmno and h.fcba=u.fcba)  " & _ 
										" INNER JOIN (SELECT fcba,bkmno,employeecode,Sum(amount) amount FROM BKMPROSES GROUP BY fcba,bkmno,employeecode)  p on (p.bkmno=h.bkmno and p.fcba=h.fcba AND u.employeecode=p.employeecode) " & _
										" INNER JOIN jobmaster a ON (u.ACTCODE=a.FCCODE)  " & _ 
										" INNER JOIN Field_hist_tab f on (u.LOCATIONCODE=f.fccode and u.chargeunit=f.fcba and f.year=to_number(TO_CHAR(h.BKMDATE,'YYYY')) and f.month=to_number(TO_CHAR(h.BKMDATE,'MM'))  AND a.job_forfieldstatus=f.dx_block_status) " & _ 
										" where h.postingflg='1' AND  TO_CHAR(h.BKMDATE,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and h.fcba='" & businessunit & "' " & _ 
										" AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
										" AND a.jobgroupcode IN ('1338900','6502100','6502200') " & _
										" and attendance <> 'KP' and u.LOCATIONCODE='" & blokhdr & "' " & _ 
										" GROUP BY h.bkmdate,u.ACTCODE,u.ACTCODE || '-' || a.fcname,a.uom_unit " & _ 
										" HAVING sum(nvl(p.AMOUNT ,0)) <> 0 " & _
										" UNION ALL "  & _
										" SELECT to_char(u.wipdate,'DD/MM/YYYY') bkmdate,u.jobcode ACTCODE,u.jobcode || '-' ||  trim(a.fcname) || ' (WIP)' as activityname,a.uom_unit, " & _   
										" sum(nvl(u.qty_final,0))UNIT, sum( nvl(u.OUTPUT_final,0)) OUTPUT,sum(0)MANDAYS,0 bunch_weight,sum(nvl(u.qty_final,0)*nvl(u.rate_final,0))AMOUNT  " & _ 
										" FROM workinprogressfinal_new u  " & _ 
										" INNER JOIN jobmaster a ON (u.jobcode=a.FCCODE)  " & _ 
										" INNER JOIN Field_hist_tab f on (u.LOCATIONCODE=f.fccode and u.chargeunit=f.fcba and f.year=to_number(TO_CHAR(u.wipdate,'YYYY')) and f.month=to_number(TO_CHAR(u.wipdate,'MM'))  AND a.job_forfieldstatus=f.dx_block_status) " & _ 
										" where TO_CHAR(u.wipdate,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and u.chargeunit='" & businessunit & "' " & _ 
										" AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
										" AND a.jobgroupcode IN ('1338900','6502100','6502200') " & _
										" and u.LOCATIONCODE='" & blokhdr & "' " & _ 
										" GROUP BY u.wipdate,u.jobcode,u.jobcode || '-' ||trim(a.fcname)||' (WIP)',u.LOCATIONCODE,a.uom_unit " & _ 
										" HAVING sum(nvl(u.qty_final ,0)*nvl(u.rate_final ,0)) <> 0 " & _
										" ORDER BY bkmdate,ACTCODE"
								else
								sqldtlpan = "SELECT to_char(h.bkmdate,'DD/MM/YYYY') bkmdate,u.ACTCODE,u.ACTCODE || '-' ||  a.fcname as activityname,a.uom_unit, " & _   
										" sum(nvl(u.UNIT,0))UNIT, sum( nvl(u.OUTPUT,0)) OUTPUT,sum(nvl(u.MANDAYS,0))MANDAYS,sum(nvl(u.bunch_weight,0))bunch_weight,sum(nvl(p.AMOUNT ,0))AMOUNT  " & _ 
										" ,u.LOCATIONCODE " & _
										" FROM BKMUPAHDTL u  " & _ 
										" INNER JOIN BKMHDR h ON (h.bkmno=u.bkmno and h.fcba=u.fcba)  " & _ 
										" INNER JOIN (SELECT fcba,bkmno,employeecode,Sum(amount) amount FROM BKMPROSES GROUP BY fcba,bkmno,employeecode)  p on (p.bkmno=h.bkmno and p.fcba=h.fcba AND u.employeecode=p.employeecode) " & _
										" INNER JOIN jobmaster a ON (u.ACTCODE=a.FCCODE)  " & _ 
										" INNER JOIN Field_hist_tab f on (u.LOCATIONCODE=f.fccode and u.chargeunit=f.fcba and f.year=to_number(TO_CHAR(h.BKMDATE,'YYYY')) and f.month=to_number(TO_CHAR(h.BKMDATE,'MM'))  AND a.job_forfieldstatus=f.dx_block_status) " & _ 
										" where h.postingflg='1' AND TO_CHAR(h.BKMDATE,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and u.chargeunit='" & businessunit & "' " & _ 
										" AND attendance <> 'KP' AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
										" AND upper(f.ownership)='" & own & "' " & _
										" AND a.jobgroupcode IN ('1338900','6502100','6502200') " & _
										" GROUP BY h.bkmdate,u.ACTCODE,u.ACTCODE || '-' || a.fcname,u.LOCATIONCODE,a.uom_unit " & _ 
										" HAVING sum(nvl(p.AMOUNT ,0)) <> 0 " & _
										" UNION ALL "  & _
										" SELECT to_char(u.wipdate,'DD/MM/YYYY') bkmdate,u.jobcode,u.jobcode || '-' || trim(a.fcname)||' (WIP)' as activityname,a.uom_unit, " & _   
										" sum(nvl(u.qty_final,0))UNIT, sum( nvl(u.OUTPUT_final,0)) OUTPUT,sum(0)MANDAYS,0 bunch_weight, sum(nvl(u.qty_final,0)*nvl(u.rate_final,0))AMOUNT  " & _ 
										" , u.LOCATIONCODE " & _
										" FROM workinprogressfinal_new u  " & _ 
										" INNER JOIN jobmaster a ON (u.jobcode=a.FCCODE)  " & _ 
										" INNER JOIN Field_hist_tab f on (u.LOCATIONCODE=f.fccode and u.chargeunit=f.fcba and f.year=to_number(TO_CHAR(u.wipdate,'YYYY')) and f.month=to_number(TO_CHAR(u.wipdate,'MM')) AND a.job_forfieldstatus=f.dx_block_status) " & _
										" where TO_CHAR(u.wipdate,'YYYYMM')='" & thnhdr & "'||LPad('" & blnhdr & "',2,0) and u.chargeunit='" & businessunit & "' " & _ 
										" AND a.job_forfieldstatus = ('" & statushdr & "') " & _ 
										" AND upper(f.ownership)='" & own & "' " & _
										" AND a.jobgroupcode IN ('1338900','6502100','6502200') " & _
										" GROUP BY u.wipdate,u.jobcode,u.jobcode || '-' ||trim(a.fcname)||' (WIP)',u.LOCATIONCODE,a.uom_unit " & _ 
										" HAVING sum(nvl(u.qty_final ,0)*nvl(u.rate_final ,0)) <> 0 " & _
										" ORDER BY ACTCODE,LOCATIONCODE,bkmdate"
								end if
								'response.write sqldtlpan
										set rsdtlpan = connStr.Execute(sqldtlpan)
																
										nopan = 0
										ttlhkpan=0
										ttlfisikha=0
										ttlfisikpan=0
										ttlfisikkg=0
										ttlamountpan=0
										While Not rsdtlpan.EOF %>
										<tr id='trp<%=nopan%>'>
											<td class="ratakiri">
												<%=rsdtlpan("bkmdate")%>
											</td>
											<td class="ratakiri">
												<%=rsdtlpan("activityname")%>
											</td>
											<td class="ratatengah">
												<%=rsdtlpan("uom_unit")%>
											</td>
											<td class="ratatengah">
												<%=rsdtlpan("MANDAYS")%>
											</td>
											<td class="ratatengah">
												<%=rsdtlpan("OUTPUT")%>
											</td>
											<td class="ratatengah">
												<%=rsdtlpan("UNIT")%>
											</td>
											<td class="ratatengah">
												<%=rsdtlpan("bunch_weight")%>
											</td>
											<td class="ratakanan">
												<%=FORMATNUMBER(rsdtlpan("AMOUNT"),0)%>
											</td>
											<td></td>
										</tr>

											
										
									<%
										ttlhkpan		= CDBL(ttlhkpan) + cdbl(rsdtlpan("MANDAYS"))
										ttlfisikha      = CDBL(ttlfisikha) + cdbl(rsdtlpan("OUTPUT"))
										ttlfisikpan		= CDBL(ttlfisikpan) + cdbl(rsdtlpan("UNIT"))
										ttlfisikkg		= CDBL(ttlfisikkg) + cdbl(rsdtlpan("bunch_weight"))
										
										ttlamountpan	= CDBL(ttlamountpan) + cdbl(rsdtlpan("AMOUNT"))

										nopan = nopan + 1
										rsdtlpan.MoveNext
										Wend
									
									%>

									<tr class="ratakanan">
										<td colspan="3" class="tebal">Total</td>
										<td class="tebal ratatengah"><%=FormatNumber(ttlhkpan,0)%></td>
										<td class="tebal ratatengah"><%=FormatNumber(ttlfisikha,0)%></td>
										<td class="tebal ratatengah"><%=FormatNumber(ttlfisikpan,0)%></td>
										<td class="tebal ratatengah"><%=FormatNumber(ttlfisikkg,0)%></td>
										<td class="tebal ratakanan"><%=FormatNumber(ttlamountpan,0)%></td>
									
										
									</tr>
									
								</table>
							</div>
							
						</div>

						
					</div>
				</div>
			</td>
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



</script>
