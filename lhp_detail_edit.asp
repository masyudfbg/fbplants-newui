<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,lhpno,tglprod 
lhpno= trim(request("lhpno"))
businessunit= trim(session("fcba"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if


' get data header
dim sqlheader,rsHeader
sqlheader = "select to_char(proddate,'dd/mm/yyyy')proddate,lhpno from productiontrn  where lhpno='" & lhpno & "'"
set rsHeader = connStr.Execute(sqlheader)
'response.write rsHeader
if not rsHeader.eof then 
	tglprod = rsHeader("proddate")	
end if


%>

<html>
	<head>
		<title>FBPlants - LHP Operasional</title>
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
	.posisikiri{
		position: relative;
		left: 20px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarabsen{
		width: 60px;
	}
	.lebaremployee{
		width: 200px;
	}
	.lebaractivity{
		width: 200px;
	}
	.lebaracllowedcode{
		width: 210px;
	}
	.lebarupah{
		width: 80px;
	}
	.lebaritembarang{
		width: 400px;
	}
	.lebaractbarang{
		width: 300px;
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

	.lebarvehicle{
		width: 200px;
	}
	.lebartime{
		width: 90px;
	}

	.lebarspkno{
		width: 250px;
	}
	.lebarqty{
		width: 30px;
	}
	
</style>

<script>

</script>

<body>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : LHP OPERASIONAL</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="lhp_detail_edit.asp">
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >LHP No</td>
			<td align="left">
				<input type="text" name="lhpno" readonly id="lhpno" size="30" value="<%=lhpno%>" title="LHP No">
			</td>
		</tr>
		<tr>
			<td align="left" >Tanggal</td>
			<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="tgllhp" readonly id="tgllhp" size="15" value="<%=tglprod%>" title="Tanggal LHP">
			</td>
			
			
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='900' id="product-table" class="xthf-site">
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">TIKET TBS INTERNAL</li>
						<li class="TabbedPanelsTab">TIKET TBS EXTERNAL</li>
						<li class="TabbedPanelsTab">TIKET CPO</li>
						<li class="TabbedPanelsTab">TIKET KERNEL</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view TIKET TBS INTERNAL-->
						<div class="TabbedPanelsContent">
							<table align=center border="1" width="1000" class='data xthf-site'>
								<thead>
								<tr id="trheader2">
									
									<th>Estate/Divisi</th>
									<th>No Tiket</th>
									<th>SPBNo</th>
									<!--<th>TimeIn</th>
										<th>TimeOut</th>-->
									<th>Truct</th>
									<th>Bruto</th>
									<th>Tara</th>
									<th>Sortasi</th>
									<th>Netto</th>
									
								</tr>
								</thead>

								<% 
								
								
								'get data detail upah
								dim sqldtltbs,rsdtltbs,no,ttlbruto,ttltara,ttlsortasi,ttlnetto
								sqldtltbs = "select a.areaname,hd.docketdate,hd.docketid,hd.spbno,hd.timein,hd.timeout,hd.noreg,dt.weightin,dt.weightout,dt.sortasiwgt,dt.ffbweight,hd.fcba " & _
								" from LHPTIKETHDR hd  " & _
								" inner join LHPFFBWBTRN dt on hd.docketid=dt.docketid  " & _
								" inner join LHPAREAMST a on dt.areaid=a.areaid AND a.fcba=dt.fcba " & _
								" where hd.DOCKETDATE=to_date('" & tglprod & "','dd/mm/yyyy') " & _
								" and hd.fcba=dt.fcba and hd.fcba='" & businessunit & "' AND a.UNIT !='EXT' " & _
								" order by hd.docketdate,a.areaname asc"
								set rsdtltbs = connStr.Execute(sqldtltbs)
								

																
								no = 0
								ttlbruto = 0
								ttltara = 0
								ttlsortasi = 0
								ttlnetto = 0
								While Not rsdtltbs.EOF %>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri lebarabsen">
										<span><%=rsdtltbs("areaname")%></span>
									</td>
									<td class="ratakiri lebarabsen">
										<span><%=rsdtltbs("docketid")%></span>
									</td>
									<td class="ratakiri lebarabsen">
										<span><%=rsdtltbs("spbno")%></span>
									</td>
									<!--<td class="ratatengah lebarabsen">
										<%=rsdtltbs("timein")%>
									</td>
									<td class="ratatengah lebarabsen">
										<%=rsdtltbs("timeout")%>
									</td>-->
									<td class="ratatengah lebarabsen">
										<%=rsdtltbs("noreg")%>
									</td>
									<td class="ratatengah lebarqty">
										<%=FormatNumber(rsdtltbs("weightin"),0)%>
									</td>
									<td class=" ratatengah lebarqty">
										<%=FormatNumber(rsdtltbs("weightout"),0)%>
									</td>
									<td class="ratatengah lebarqty">
										<%=FormatNumber(rsdtltbs("sortasiwgt"),0)%>
									</td>

									<td class="ratatengah lebarqty">
										<%=FormatNumber(rsdtltbs("ffbweight"),0)%>
									</td>
									
									
									
								</tr>
								<%
									no = no +1
									ttlbruto = cdbl(ttlbruto) + cdbl(rsdtltbs("weightin"))
									ttltara = cdbl(ttltara) + cdbl(rsdtltbs("weightout"))
									ttlsortasi = cdbl(ttlsortasi) + cdbl(rsdtltbs("sortasiwgt"))
									ttlnetto = cdbl(ttlnetto) + cdbl(rsdtltbs("ffbweight"))
									rsdtltbs.MoveNext
									Wend
								%>
									<tr class="">
										<td colspan="4" class="tebal ratakanan">Total</td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlbruto,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttltara,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlsortasi,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlnetto,0)%></span></td>
									</tr>
							</table>
						</div>

						<!-- view TIKET TBS EXTERNAL-->
						<div class="TabbedPanelsContent">
							<table align=center border="1" width="1000" class='data xthf-site'>
								<thead>
								<tr id="trheader2">
									
									<th>Estate/Supplier</th>
									<th>No Tiket</th>
									<th>SPBNo</th>
									<!--<th>TimeIn</th>
										<th>TimeOut</th>-->
									<th>Truct</th>
									<th>Bruto</th>
									<th>Tara</th>
									<th>Sortasi</th>
									<th>Netto</th>
									
								</tr>
								</thead>

								<% 
								
								
								'get data detail upah
								'dim sqldtltbs,rsdtltbs,no,ttlbruto,ttltara,ttlsortasi,ttlnetto
								sqldtltbs = "select a.areaname,hd.docketdate,hd.docketid,hd.spbno,hd.timein,hd.timeout,hd.noreg,dt.weightin,dt.weightout,dt.sortasiwgt,dt.ffbweight,hd.fcba " & _
								" from LHPTIKETHDR hd  " & _
								" inner join LHPFFBWBTRN dt on hd.docketid=dt.docketid  " & _
								" inner join LHPAREAMST a on dt.areaid=a.areaid AND a.fcba=dt.fcba " & _
								" where hd.DOCKETDATE=to_date('" & tglprod & "','dd/mm/yyyy') " & _
								" and hd.fcba=dt.fcba and hd.fcba='" & businessunit & "' AND a.UNIT='EXT' " & _
								" order by hd.docketdate,a.areaname asc"
								set rsdtltbs = connStr.Execute(sqldtltbs)
								

																
								no = 0
								ttlbruto = 0
								ttltara = 0
								ttlsortasi = 0
								ttlnetto = 0
								While Not rsdtltbs.EOF %>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri lebarabsen">
										<span><%=rsdtltbs("areaname")%></span>
									</td>
									<td class="ratakiri lebarabsen">
										<span><%=rsdtltbs("docketid")%></span>
									</td>
									<td class="ratakiri lebarabsen">
										<span><%=rsdtltbs("spbno")%></span>
									</td>
									<!--<td class="ratatengah lebarabsen">
										<%=rsdtltbs("timein")%>
									</td>
									<td class="ratatengah lebarabsen">
										<%=rsdtltbs("timeout")%>
									</td>-->
									<td class="ratatengah lebarabsen">
										<%=rsdtltbs("noreg")%>
									</td>
									<td class="ratatengah lebarqty">
										<%=FormatNumber(rsdtltbs("weightin"),0)%>
									</td>
									<td class=" ratatengah lebarqty">
										<%=FormatNumber(rsdtltbs("weightout"),0)%>
									</td>
									<td class="ratatengah lebarqty">
										<%=FormatNumber(rsdtltbs("sortasiwgt"),0)%>
									</td>

									<td class="ratatengah lebarqty">
										<%=FormatNumber(rsdtltbs("ffbweight"),0)%>
									</td>
									
									
									
								</tr>
								<%
									no = no +1
									ttlbruto = cdbl(ttlbruto) + cdbl(rsdtltbs("weightin"))
									ttltara = cdbl(ttltara) + cdbl(rsdtltbs("weightout"))
									ttlsortasi = cdbl(ttlsortasi) + cdbl(rsdtltbs("sortasiwgt"))
									ttlnetto = cdbl(ttlnetto) + cdbl(rsdtltbs("ffbweight"))
									rsdtltbs.MoveNext
									Wend
								%>
									<tr class="">
										<td colspan="4" class="tebal ratakanan">Total</td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlbruto,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttltara,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlsortasi,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlnetto,0)%></span></td>
									</tr>
							</table>
						</div>

						<!-- view tiket cpo-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="1" width="1050" class='data' id='table1'>
									<tr id="trheader1">
										<th>DO</th>
										<th>Kontrak</th>
										<th>No Tiket</th>
										
										<th>Customer</th>
										<th>Contractor</th>
										<th>Truct</th>
										<th>Bruto</th>
										<th>Tara</th>
										<th>Netto</th>
										
									</tr>

									<%
										dim sqlcpo,rscpo,ttlbrutocpo,ttltaracpo,ttlnettocpo,nocpo
										sqlcpo = "select h.docketdate,dom.doorderid,cr.ContractNo,h.docketid,h.PRODUCTID ,h.TimeIn,h.TimeOut,cst.CustomerName,h.NOREG ,c.ContractorName, " & _
										" h.WEIGHTOUT, h.WEIGHTIN,h.WEIGHTOUT-h.WEIGHTIN as Netto " & _
										" from LHPTIKETHDR h  " & _
										" INNER JOIN LHPTIKETCPO d ON h.docketid=d.docketid AND h.fcba=d.fcba " & _
										" INNER JOIN LHPDOMST dom ON dom.doid=d.doid and dom.fcba=d.fcba " & _
										
										" INNER JOIN LHPCONTRACTMST cr ON dom.ContractID=cr.ContractID and h.fcba=cr.fcba " & _
										" INNER JOIN LHPCUSTOMERMST cst ON cst.CUSTOMERID=cr.CustomerID and cst.fcba=cr.fcba " & _
										" INNER JOIN LHPCONTRACTORMST c ON d.ContractorID = c.ContractorID and d.fcba=c.fcba " & _
										" where h.docketdate= to_Date('" & tglprod & "','dd/mm/yyyy') and h.fcba='" & businessunit & "' " & _
										" AND h.PRODUCTID='1' ORDER BY h.DocketDate, h.DocketID"
										set rscpo = connStr.Execute(sqlcpo)
									''	Response.write(sqlcpo)

																
								nocpo = 0
								ttlbrutocpo = 0
								ttltaracpo = 0
								ttlnettocpo = 0
								While Not rscpo.EOF %>
								<tr id="trc<%=nocpo%>">
									
									<td class="ratakiri lebarabsen">
										<span><%=rscpo("doorderid")%></span>
									</td>
									<td class="ratakiri lebarabsen">
										<span><%=rscpo("ContractNo")%></span>
									</td>
									<td class="ratakiri lebarabsen">
										<span><%=rscpo("docketid")%></span>
									</td>
									
									<td class="ratatengah lebarabsen">
										<%=rscpo("CustomerName")%>
									</td>
									<td class="ratatengah lebarabsen">
										<%=rscpo("ContractorName")%>
									</td>
									<td class="ratatengah lebarabsen">
										<%=rscpo("noreg")%>
									</td>
									<td class=" ratatengah lebarqty">
										<%=FormatNumber(rscpo("weightout"),0)%>
									</td>
									<td class="ratatengah lebarqty">
										<%=FormatNumber(rscpo("weightin"),0)%>
									</td>
									
									<td class="ratatengah lebarqty">
										<%=FormatNumber(rscpo("Netto"),0)%>
									</td>
									
									
									
								</tr>
								<%
									nocpo = nocpo +1
									ttlbrutocpo = cdbl(ttlbrutocpo) + cdbl(rscpo("weightout"))
									ttltaracpo = cdbl(ttltaracpo) + cdbl(rscpo("weightin"))
									ttlnettocpo = cdbl(ttlnettocpo) + cdbl(rscpo("Netto"))
									rscpo.MoveNext
									Wend
								%>
									<tr class="">
										<td colspan="6" class="tebal ratakanan">Total</td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlbrutocpo,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttltaracpo,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlnettocpo,0)%></span></td>
										
									</tr>
									
								</table>
							</div>
							
						</div>

						<!-- view lhp kernel-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="1" width="1050" class='data' id='table1'>
									<tr id="trheader3">
										<th>DO</th>
										<th>Kontrak</th>
										<th>No Tiket</th>
										
										<th>Customer</th>
										<th>Contractor</th>
										<th>Truct</th>
										<th>Bruto</th>
										<th>Tara</th>
										<th>Netto</th>
									</tr>

									<%
										dim sqlpk,rspk,ttlbrutopk,ttltarapk,ttlnettopk,nopk
										sqlpk = "select h.docketdate,dom.doorderid,cr.ContractNo,h.docketid,h.PRODUCTID ,h.TimeIn,h.TimeOut,cst.CustomerName,h.NOREG ,c.ContractorName, " & _
										" h.WEIGHTOUT, h.WEIGHTIN,h.WEIGHTOUT-h.WEIGHTIN as Netto " & _
										" from LHPTIKETHDR h  " & _
										" INNER JOIN LHPTIKETKERNEL d ON h.docketid=d.docketid AND h.fcba=d.fcba " & _
										" INNER JOIN LHPDOMST dom ON dom.doid=d.doid and dom.fcba=d.fcba " & _
										" INNER JOIN LHPCONTRACTMST cr ON dom.ContractID=cr.ContractID and h.fcba=cr.fcba " & _
										" INNER JOIN LHPCUSTOMERMST cst ON cst.CUSTOMERID=cr.CustomerID and cst.fcba=cr.fcba " & _
										" INNER JOIN LHPCONTRACTORMST c ON d.ContractorID = c.ContractorID and d.fcba=c.fcba " & _
										" where h.docketdate = to_Date('" & tglprod & "','dd/mm/yyyy') and h.fcba='" & businessunit & "' " & _
										" AND h.PRODUCTID='3' ORDER BY h.DocketDate, h.DocketID"
										set rspk = connStr.Execute(sqlpk)
									''	Response.write(sqlpk)

																
								nopk = 0
								ttlbrutopk = 0
								ttltarapk = 0
								ttlnettopk = 0
								While Not rspk.EOF %>
								<tr id="trk<%=nopk%>">
									
									<td class="ratakiri lebarabsen">
										<span><%=rspk("doorderid")%></span>
									</td>
									<td class="ratakiri lebarabsen">
										<span><%=rspk("ContractNo")%></span>
									</td>
									<td class="ratakiri lebarabsen">
										<span><%=rspk("docketid")%></span>
									</td>
								
									<td class="ratatengah lebarabsen">
										<%=rspk("CustomerName")%>
									</td>
									<td class="ratatengah lebarabsen">
										<%=rspk("ContractorName")%>
									</td>
									<td class="ratatengah lebarabsen">
										<%=rspk("noreg")%>
									</td>
									<td class=" ratatengah lebarqty">
										<%=FormatNumber(rspk("weightout"),0)%>
									</td>
									<td class="ratatengah lebarqty">
										<%=FormatNumber(rspk("weightin"),0)%>
									</td>
									
									<td class="ratatengah lebarqty">
										<%=FormatNumber(rspk("Netto"),0)%>
									</td>
									
									
									
								</tr>
								<%
									nopk = nopk +1
									ttlbrutopk = cdbl(ttlbrutopk) + cdbl(rspk("weightout"))
									ttltarapk = cdbl(ttltarapk) + cdbl(rspk("weightin"))
									ttlnettopk = cdbl(ttlnettopk) + cdbl(rspk("Netto"))
									rspk.MoveNext
									Wend
								%>
									<tr class="">
										<td colspan="6" class="tebal ratakanan">Total</td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlbrutopk,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttltarapk,0)%></span></td>
										<td class="tebal ratatengah"><span><%=FormatNumber(ttlnettopk,0)%></span></td>
										
									</tr>

									
								</table>
							</div>
							
						</div>
					</div>
				</div>
			</td>
		</tr>
		<tr class="ratatengah">
			<td  colspan="11"><input type=button name="btncancel" value="KEMBALI" onclick=window.close()></td>
			
		</tr>
	</table>

</form>
</div>
</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript" LANGUAGE="VBScript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

$(function() {	
	$("#tglbkm").datepicker();
		
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
