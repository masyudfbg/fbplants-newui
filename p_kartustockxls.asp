<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Invoice</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
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

<%
	Const adOpenForwardOnly = 0
	Const adLockReadOnly = 1
	Const adCmdTableDirect = &H0200
	Const adUseClient = 3
	cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
	cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
	cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

	dim tglawal,tglakhir,businessunit,productstock

	businessunit 	= trim(request("fcba"))
	productstock 	= trim(request("productstock"))
	tglawal 		= trim(request("tglawal"))
	tglakhir 		= trim(request("tglakhir"))
	wrhouse 		= trim(Request("wrhouse"))

	Dim formattedDate
	formattedDate = Right("0" & Month(tglawal), 2) & "/" & Right("0" & Day(tglawal), 2) & "/" & Year(tglawal)

	if trim(wrhouse) = "All" then
		addfilterSelect = " AND STORE IN (select fccode from store where fcba = '"&businessunit&"' and activation = 'Y')"
	else
		addfilterSelect = " AND STORE IN ('"&wrhouse&"', '-')"
    end if

	Dim sqlviewq, rsview
	if Day(formattedDate) = "1" then
		sqlView = 	" select  " & _ 
					" tgltrx, type, itemcode, origin, dest, inqty, outqty, fcba, " & _ 
					" store, note, fcentry, stockawal, stockakhir, trxcode, nominal, companyname" & _
					" from stock_history" & _
					" where fcba = '"& businessunit &"' and itemcode = '"& productstock &"' " & _
					" AND TO_DATE(tgltrx,'MM/DD/YYYY') between TO_DATE('"&tglawal&"','MM/DD/YYYY') and TO_DATE('"&tglakhir&"','MM/DD/YYYY') " & _
					addfilterSelect & _
					" group by tgltrx, type, itemcode, origin, store, " & _ 
					" dest, inqty, outqty, fcba, store, note, fcentry, stockawal, stockakhir, trxcode, nominal, companyname order by tgltrx "
	else
		sqlView =	" select  " & _ 
					" tgltrx, type, itemcode, origin, dest, inqty, outqty, fcba, " & _ 
					" store, note, fcentry, sum(stockawal+inqty)-sum(outqty) stockawal, stockakhir, trxcode, nominal, companyname" & _
					" from stock_history" & _
					" where fcba = '"& businessunit &"' and itemcode = '"& productstock &"' " & _
					" AND TO_DATE(tgltrx,'MM/DD/YYYY') between TO_DATE('"&tglawal&"','MM/DD/YYYY') and TO_DATE('"&tglakhir&"','MM/DD/YYYY') " & _
					addfilterSelect & _
					" group by tgltrx, type, itemcode, origin, store, " & _
					" dest, inqty, outqty, fcba, store, note, fcentry, stockakhir, trxcode, nominal, companyname order by tgltrx "
	end if
	' response.Write sqlview
	set rsview = connStr.Execute(sqlview)
	'response.Flush()
%>
	<body>
		<div id="wrapper">
			<!--#include file="print_kartustock.asp"-->
			<%
				sqlCompany = "SELECT companyname FROM stock_history" & _
								" where fcba = '"& businessunit &"' and itemcode = '"& productstock &"' " & _
								" AND TGLTRX BETWEEN '"& tglawal &"' " & _
								" AND '"& tglakhir &"' "
				' response.Write SaldoAwalSql
				set rsCmpny = connStr.Execute(sqlCompany)
			%>
			<table border="0" id="kartustock-table" width="100%">
				<thead>
					<tr>
						<th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
							<div id="header-wrapper">
								<div id="head_judul" style="float:left; text-align:left;">
									<h3><%=rsCmpny("companyname")%></h3>
									<h5><%=fcbaname%></h5>
								</div>
								<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
									<h2>KARTU STOCK</h2>
								</div>
							</div>
						</th>
					</tr>
					<tr style="height: 50px; background-color: #45aa12ff; border: 1px solid #000">
						<th>COMPANY</th>
						<th>FCBA</th>
						<th>STORE</th>
						<th>TANGGAL</th>
						<th>NO. BUKTI</th>
						<th>MODUL</th>
						<th>MASUK (IN)</th>
						<th>KELUAR (OUT)</th>	  
						<th>SALDO</th>
						<th>NOMINAL</th>
					</tr>
					<%
						if trim(wrhouse) = "All" then
							addfilterSelect1 = " AND STORECODE IN (select fccode from store where fcba = '"&businessunit&"' and activation = 'Y')"
						else
							addfilterSelect1 = " AND STORECODE IN ('"&wrhouse&"', '-')"
						end if

						SaldoAwalSql = "select sum(openingbalanceqty) openingbalanceqty, sum(openingbalanceamt) amt from stockledger where year = TO_CHAR(TO_DATE('"& tglawal &"', 'MM/DD/YYYY'), 'YYYY') " & _ 
				   					   " and month = TO_CHAR(TO_DATE('"& tglakhir &"', 'MM/DD/YYYY'), 'MM') AND itemcode = '"& productstock &"' AND fcba = '"& businessunit &"' "&addfilterSelect1&"" 

						' response.Write SaldoAwalSql
						set rsSaldoAwal = connStr.Execute(SaldoAwalSql)
					%>
					<tr style="height: 25px; background-color: #ccc; font-size: 12px; border: 1px solid #000">
						<th><%=rsview("COMPANYNAME")%></th>
						<th><%=rsview("FCBA")%></th>
						<th colspan="6" align="right">SALDO AWAL&nbsp;</th>
						<th align="right"><%=FormatNumber(rsSaldoAwal("openingbalanceqty"), 2)%>&nbsp;</th>
						<th align="right"><%=FormatNumber(rsSaldoAwal("amt"), 2)%>&nbsp;</th>
					</tr>
				</thead>
				<%
					dim inqtyTtl, endingTtl, saldoAwal, saldoJalan
						inqtyTtl = 0
						endingTtl = 0
						saldoAwal = cdbl(rsSaldoAwal("openingbalanceqty"))  ' saldo awal dari query
						saldoJalan = saldoAwal  
						nominalTtl = 0

					While Not rsview.EOF
						dim inqty, outqty, ttlamt
						inqty = cdbl(to_nol(rsview("INQTY")))
						outqty = cdbl(to_nol(rsview("OUTQTY")))
						nominal  = cdbl(to_nol(rsview("NOMINAL")))
						' hitung saldo jalan
						saldoJalan = saldoJalan + inqty - outqty
						nominalTtl = nominalTtl + nominal
				%>
				<tbody>
					<tr style="height: 25px; font-size: 12px; border: 1px solid #000">					
						<td align="center" style="<%=cmdstyle%>"><%=rsview("COMPANYNAME")%></td> 
						<td align="center" style="<%=cmdstyle%>"><%=rsview("FCBA")%></td> 
						<td align="center" style="<%=cmdstyle%>"><%=rsview("STORE")%>&nbsp;</td>
						<td align="center" style="<%=cmdstyle%>"><%=rsview("TGLTRX")%>&nbsp;</td>
						<td align="center" style="<%=cmdstyle%>"><%=rsview("TRXCODE")%> </td>
						<td align="center" style="<%=cmdstyle%>"><%=rsview("NOTE")%> </td>
						<td align="right" style="<%=cmdstyle%>"><%=FormatNumber(inqty,2)%>&nbsp; </td>
						<td align="right" style="<%=cmdstyle%>"><%=FormatNumber(outqty,2)%>&nbsp; </td>
						<td align="right" style="<%=cmdstyle%>"><%=FormatNumber(saldoJalan,2)%>&nbsp;</td>
						<td align="right" style="<%=cmdstyle%>"><%=FormatNumber(nominal,2)%>&nbsp;</td>
					</tr>
				</tbody>
				<%
					inqtyTtl = inqtyTtl + inqty
        			endingTtl = endingTtl + outqty
					rsview.MoveNext
					Wend
					rsview.close()
				%>
				<tfoot>
					<tr style="height: 25px; background-color: #ccc; font-size: 12px; border: 1px solid #000">
						<th align="right">&nbsp;</th>
						<th align="right">&nbsp;</th>
						<th align="right">&nbsp;</th>
						<th align="right">&nbsp;</th>
						<th align="right">&nbsp;</th>
						<th align="right">&nbsp;</th>
						<th align="right"><%=FormatNumber(cdbl(to_nol(inqtyTtl)), 2)%>&nbsp;</th>
						<th align="right"><%=FormatNumber(cdbl(to_nol(endingTtl)), 2)%>&nbsp;</th>
						<th align="right"><%=FormatNumber(cdbl(rsSaldoAwal("openingbalanceqty"))+cdbl(to_nol(inqtyTtl))-cdbl(to_nol(endingTtl)), 2)%>&nbsp;</th>
						<th align="right"><%=FormatNumber(nominalTtl,2)%>&nbsp;</th>
					</tr>
				</tfoot>
			</table>
		</div>
	</body>
</html>	