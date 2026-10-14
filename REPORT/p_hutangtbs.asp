<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtDateFrom, v_txtDateTo, v_txtsupllier

v_txtfcba 			  = Request("txtfcba")

v_txtDateFrom = trim(Request("txtDateFrom"))
v_txtDateTo = trim(Request("txtDateTo"))
v_txtsupllier = trim(Request("v_txtsupllier"))
ScriptName = Request.ServerVariables("SCRIPT_NAME")

if "" & v_txtDateFrom = "" then v_txtDateFrom = textDate(date)
if "" & v_txtDateTo = "" then v_txtDateTo = textDate(date)


%>
<html>

<head>
<title>ePlaS Reporting - Hutang TBS</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
		xAddEventListener(window, 'load',
		  function() {
			new xTableHeaderFixed('xthf-site', window);
		  }, false
		);
	</script>
	<link rel="stylesheet" href="/css/jquery-ui.css" />
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
    <script>
		$(function() {
			$( ".datepicker" ).datepicker();
			getSup();
		});
	</script>
	<script language="JavaScript" src="/js/date-picker.js"></script> 
	<script language="JavaScript">
		var newwin;
		function launchwin(winurl,winname,winfeatures)
		{
			//This launches a new window and then
			//focuses it if window.focus() is supported.
			newwin = window.open(winurl,winname,winfeatures);

		}

		function getSup(jnsinvoice) {
				$('#v_txtsupllier').select2({
					ajax: {
						delay: 1000,
						url : '/getviewsuppliertbs.asp',
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

	<style type="text/css">
		@media print
		{
		.printhide{display:none;}
		}

		.lebarselect{
				width:240px;
		}
	</style>
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
		<form name="frmEntry" method="post" action="<%=ScriptName%>">
		<div class="printhide">
		
		<table id="form-table" class="centered-margin">
			<tr>
				<form name="frmEntry" method="post" action="<%=ScriptName%>">
				<td><font face="verdana" size="1">Tgl Invoice</font></td>
				<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
				</td>
				<td><font face="verdana" size="1">s/d</font></td>
				<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
				</td>
				<td><font face="verdana" size="1">Supplier TBS</font></td>
				<td><select name="v_txtsupllier" id="v_txtsupllier" class="lebarselect"/>
						<option value="">Pilih Supplier</option>
					</select>
				</td>
				<td><input type="submit" value="View" /></td>
				</form>
			</tr>
		</table>
		</div>
		</form>
	</div>
	<table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
			<tr>
				<th colspan="<%=add_newspan+10%>" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
				
					</div> 
					<br>
					<br>
	 
					<div id="subheader" class="centered">
						<h2>LAPORAN HUTANG TBS</h2>		
						<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
							<h5>TANGGAL: <%=v_txtdateto%></h5>
						<% else %>
							<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
						<% end if %>
					</div>
					<br>
					</div>
				</th>
			</tr>
			<br>
     <tr>
		<th  style="color:#fff;">TGL INVOICE</th>
		<th  style="color:#fff;">NO INVOICE</th>		
		<th  style="color:#fff;">PERIODE TIKET</th>
		<th  style="color:#fff;">TOTAL INV</th>
		<th  style="color:#fff;">PEMBAYARAN</th>
		<th  style="color:#fff;">SISA</th>
	</tr>
	</thead>
	<tbody>
	<%
	Dim rsHutangTBS, strHutangTBS , noline
	
	strHutangTBS	=	" select h.invoiceno,h.transdate,c.description,to_char(h.valid_from, 'dd/mm/yyyy') || ' - '|| to_char(h.valid_until,'dd/mm/yyyy') as periode, " & _
						" h.grandtotal,h.payment " & _
						" from invoiceservicehdr h,customer c " & _
						" where h.suppliername=c.customercode AND h.fcba=c.fcba "
						

if  len(trim(v_txtDateFrom)) > 0 then
	strHutangTBS = strHutangTBS & " AND h.transdate >=	TO_DATE('" & v_txtDateFrom & "','MM/DD/YYYY') "
					
end if
if  len(trim(v_txtDateTo)) > 0 then
	strHutangTBS = strHutangTBS &  " AND h.transdate <= TO_DATE('" & v_txtDateTo & "','MM/DD/YYYY')"		   
end if


if  len(trim(v_txtsupllier)) <> 0 then
	strHutangTBS = strHutangTBS &  " AND h.suppliername = '" & v_txtsupllier & "' " & _
	" group by h.invoiceno,h.transdate,c.description,h.grandtotal,h.payment, " & _
	" to_char(h.valid_from, 'dd/mm/yyyy') || ' - '|| to_char(h.valid_until,'dd/mm/yyyy')"
	'response.Write strHutangTBS
	
end if	

	
   '' response.write strHutangTBS	
	Set rsHutangTBS = Server.CreateObject("ADODB.Recordset")
	rsHutangTBS.Open strHutangTBS, connstr, adOpenStatic
    
	if rsHutangTBS.bof and rsHutangTBS.eof then
	%>
	<tr>
			<td align="left" colspan="<%=add_newspan+10%>"><strong>DATA TIDAK DITEMUKAN</strong></td>
	</tr>	
	<%
	elseif Not rsHutangTBS.eof then
		Dim sisa,ttlinv, ttlbyr,ttlsisa
		
		ttlbyr			=	0
		ttlsisa 		=	0
		while Not rsHutangTBS.eof 

		sisa = cdbl(rsHutangTBS("grandtotal"))-cdbl(rsHutangTBS("payment"))
		
		%>
		<tr>
			<td align="left"><%=rsHutangTBS("transdate")%></td>
			<td align="left"><%=rsHutangTBS("invoiceno")%></td>
			<td align="left"><%=rsHutangTBS("periode")%></td>
			<td align="right"><%=FormatNumber(rsHutangTBS("grandtotal"),0)%></td>
			<td align="right"><%=FormatNumber(rsHutangTBS("payment"),0)%></td>
			<td align="right"><%=FormatNumber(sisa,0)%></td>
		  </tr>	
		      
<%
			ttlinv = ttlinv + cdbl(rsHutangTBS("grandtotal"))
			ttlbyr = ttlbyr + cdbl(rsHutangTBS("payment"))
			ttlsisa = ttlsisa + cdbl(sisa)
		
			rsHutangTBS.MoveNext
		wend
%>
		  <tr bgcolor="#F4F4A0">
			<td align="center" colspan="<%=add_newspan+3%>"><strong>TOTAL</td>
			<td align="right"><strong><%=FormatNumber(ttlinv,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlbyr,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlsisa,0)%></strong></td>
		  </tr>	
<%
	end if
rsHutangTBS.Close
set rsHutangTBS = Nothing
%>
	</tbody>
</table>
<!--#include file="bottom.Asp"-->