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
<title>ePlaS Reporting - Listing Tiket TBS</title>
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
				<td><font face="verdana" size="1">Tgl Tiket</font></td>
				<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
				</td>
				<td><font face="verdana" size="1">s/d</font></td>
				<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
				</td>
				<td><font face="verdana" size="1">Supplier TBS</font></td>
				<td><select name="v_txtsupllier" id="v_txtsupllier" class="lebarselect"/>
						<option value="all">Semua Supplier</option>
					</select>
				</td>
				<td><input type="submit" value="View" /></td>
				</form>
			</tr>
		</table>
		</div>
		</form>
	</div>

	<% if v_txtsupllier = "all" then%>

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
						<h2>LAPORAN LISTING TBS</h2>		
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
		<th  style="color:#fff;">SUPPLIER</th>
		<th  style="color:#fff;">TANGGAL</th>		
		<th  style="color:#fff;">NETTO 1</th>
		<th  style="color:#fff;">GRADING</th>
		<th  style="color:#fff;">NETTO 2</th>
		<th  style="color:#fff;">HARGA TBS</th>
		<th  style="color:#fff;">TOTAL</th>
	</tr>
	</thead>
	<tbody>
	<%
	Dim rsBeliTBSall, strBeliTBSall

	strBeliTBSall = " select docketdate,description,price,sum(weightin),sum(weightout),sum(weightin-weightout) as netto1, " & _
	" sum(sortasiwgt) as sortasiwgt,sum(weightin-weightout-sortasiwgt) as netto2 " & _
	" from( " & _
	"	select distinct h.docketdate,c.description,i.price,d.weightin,d.weightout,d.sortasiwgt " & _
	"	from lhptikethdr h,lhpffbwbtrn d,customer c,invoice_pricetbs i " & _
	"	where h.docketid=d.docketid and d.suppliercode=c.customercode " & _
	"	and h.docketdate between to_date('" & v_txtDateFrom & "', 'mm/dd/yyyy') " & _
	"	and to_date('" & v_txtDateTo & "', 'mm/dd/yyyy') and d.suppliercode not in('SP0001','SP0002','SP0003') " & _
	"   and i.suppliercode=d.suppliercode and i.ondate=h.docketdate " & _
	" )" & _
	" group by docketdate,description,price " & _
	" ORDER BY docketdate,description"
		
    'response.write strBeliTBSall	
	Set rsBeliTBSall = Server.CreateObject("ADODB.Recordset")
	rsBeliTBSall.Open strBeliTBSall, connstr, adOpenStatic
   	if Not rsBeliTBSall.eof then
	    Dim netto1,grading,netto2,harga,totalrp 
		Dim ttlnetto1,ttlgrading, ttlnetto2,ttlgrand,granttl
		
		ttlkg1			=	0
		ttlgrading 		=	0
		ttlkg    		=	0
		ttlgrand 		=	0
		while Not rsBeliTBSall.eof 
			docketdate 			= rsBeliTBSall("docketdate")
			description 		= rsBeliTBSall("description")
			netto1 				= round(cdbl(rsBeliTBSall("netto1")))
			grading 			= round(cdbl(rsBeliTBSall("sortasiwgt")))
			netto2 				= round(cdbl(rsBeliTBSall("netto2")))
			harga 				= round(cdbl(rsBeliTBSall("price")))
			
			totalrp 			= cdbl(netto2) * CDbl(harga)
		%>
		<tr>
			<td align="left"><%=rsBeliTBSall("description")%></td>
			<td align="left"><%=rsBeliTBSall("docketdate")%></td>
			<td align="right"><%=FormatNumber(netto1,2)%></td>
			<td align="right"><%=FormatNumber(grading,2)%></td>
			<td align="right"><%=FormatNumber(netto2,2)%></td>
			<td align="right"><%=FormatNumber(harga,2)%></td>
			<td align="right"><%=FormatNumber(totalrp,2)%></td>
			
		  </tr>	
		      
		<%
		
			ttlnetto1		=	ttlnetto1 + cdbl(netto1)
            ttlgrading 		=	ttlgrading + cdbl(grading)
            ttlnetto2    	=	ttlnetto2 + cdbl(netto2)
            ttlgrand 		=	ttlgrand + cdbl(totalrp)
			ttlharga		= 	ttlgrand/ttlnetto2
			rsBeliTBSall.MoveNext
		wend
		%>
		  <tr bgcolor="#F4F4A0">
			<td align="center" colspan="<%=add_newspan+2%>"><strong>TOTAL</td>
			<td align="right"><strong><%=FormatNumber(ttlnetto1,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlgrading,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlnetto2,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlharga,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlgrand,2)%></strong></td>
		  </tr>	
		<%
	end if
		rsBeliTBSall.Close
		set rsBeliTBSall = Nothing
		%>
	</tbody>
	</table>

<%else%>

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
						<h2>LAPORAN LISTING TBS</h2>		
						<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
							<h5>TANGGAL: <%=v_txtdateto%></h5> <br>
							<h5>SUPPLIER:</h5>
						<% else %>
							<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
							<h5>SUPPLIER:</h5>
						<% end if %>
					</div>
					<br>
					</div>
				</th>
			</tr>
			<br>
     <tr>
		<th  style="color:#fff;">TANGGAL</th>		
		<th  style="color:#fff;">NETTO 1</th>
		<th  style="color:#fff;">GRADING</th>
		<th  style="color:#fff;">NETTO 2</th>
		<th  style="color:#fff;">HARGA TBS</th>
		<th  style="color:#fff;">TOTAL</th>
	</tr>
	</thead>
	<tbody>
	<%
	'Dim rsBeliTBSall, strBeliTBSall

	strBeliTBSall = "select docketdate,description,price,sum(weightin),sum(weightout),sum(weightin-weightout)as netto1, " & _
	" sum(sortasiwgt)as sortasiwgt,sum(weightin-weightout-sortasiwgt) as netto2 " & _
	" from( " & _
	"	select distinct h.docketdate,c.description,i.price,d.weightin,d.weightout,d.sortasiwgt " & _
	"	from lhptikethdr h,lhpffbwbtrn d,customer c,invoice_pricetbs i " & _
	"	where h.docketid=d.docketid and d.suppliercode=c.customercode " & _
	"	and h.docketdate between to_date('" & v_txtDateFrom & "', 'mm/dd/yyyy') " & _
	"	and to_date('" & v_txtDateTo & "', 'mm/dd/yyyy') AND d.suppliercode='" & v_txtsupllier & "' " & _
	"   and d.suppliercode not in('SP0001','SP0002','SP0003') and i.ondate=h.docketdate and i.suppliercode=d.suppliercode " & _
	")" & _
	" group by docketdate,description,price " & _
	" ORDER BY docketdate,description"
		
    'response.write strBeliTBSall	
	Set rsBeliTBSall = Server.CreateObject("ADODB.Recordset")
	rsBeliTBSall.Open strBeliTBSall, connstr, adOpenStatic
   	if Not rsBeliTBSall.eof then
		'Dim ttlkg1,ttlgrading, ttlkg,ttlgrand,granttl
		Dim bydocketdate,bydescription,bynetto1,bygrading,bynetto2,byharga,bytotalrp 
		Dim byttlnetto1,byttlgrading, byttlnetto2,byttlgrand
		byttlnetto1		=	0
		byttlgrading 	=	0
		byttlnetto2    	=	0
		byttlharga   	=	0
		byttlgrand 		=	0
		while Not rsBeliTBSall.eof 
			bydocketdate 		= rsBeliTBSall("docketdate")
			bydescription 		= rsBeliTBSall("description")
			bynetto1 			= round(cdbl(rsBeliTBSall("netto1")))
			bygrading 			= round(cdbl(rsBeliTBSall("sortasiwgt")))
			bynetto2 			= round(cdbl(rsBeliTBSall("netto2")))
			byharga 			= round(cdbl(rsBeliTBSall("price")))
			
			bytotalrp 			= cdbl(bynetto2) * CDbl(byharga)
		%>
		<tr>
			<td align="left"><%=bydocketdate%></td>
			<td align="right"><%=FormatNumber(bynetto1,2)%></td>
			<td align="right"><%=FormatNumber(bygrading,2)%></td>
			<td align="right"><%=FormatNumber(bynetto2,2)%></td>
			<td align="right"><%=FormatNumber(byharga,2)%></td>
			<td align="right"><%=FormatNumber(bytotalrp,2)%></td>
			
		  </tr>	
		      
		<%
		
			byttlnetto1		=	byttlnetto1 + cdbl(bynetto1)
            byttlgrading 	=	byttlgrading + cdbl(bygrading)
            byttlnetto2    	=	byttlnetto2 + cdbl(bynetto2)
            byttlgrand 		=	byttlgrand + cdbl(bytotalrp)
			byttlharga		= 	byttlharga/byttlnetto2
			
			rsBeliTBSall.MoveNext
		wend
		%>
		  <tr bgcolor="#F4F4A0">
			<td align="center" colspan="<%=add_newspan+1%>"><strong>TOTAL</td>
			<td align="right"><strong><%=FormatNumber(byttlnetto1,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(byttlgrading,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(byttlnetto2,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(byttlharga,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(byttlgrand,2)%></strong></td>
		  </tr>	
		<%
	end if
		rsBeliTBSall.Close
		set rsBeliTBSall = Nothing
		%>
	</tbody>
	</table>

<% end if%>
<!--#include file="bottom.Asp"-->