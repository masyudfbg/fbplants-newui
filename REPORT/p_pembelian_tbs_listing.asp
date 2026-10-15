<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtDateFrom, v_txtDateTo, v_txtsupllier

v_txtfcba 			  = session("fcba")

v_txtDateFrom = trim(Request("txtDateFrom"))
v_txtDateTo = trim(Request("txtDateTo"))
v_txtsupllier = trim(Request("v_txtsupllier"))
ScriptName = Request.ServerVariables("SCRIPT_NAME")

if "" & v_txtDateFrom = "" then v_txtDateFrom = textDate(date)
if "" & v_txtDateTo = "" then v_txtDateTo = textDate(date)


%>
<html>

<head>
<title>ePlaS Reporting - Pembelian TBS</title>
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
				<th colspan="<%=add_newspan+12%>" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
				
					</div> 
					<br>
					<br>
	 
					<div id="subheader" class="centered">
						<h2>LAPORAN PEMBELIAN TBS</h2>		
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
	 	<th  style="color:#fff;">BULAN</th>
		<th  style="color:#fff;">TAHUN</th>
		<th  style="color:#fff;">SUPPLIER</th>
		<th  style="color:#fff;">TANGGAL</th>		
		<th  style="color:#fff;">NETTO 1</th>
		<th  style="color:#fff;">GRADING</th>
		<th  style="color:#fff;">NETTO 2</th>
		<th  style="color:#fff;">HARGA TBS</th>
		<th  style="color:#fff;">TOTAL</th>
		<th  style="color:#fff;">PPH 22</th>
		<th  style="color:#fff;">JUMLAH TAGIHAN</th>
	</tr>
	</thead>
	<tbody>
	<%
	Dim rsBeliTBSall, strBeliTBSall,netto1

'	strBeliTBSall = " select c.description,h.transdate,sum(d.kg1) netto1,sum(d.gradding)sortasiwgt,sum(d.kg)ffbweight,nvl(d.rate,0) price " & _ 
'	" from invoiceservicehdr h " & _
'	" inner join invoiceservicedtl d on h.INVOICENO=d.INVOICENO and h.fcba=d.fcba " & _
'	" inner join customer c on h.SUPPLIERNAME=c.CUSTOMERCODE and h.fcba=c.fcba " & _
'	" where h.transdate between to_date('" & v_txtDateFrom & "','mm/dd/yyyy') and to_date('" & v_txtDateTo & "', 'mm/dd/yyyy') " & _
'	" and h.fcba='" & v_txtfcba & "' " & _
'	" group by c.description,h.transdate,d.rate order by h.transdate,c.description asc"

    strBeliTBSall = " select to_char(hd.docketdate,'mm') as bulan,to_char(hd.docketdate,'yyyy') as tahun, a.areaname AS description,to_char(hd.docketdate,'dd/mm/yyyy') AS transdate,nvl(p.price,0) price,nvl(sum(dt.weightin),0) as weightin,nvl(sum(dt.weightout),0)as weightout,nvl(sum(dt.sortasiwgt),0) as sortasiwgt,nvl(sum(dt.ffbweight),0) as ffbweight " & _
  	" from LHPTIKETHDR hd " & _
  	" left join LHPFFBWBTRN dt on hd.docketid=dt.docketid AND hd.fcba=dt.fcba " & _
	" left join LHPAREAMST a on dt.areaid=a.areaid AND dt.fcba=a.fcba " & _
	" left join invoice_pricetbs p on p.suppliercode=a.eplascode and p.ONDATE=hd.DOCKETDATE  " & _
	" where hd.DOCKETDATE between to_date('" & v_txtDateFrom & "','mm/dd/yyyy') and to_date('" & v_txtDateTo & "', 'mm/dd/yyyy') " & _
	" and hd.fcba='" & v_txtfcba & "' " & _
	" group by to_char(hd.docketdate,'mm'),to_char(hd.docketdate,'yyyy'),a.areaname,hd.docketdate,p.price order by hd.docketdate,a.areaname asc"
		
   ' response.write strBeliTBSall	
	Set rsBeliTBSall = Server.CreateObject("ADODB.Recordset")
	rsBeliTBSall.Open strBeliTBSall, connstr, adOpenStatic
   	if Not rsBeliTBSall.eof then
	    Dim gradingqty,netto2,totalrp,weightin,weightout,pricetbs
		Dim ttlkg1,ttlgrading, ttlkg,ttlgrand,granttl
		Dim pphrp,amountdanpph,totalpph,totalamountpph
		
		ttlkg1			=	0
		ttlgrading 		=	0
		ttlkg    		=	0
		ttlgrand 		=	0
		ttlharga 		=	0
		while Not rsBeliTBSall.eof 
			
			netto1 		= cdbl(rsBeliTBSall("weightin"))-cdbl(rsBeliTBSall("weightout"))
			gradingqty  = cdbl(rsBeliTBSall("sortasiwgt"))
			gradingqty	= FormatNumber(gradingqty,0)
			netto2 	    = cdbl(netto1)-cdbl(gradingqty)
			pricetbs 	= cdbl(rsBeliTBSall("price"))
			granttl 	= cdbl(netto2) * cdbl(pricetbs)
			pphrp		= cdbl(granttl)*(0.25/100)
			amountdanpph= cdbl(granttl)-cdbl(pphrp)
		%>
		<tr>
			<td align="left"><%=rsBeliTBSall("bulan")%></td>
			<td align="left"><%=rsBeliTBSall("tahun")%></td>
			<td align="left"><%=rsBeliTBSall("description")%></td>
			<td align="left"><%=rsBeliTBSall("transdate")%></td>
			<td align="right"><%=FormatNumber(netto1,2)%></td>
			<td align="right"><%=FormatNumber(gradingqty,2)%></td>
			<td align="right"><%=FormatNumber(netto2,2)%></td>
			<td align="right"><%=FormatNumber(pricetbs,2)%></td>
			<td align="right"><%=FormatNumber(granttl,2)%></td>
			<td align="right"><%=FormatNumber(pphrp,2)%></td>
			<td align="right"><%=FormatNumber(amountdanpph,2)%></td>
			
		  </tr>	
		      
		<%
		
            ttlkg1			=	ttlkg1 + cdbl(netto1)
            ttlgrading 		=	ttlgrading + cdbl(gradingqty)
            ttlkg    		=	cdbl(ttlkg) + cdbl(netto2)
            ttlgrand 		=	cdbl(ttlgrand) + cdbl(granttl)
			ttlharga		= 	cdbl(ttlharga)+cdbl(pricetbs)
			'ttlhargaview	=   cdbl(ttlharga)/cdbl(ttlkg)
			totalpph		= 	totalpph+pphrp
			totalamountpph	= 	totalamountpph+amountdanpph
			rsBeliTBSall.MoveNext
		wend
		%>
		  <tr bgcolor="#F4F4A0">
			<td align="center" colspan="4"><strong>TOTAL</td>
			<td align="right"><strong><%=FormatNumber(ttlkg1,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlgrading,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlkg,2)%></strong></td>
			<td align="right"><strong>0</strong></td>
			<td align="right"><strong><%=FormatNumber(ttlgrand,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(totalpph,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(totalamountpph,2)%></strong></td>
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
						<h2>LAPORAN PEMBELIAN TBS</h2>		
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
		<th  style="color:#fff;">BULAN</th>
		<th  style="color:#fff;">TAHUN</th>
		<th  style="color:#fff;">SUPPLIER</th>
		<th  style="color:#fff;">TANGGAL</th>		
		<th  style="color:#fff;">NETTO 1</th>
		<th  style="color:#fff;">GRADING</th>
		<th  style="color:#fff;">NETTO 2</th>
		<th  style="color:#fff;">HARGA TBS</th>
		<th  style="color:#fff;">TOTAL</th>
		<th  style="color:#fff;">PPH 22</th>
		<th  style="color:#fff;">JUMLAH TAGIHAN</th>
	</tr>
	</thead>
	<tbody>
	<%
	'Dim rsBeliTBSall, strBeliTBSall

'	strBeliTBSall = " select c.description,h.transdate,sum(d.kg1) netto1,sum(d.gradding)sortasiwgt,sum(d.kg)ffbweight,nvl(d.rate,0) price " & _ 
'	" from invoiceservicehdr h " & _
'	" inner join invoiceservicedtl d on h.INVOICENO=d.INVOICENO and h.fcba=d.fcba " & _
'	" inner join customer c on h.SUPPLIERNAME=c.CUSTOMERCODE and c.fcba=h.fcba " & _
'	" where h.transdate between to_date('" & v_txtDateFrom & "','mm/dd/yyyy') and to_date('" & v_txtDateTo & "', 'mm/dd/yyyy') " & _
'	" and h.fcba='" & v_txtfcba & "' and h.suppliername='" & v_txtsupllier & "' " & _
'	" group by c.description,h.transdate,d.rate order by h.transdate,c.description asc"

	strBeliTBSall = " select to_char(hd.docketdate,'mm') as bulan,to_char(hd.docketdate,'yyyy') as tahun,a.areaname AS description,to_char(hd.docketdate,'dd/mm/yyyy') AS transdate,nvl(p.price,0) price,sum(dt.weightin) as weightin,sum(dt.weightout)as weightout,sum(dt.sortasiwgt) as sortasiwgt,sum(dt.ffbweight) as ffbweight " & _
  	" from LHPTIKETHDR hd " & _
  	" inner join LHPFFBWBTRN dt on hd.docketid=dt.docketid AND hd.fcba=dt.fcba " & _
	" inner join LHPAREAMST a on dt.areaid=a.areaid AND dt.fcba=a.fcba " & _
	" left join invoice_pricetbs p on p.suppliercode=a.eplascode and p.ONDATE=hd.DOCKETDATE  " & _
	" where hd.DOCKETDATE between to_date('" & v_txtDateFrom & "','mm/dd/yyyy') and to_date('" & v_txtDateTo & "', 'mm/dd/yyyy') " & _
	" and hd.fcba='" & v_txtfcba & "'  AND a.eplascode='" & v_txtsupllier & "'  " & _
	" group by a.areaname,hd.docketdate,p.price order by hd.docketdate,a.areaname asc"
		
  ' response.write strBeliTBSall	
	Set rsBeliTBSall = Server.CreateObject("ADODB.Recordset")
	rsBeliTBSall.Open strBeliTBSall, connstr, adOpenStatic
   	if Not rsBeliTBSall.eof then
		'Dim ttlkg1,ttlgrading, ttlkg,ttlgrand,granttl
		
		ttlkg1			=	0
		ttlgrading 		=	0
		ttlkg    		=	0
		ttlharga   		=	0
		ttlgrand 		=	0
		while Not rsBeliTBSall.eof 
			netto1 		= cdbl(rsBeliTBSall("weightin"))-cdbl(rsBeliTBSall("weightout"))
			gradingqty  = cdbl(rsBeliTBSall("sortasiwgt"))
			gradingqty	= FormatNumber(gradingqty,0)
			netto2 	    = cdbl(netto1)-cdbl(gradingqty)
			pricetbs 	= cdbl(rsBeliTBSall("price"))
			granttl 	= cdbl(netto2) * CDbl(pricetbs)
			pphrp		= cdbl(granttl)*(0.25/100)
			amountdanpph= cdbl(granttl)-cdbl(pphrp)
		%>
		<tr>
			<td align="left"><%=rsBeliTBSall("bulan")%></td>
			<td align="left"><%=rsBeliTBSall("tahun")%></td>
			<td align="left"><%=rsBeliTBSall("description")%></td>
			<td align="left"><%=rsBeliTBSall("transdate")%></td>
			<td align="right"><%=FormatNumber(netto1,2)%></td>
			<td align="right"><%=FormatNumber(gradingqty,2)%></td>
			<td align="right"><%=FormatNumber(netto2,2)%></td>
			<td align="right"><%=FormatNumber(pricetbs,2)%></td>
			<td align="right"><%=FormatNumber(granttl,2)%></td>
			<td align="right"><%=FormatNumber(pphrp,2)%></td>
			<td align="right"><%=FormatNumber(amountdanpph,2)%></td>
		  </tr>	
		      
		<%
		
			ttlkg1			=	ttlkg1 + cdbl(netto1)
            ttlgrading 		=	ttlgrading + cdbl(gradingqty)
            ttlkg    		=	ttlkg + cdbl(netto2)
            ttlgrand 		=	ttlgrand + cdbl(granttl)
			ttlharga		= 	ttlgrand/ttlkg
			totalpph		= 	totalpph+pphrp
			totalamountpph	= 	totalamountpph+amountdanpph
			rsBeliTBSall.MoveNext

		wend
		%>
		<tr bgcolor="#F4F4A0">
			<td align="center" colspan="4"><strong>TOTAL</td>
			<td align="right"><strong><%=FormatNumber(ttlkg1,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlgrading,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlkg,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlharga,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(ttlgrand,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(totalpph,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(totalamountpph,2)%></strong></td>
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



<%
  ' strBeliTBSall = " select a.areaname,hd.docketdate,nvl(p.price,0) price,sum(dt.weightin) as weightin,sum(dt.weightout)as weightout,sum(dt.sortasiwgt) as sortasiwgt,sum(dt.ffbweight) as ffbweight " & _
  '	" from LHPTIKETHDR hd " & _
  '	" inner join LHPFFBWBTRN dt on hd.docketid=dt.docketid " & _
'	" inner join LHPAREAMST a on dt.areaid=a.areaid " & _
'	" left join invoice_pricetbs p on p.suppliercode=a.eplascode and p.ONDATE=hd.DOCKETDATE  " & _
'	" where hd.DOCKETDATE between to_date('" & v_txtDateFrom & "','mm/dd/yyyy') and to_date('" & v_txtDateTo & "', 'mm/dd/yyyy') " & _
'	" and a.UNIT='EXT' and hd.fcba='" & v_txtfcba & "'  AND a.eplascode='" & v_txtsupllier & "' " & _
'	" group by a.areaname,hd.docketdate,p.price order by hd.docketdate,a.areaname asc"
%>