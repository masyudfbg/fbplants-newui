<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>

<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Analisa Biaya Pabrik</title>
<script language="JavaScript" src="/js/date-picker.js"></script> 
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

<body>
	<div id="wrapper">
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<!--#include file="FromSelectHDRMILL.html"-->
	
	
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
	   
				
		
		<h3>LAPORAN OUT STANDING DELIVERY NOTE</h3>
		
	</div>
	</div>
<%
Dim strtbs, rsha,TBSprocess_BI,TBSprocess_SBI,ProdCPO_BI,ProdCPO_SBI,ProdKER_BI,ProdKER_SBI,Palm_BI,Palm_SBI, strprod

%>	

<table border="0" width="100%" id="product-table">
	
	<tr>
		
		<th style="color:#FFFFFF" >Tanggal</th>
		<th style="color:#FFFFFF" >No Kontrak</th>
		
		<th style="color:#FFFFFF" >Qty Kontrak</th>
		<th style="color:#FFFFFF" >No Kendaraan</th>
		
		<th style="color:#FFFFFF" >Delivery Note No</th>
		<th style="color:#FFFFFF" >Qty Delivery</th>
		
		<th style="color:#FFFFFF" >Sisa</th>
	
	
	</tr>
	
	<%
	Dim strdtl , strgrp ,cdgrp,cdmesin,cdgrpdesc  ,noline , subAmtgrpBI , GRAmtgrpBI, subAmtgrpSBI , GRAmtgrpSBI , grpDesc
	
			strdtl = " select * from ( select a.dnote_date tanggal, a.contractcode no_kontrak,  " & _
					 " (select totalweight from deliverycontrac where fcba = a.fcba and contractcode = a.contractcode) qty_contract, " & _
					 " (  select distinct vehicleno from weighbridgeticket where fcba = a.fcba and chitno = a.wbt_chitno and product_productcode ='CPO' ) no_kendaraan  , " & _
					 "  a.DELIVERYNOTECODE, a.quantity from deliverynote a ) order by no_kontrak, tanggal "
    			
			
		
	
	Set rsha = Server.CreateObject("ADODB.Recordset")
	'rsha.CursorLocation = adUseClient
	rsha.Open strdtl, connstr, adOpenStatic
    response.write strdtl
	
	if Not rsha.eof then
		  while Not rsha.eof %>
	      <tr>  
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0%</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
		    </tr>		
		<%
		rsha.MoveNext
		  wend

	end if		
	%>
	
	
</table>
<!--#include file="bottom.Asp"-->