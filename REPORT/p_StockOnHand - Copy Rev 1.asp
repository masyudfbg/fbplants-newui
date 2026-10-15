<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockOnHand, strStockOnHand
	Dim v_txtStockCode
	Dim NO, cClass, nTotal

	v_txtStockCode		=	trim(request("txtStockCode"))
	
	
	strStockOnHand		=	" select fcba" & _
	                        "        , storecode" & _
							"        , store_api.get_fcname(storecode, fcba) storename" & _
							"        , stockcode" & _
							"        , stock_api.get_stockname(stockcode, fcba) stockname" & _
							"        , stock_api.get_uom(stockcode, fcba) uom" & _
							"        , NVL(quantity, 0) quantity" & _
							" from stockonhand " & _
							" where stockcode not in (" & _
							"                 select fccode from stock" & _
							"                 where fcba = '" & Session("Fcba") & "'" & _
							"                 and fcba||fccode in (" & _
							"                                  select fcba||stockcode from nursery" & _
							"                                  where fcba = '" & Session("Fcba") & "'" & _
							"                                  )" & _
							"                 )" & _
							"       and fcba = '" & Session("Fcba") & "'"


	if  len(trim(v_txtStockCode)) > 0 then
		strStockOnHand	=	strStockOnHand & " and stockcode = '"& v_txtStockCode &"'"
	end if
	
		strStockOnHand	=	strStockOnHand & " order by storecode, stockcode"

	Set rsStockOnHand = Server.CreateObject("ADODB.Recordset")
	rsStockOnHand.Open strStockOnHand, connstr, adOpenStatic 
	
%>


<html>

<head>
	<title>ePlaS Reporting - Stock On Hand</title>
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
	
	<style type="text/css">
	@media print
	{
	.printhide{display:none;}
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
			<form name="frmEntry" method="post">
			<div class="printhide">
				<table id="form-table" class="centered-margin">
					<tr>
						<td><font face="verdana" size="1">Kode Stock</font></td>						
						<td align="center">
							<input type="text" maxlength="20" name="txtStockCode" size="20" value="<%=v_txtStockCode%>" title="Input Stock code">
						</td>
						<td align="center"><input type="submit" value="View" /></td>
					</tr>
				</table>
			</div>
			</form>
		</div>
		
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					<th colspan="6" style="background-color:#FFFFFF; border:none;">  
						<div id="header-wrapper">
							<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
								<h4>PT. <%=Session("COMPANY")%></h4>
								<h5><%=Session("BA")%></h5>
							</div>
							<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
								<h2>LAPORAN STOCK ON HAND</h2>
								<% If Len(Trim(v_txtStockCode)) > 0 Then %>
								<h4>Stock Code : <%=v_txtStockCode%></h4>
								<br/>
								<%End If%>
							</div>
						</div>
					</th>
				</tr>
				<tr>
					<th>STORE CODE</th>
					<th>STORE NAME</th>
					<th>STOCK CODE</th>
					<th>STOCK NAME</th>
					<th>SATUAN</th>
					<th>QUANTITY</th>
				</tr>
			</thead>
		<tbody>
<%if rsStockOnHand.eof and rsStockOnHand.eof then%>
	<tr>
		<td colspan="6" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		nTotal = 0
		while not rsStockOnHand.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="left"><%=rsStockOnHand("STORECODE")%></td>
				<td align="left"><%=rsStockOnHand("STORENAME")%></td>
				<td align="left"><%=rsStockOnHand("STOCKCODE")%></td>
				<td align="left"><%=rsStockOnHand("STOCKNAME")%></td>
				<td align="left"><%=rsStockOnHand("UOM")%></td>
				<td align="right"><%=FbpNum2(rsStockOnHand("QUANTITY"))%></td>				
			</tr>
<%
			NO = NO + 1
			nTotal = nTotal + Cdbl(rsStockOnHand("QUANTITY"))
			rsStockOnHand.MoveNext
		Wend
	rsStockOnHand.Close
	set rsStockOnHand = Nothing
%>
			<tr style="background-color:#F4F4A0;">
				<td align="center" colspan="5"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FbpNum2(nTotal)%></strong></td>
			</tr>
 </tbody>
</table>
   <table width="90%" class="data" style="margin:auto;">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("FCMANAGER")%>&nbsp;&nbsp;</u></strong></td>
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("KTU")%>&nbsp;&nbsp;</u></strong></td>
	   </tr>
	   	  <tr> 
		<td align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
</div>