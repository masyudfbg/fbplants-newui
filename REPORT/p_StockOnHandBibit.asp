<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockOnhandBibit, strStockOnhandBibit
	Dim v_txtNurseryCode
	Dim NO, cClass, nTotal

	v_txtNurseryCode		=	trim(request("txtNurseryCode"))
	
	
	strStockOnhandBibit		=	" select fcba" & _
	                        "        , storecode" & _
	                        "        , store_api.get_fcname(storecode, fcba) storename" & _
	                        "        , stockcode" & _
	                        "        , stock_api.get_stockname(stockcode, fcba) stockname" & _
	                        "        , nurserycode" & _
	                        "        , nursery_api.get_description(nurserycode, fcba) nurseryname" & _
	                        "        , quantity" & _
	                        " from stockonhandnursery" & _
							" where fcba||stockcode in (select fcba||fccode from stock where fcba||fccode in (select fcba||stockcode from nursery))" & _
							"       and fcba = '" & Session("Fcba") & "'"

	if  len(trim(v_txtNurseryCode)) > 0 then
		strStockOnhandBibit	=	strStockOnhandBibit & " and nurserycode = '"& v_txtNurseryCode & "'"
	end if
	
	strStockOnhandBibit	=	strStockOnhandBibit & " order by stockcode, storecode, substr(nurserycode, 3, 2), substr(nurserycode, 1, 2)"

	Set rsStockOnhandBibit = Server.CreateObject("ADODB.Recordset")
	rsStockOnhandBibit.Open strStockOnhandBibit, connstr, adOpenStatic 
%>


<html>

<head>
	<title>ePlaS Reporting - Stock On Hand Bibit</title>
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
						<td><font face="verdana" size="1">Kode Batch</font></td>					
						<td align="center">
							<input type="text" maxlength="11" name="txtNurseryCode" size="11" value="<%=v_txtNurseryCode%>" title="Input Stock code">
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
					<th colspan="7" style="background-color:#FFFFFF; border:none;">  
						<div id="header-wrapper">
							<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
								<h4>PT. <%=Session("COMPANY")%></h4>
								<h5><%=Session("BA")%></h5>
							</div>
							<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
								<h2>LAPORAN STOCK ON HAND BIBITAN</h2>
								<% If Len(Trim(v_txtNurseryCode)) > 0 Then %>
								<h4>Nursery Code : <%=v_txtNurseryCode%></h4>
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
					<th>NURSERY CODE</th>
					<th>NURSERY NAME</th>
					<th>QUANTITY</th>
				</tr>
</thead>
<tbody>
<%if rsStockOnhandBibit.eof and rsStockOnhandBibit.eof then%>
	<tr>
		<td colspan="7" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		nTotal = 0
		while not rsStockOnhandBibit.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="left"><%=rsStockOnhandBibit("storecode")%></td>
				<td align="left"><%=rsStockOnhandBibit("storename")%></td>
				<td align="left"><%=rsStockOnhandBibit("stockcode")%></td>
				<td align="left"><%=rsStockOnhandBibit("stockname")%></td>
				<td align="left"><%=rsStockOnhandBibit("nurserycode")%></td>
				<td align="left"><%=rsStockOnhandBibit("nurseryname")%></td>
				<td align="right"><%=FormatNumber(rsStockOnhandBibit("QUANTITY"), 0)%></td>
				
			</tr>
<%
			NO = NO + 1
			nTotal = nTotal + Cdbl(rsStockOnhandBibit("QUANTITY"))
			rsStockOnhandBibit.MoveNext
		Wend
	rsStockOnhandBibit.Close
	set rsStockOnhandBibit = Nothing
%>
			<tr style="background-color:#F4F4A0;">	
				<td align="center" colspan="6"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FormatNumber(nTotal, 0)%></strong></td>
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