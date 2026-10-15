<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtDateFrom, v_txtDateTo, v_txtSupplier, v_txtProduct
Dim ScriptName, NO, cClass
Dim nQTYDELIVERY


v_txtDateFrom	=	Trim(Request("txtDateFrom"))
v_txtDateTo		=	Trim(Request("txtDateTo"))
v_txtSupplier	=	Trim(Request("txtSupplier"))
v_txtProduct	=	Trim(Request("txtProduct"))

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

Dim rsPengirimanCPOKER, strPengirimanCPOKER

strPengirimanCPOKER	=	" SELECT FCBA," & _
						"        TANGGAL," & _
						"        KODEPRODUK," & _
						"        NAMAPRODUK," & _
						"        NOKONTRAK," & _
						"        KODESUPPLIER," & _
						"        NAMASUPPLIER," & _
						"        TRANSPORTCO," & _
						"        NODELIVERYNOTE," & _
						"        NOKENDARAAN," & _
						"        SUPIR," & _
						"        QTYDELIVERY," & _
						"        FFA," & _
						"        DIRT," & _
						"        MOISTURE" & _
						"   FROM (  SELECT DN.FCBA," & _
						"                  DN.DNOTE_DATE TANGGAL," & _
						"                  WBT.PRODUCT_PRODUCTCODE KODEPRODUK," & _
						"                  PR.PRODUCTDESCRIPTION NAMAPRODUK," & _
						"                  DN.CONTRACTCODE NOKONTRAK," & _
						"                  WBT.CUST_CUSTOMERCODE KODESUPPLIER," & _
						"                  CS.DESCRIPTION NAMASUPPLIER," & _
						"                  DN.TRANSPORTCO," & _
						"                  DN.DELIVERYNOTECODE NODELIVERYNOTE," & _
						"                  WBT.VEHICLENO NOKENDARAAN," & _
						"                  DN.DRIVERNAME SUPIR," & _
						"                  NVL (QUANTITY, 0) QTYDELIVERY," & _
						"                  NVL (DN.FFA, 0) FFA," & _
						"                  NVL (DN.DIRT, 0) DIRT," & _
						"                  NVL (DN.MOISTURE, 0) MOISTURE" & _
						"             FROM DELIVERYNOTE DN" & _
						"                  LEFT JOIN WEIGHBRIDGETICKET WBT" & _
						"                     ON (DN.FCBA = WBT.FCBA AND DN.WBT_CHITNO = WBT.CHITNO)" & _
						"                  INNER JOIN CUSTOMER CS" & _
						"                     ON (WBT.FCBA = CS.FCBA" & _
						"                         AND WBT.CUST_CUSTOMERCODE = CS.CUSTOMERCODE)" & _
						"                  INNER JOIN PRODUCT PR" & _
						"                     ON (WBT.FCBA = PR.FCBA" & _
						"                         AND WBT.PRODUCT_PRODUCTCODE = PR.PRODUCTCODE)" & _
						"         ORDER BY FCBA," & _
						"                  PRODUCT_PRODUCTCODE," & _
						"                  DNOTE_DATE," & _
						"                  CONTRACTCODE," & _
						"                  CUST_CUSTOMERCODE ASC)" & _
						" WHERE FCBA = '" & Session("FCBA") & "'"

if  len(trim(v_txtDateFrom)) > 0 then
	strPengirimanCPOKER = strPengirimanCPOKER & " AND TANGGAL >= TO_DATE('" & v_txtDateFrom & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDateTo)) > 0 then
	strPengirimanCPOKER = strPengirimanCPOKER &  " AND TANGGAL <= TO_DATE('" & v_txtDateTo & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtProduct)) > 0 then
	strPengirimanCPOKER = strPengirimanCPOKER &  " AND KODEPRODUK = '" & v_txtProduct & "'"
end if
if  len(trim(v_txtSupplier)) > 0 then
	strPengirimanCPOKER = strPengirimanCPOKER &  " AND KODESUPPLIER = '" & v_txtSupplier & "'"
end if
'response.write strPengirimanCPOKER
Set rsPengirimanCPOKER = Server.CreateObject("ADODB.Recordset")
rsPengirimanCPOKER.Open strPengirimanCPOKER, connstr, adOpenStatic

%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Pengiriman</title>
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
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
     <script>
	$(function() {
	$( ".datepicker" ).datepicker();
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
</script>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
    <table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">From</font></td>
		<td align="center">
		<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			
		</td>
		<td><font face="verdana" size="1">To</font></td>
		<td align="center">
		<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			
		</td>
		</td>
		<td><font face="verdana" size="1">Product</font></td>
		<td><select name="txtProduct" ><option value=""></option>
		<%
		Dim rsProductCode, strProductCode
		strProductCode	=	" SELECT PRODUCTCODE, PRODUCTDESCRIPTION FROM PRODUCT WHERE FCBA = '" & Session("FCBA") & "'"
		strProductCode	=	strProductCode & " AND PRODUCTCODE NOT IN ('FFB') ORDER BY PRODUCTCODE ASC"
		
		Set rsProductCode = Server.CreateObject("ADODB.Recordset")
		rsProductCode.Open strProductCode, connstr, adOpenStatic 
		
		if not rsProductCode.eof then
		Do While Not rsProductCode.eof
		if rsProductCode("PRODUCTCODE") = v_txtProduct then%>
		        <option value="<%=rsProductCode("PRODUCTCODE")%>" Selected><%=rsProductCode("PRODUCTCODE")%>-<%=rsProductCode("PRODUCTDESCRIPTION")%></option> 
		     <%Else%>
		        <option value="<%=rsProductCode("PRODUCTCODE")%>"><%=rsProductCode("PRODUCTCODE")%>-<%=rsProductCode("PRODUCTDESCRIPTION")%></option>  
		    <%End if
			rsProductCode.movenext
		loop
		end if
		rsProductCode.Close
		set rsProductCode = Nothing
		%>
		</select></td>
		<td><font face="verdana" size="1">Supplier</font></td>
		<td><select name="txtSupplier" ><option value=""></option>
		<%
		Dim rsSupplier, strSupplier
		strSupplier	=	" SELECT CUSTOMERCODE, DESCRIPTION FROM CUSTOMER WHERE SUPPLIER = 'N' AND FCBA = '" & Session("FCBA") & "'"
		strSupplier	=	strSupplier & " ORDER BY CUSTOMERCODE ASC"
		
		Set rsSupplier = Server.CreateObject("ADODB.Recordset")
		rsSupplier.Open strSupplier, connstr, adOpenStatic 
		
		if not rsSupplier.eof then
		Do While Not rsSupplier.eof
		if rsSupplier("CUSTOMERCODE") = v_txtSupplier then%>
		        <option value="<%=rsSupplier("CUSTOMERCODE")%>" Selected><%=rsSupplier("CUSTOMERCODE")%>-<%=rsSupplier("DESCRIPTION")%></option> 
		     <%Else%>
		        <option value="<%=rsSupplier("CUSTOMERCODE")%>"><%=rsSupplier("CUSTOMERCODE")%>-<%=rsSupplier("DESCRIPTION")%></option>  
		    <%End if
			rsSupplier.movenext
		loop
		end if
		rsSupplier.Close
		set rsSupplier = Nothing
		%>
		</select></td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;">
		<h2>LAPORAN PENGIRIMAN</h2>
		<br/>
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>
	</div>
	</div>
       </th>
  </tr>
	<tr>
		<th style="color:#FFFFFF" rowspan="2">NO</th>
		<th style="color:#FFFFFF" rowspan="2">TANGGAL</th>
		<th style="color:#FFFFFF" rowspan="2">NO KONTRAK</th>
		<th style="color:#FFFFFF" rowspan="2">KODE SUPPLIER</th>
		<th style="color:#FFFFFF" rowspan="2">NAMA SUPPLIER</th>
		<th style="color:#FFFFFF" rowspan="2">TRANSPORTCO</th>
		<th style="color:#FFFFFF" rowspan="2">NO DELIVERYNOTE</th>
		<th style="color:#FFFFFF" rowspan="2">NO KENDARAAN</th>
		<th style="color:#FFFFFF" rowspan="2">SUPIR</th>
		<th style="color:#FFFFFF" rowspan="2">QUANTITY</th>
		<th style="color:#FFFFFF" colspan="3">KUALITAS</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF">FFA</th>
		<th style="color:#FFFFFF">DIRT</th>
		<th style="color:#FFFFFF">MOSTURE</th>
	</tr>
  </thead>
  <tbody>
<%
if rsPengirimanCPOKER.bof and rsPengirimanCPOKER.eof then
%>
	<tr><td colspan="13">NO DATA AVAILABLE</td></tr>
<%else
		NO = 1
		nQTYDELIVERY = 0
		Do while not rsPengirimanCPOKER.eof
			if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=NO%></td>
		<td align="right"><%=FbpNum2(rsPengirimanCPOKER("TANGGAL"))%></td>
		<td align="right"><%=FbpNum2(rsPengirimanCPOKER("NOKONTRAK"))%></td>
		<td align="left"><%=rsPengirimanCPOKER("KODESUPPLIER")%></td>
		<td align="left"><%=rsPengirimanCPOKER("NAMASUPPLIER")%></td>
		<td align="left"><%=rsPengirimanCPOKER("TRANSPORTCO")%></td>
		<td align="left"><%=rsPengirimanCPOKER("NODELIVERYNOTE")%></td>
		<td align="left"><%=rsPengirimanCPOKER("NOKENDARAAN")%></td>
		<td align="left"><%=rsPengirimanCPOKER("SUPIR")%></td>
		<td align="right"><%=FormatNumber(rsPengirimanCPOKER("QTYDELIVERY"),2)%></td>
		<td align="right"><%=FormatNumber(rsPengirimanCPOKER("FFA"),2)%></td>
		<td align="right"><%=FormatNumber(rsPengirimanCPOKER("DIRT"),2)%></td>
		<td align="right"><%=FormatNumber(rsPengirimanCPOKER("MOISTURE"),2)%></td>
	</tr>
<%
			NO = NO + 1
			nQTYDELIVERY = nQTYDELIVERY + Cdbl(rsPengirimanCPOKER("QTYDELIVERY"))
			rsPengirimanCPOKER.MoveNext
		loop
end if
rsPengirimanCPOKER.Close
set rsPengirimanCPOKER = Nothing
%>

			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="9"><strong>TOTAL</td>
				<td align="right"><strong><%=FormatNumber(nQTYDELIVERY,2)%></strong></td>
				<td align="right" colspan="3"><strong>&nbsp;</strong></td>
			</tr>
            </tbody>
</table>

   <table width="90%" class="data">
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
