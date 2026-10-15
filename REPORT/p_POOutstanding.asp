<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtSupplierCode, v_txtPOCode

v_txtSupplierCode	=	trim(Request("txtSupplierCode"))
v_txtPOCode	=	trim(Request("txtPOCode"))
ScriptName			=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - PO Outstanding</title>
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
			<tr>
				<form name="frmAR" method="post" action="<%=ScriptName%>">
				<td><font face="verdana" size="1">SUPPLIER</font></td><td><select name="txtSupplierCode" ><option value=""></option>
				<%
				Dim rsSupplierCode, strSupplierCode
				strSupplierCode	=	" SELECT FCCODE, FCNAME FROM SUPPLIER WHERE ACTIVATION = 'Y' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
				
				Set rsSupplierCode = Server.CreateObject("ADODB.Recordset")
				rsSupplierCode.Open strSupplierCode, connstr, adOpenStatic 
				
				if not rsSupplierCode.eof then
				Do While Not rsSupplierCode.eof
				if rsSupplierCode("FCCODE") = v_txtSupplierCode then%>
						<option value="<%=rsSupplierCode("FCCODE")%>" Selected><%=rsSupplierCode("FCCODE")%>-<%=rsSupplierCode("FCNAME")%></option> 
					 <%Else%>
						<option value="<%=rsSupplierCode("FCCODE")%>"><%=rsSupplierCode("FCCODE")%>-<%=rsSupplierCode("FCNAME")%></option>  
					<%End if
					rsSupplierCode.movenext
				loop
				end if
				rsSupplierCode.Close
				set rsSupplierCode = Nothing
				%>
				</select></td>
				<td><font face="verdana" size="1">NO PO</font></td><td><input type="text" name="txtPOCode" value="<%=v_txtPOCode%>" /></td>
				<td><input type="submit" value="View" /></td>
				</form>
			</tr>
		</table>
		<div id="header">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered">
			<h2>LAPORAN PO OUTSTANDING</h2>			 
		</div>
	</div>

<table border="0" width="100%" id="product-table">
     <tr>
		<th rowspan="2" style="color:#fff;">NO PO</th>		
		<th rowspan="2" style="color:#fff;">TANGGAL</th>
		<th rowspan="2" style="color:#fff;">SUPPLIER</th>
		<th rowspan="2" style="color:#fff;">KODE BARANG</th>
		<th rowspan="2" style="color:#fff;">NAMA BARANG</th>
		<th colspan="3" style="color:#fff;">QUANTITY</th>
		<th rowspan="2" style="color:#fff;">LAMA</th>
		<th rowspan="2" style="color:#fff;">KETERANGAN</th>
	</tr>
	<tr>		
		<th style="color:#fff;">ORDER</th>
		<th style="color:#fff;">TERIMA</th>
		<th style="color:#fff;">SISA</th>	
	</tr>
	<%
	Dim rsPOOut, strPOOut , noline
	
	strPOOut	=	" SELECT FCBA," & _
					" POCODE," & _
					" PODATE," & _
					" SUPPLIERCODE," & _
					" SUPPLIERNAME," & _
					" STOCKCODE," & _
					" STOCKNAME," & _
					" REMARKS," & _
					" QUANTITYORDER," & _
					" AMOUNTORDER," & _
					" QUANTITYRECEIVE," & _
					" QUANTITYSISA," & _
					" LAMA" & _
					" FROM ( SELECT PO.FCBA FCBA," & _
					" PO.POCODE POCODE," & _
					" PO.PODATE PODATE," & _
					" PO.SUPPLIERCODE SUPPLIERCODE," & _
					" SP.FCNAME SUPPLIERNAME," & _
					" POD.ITEMCODE STOCKCODE," & _
					" ITEM.FCNAME STOCKNAME," & _
					" PO.REMARKS REMARKS," & _
					" SUM (POD.QUANTITY) QUANTITYORDER," & _
					" SUM (POD.AMOUNT) AMOUNTORDER," & _
					" SUM (POD.RECEIVED) QUANTITYRECEIVE," & _
					" SUM (POD.QUANTITY) - SUM (POD.RECEIVED) QUANTITYSISA," & _
					" SUM(CEIL(SYSDATE-PO.PODATE)) LAMA" & _
					" FROM PURCHASEORDER PO" & _
					" INNER JOIN PURCHASEORDERDETAILS POD" & _
					" ON (PO.FCBA = POD.FCBA AND PO.POCODE = POD.POCODE)" & _
					" INNER JOIN SUPPLIER SP" & _
					" ON (PO.FCBA = SP.FCBA AND PO.SUPPLIERCODE = SP.FCCODE)" & _
					" INNER JOIN STOCK ITEM" & _
					" ON (ITEM.FCBA = POD.FCBA AND ITEM.FCCODE = POD.ITEMCODE)" & _
					" GROUP BY PO.FCBA," & _
					" PO.POCODE," & _
					" PO.PODATE," & _
					" PO.SUPPLIERCODE," & _
					" SP.FCNAME," & _
					" POD.ITEMCODE," & _
					" ITEM.FCNAME," & _
					" PO.REMARKS" & _
					" ORDER BY PO.FCBA, PO.SUPPLIERCODE, PO.PODATE)" & _
					" WHERE QUANTITYSISA <> 0" & _
					" AND FCBA = '" & Session("FCBA") & "'"	 
					
if  len(trim(v_txtSupplierCode)) > 0 then
    strPOOut	=	strPOOut &	" AND SUPPLIERCODE = '" & v_txtSupplierCode & "'"
end if
if  len(trim(v_txtPOCode)) > 0 then
    strPOOut	=	strPOOut &	" AND POCODE = '" & v_txtPOCode & "'"
end if

	
	Set rsPOOut = Server.CreateObject("ADODB.Recordset")
	rsPOOut.Open strPOOut, connstr, adOpenStatic
    

	if Not rsPOOut.eof then		
		while Not rsPOOut.eof  %>
		           
		  <tr>
			<td align="left"><%=rsPOOut("POCODE")%></td>
			<td align="left"><%=rsPOOut("PODATE")%></td>
			<td align="left"><%=rsPOOut("SUPPLIERNAME")%></td>
			<td align="left"><%=rsPOOut("STOCKCODE")%></td>
			<td align="left"><%=rsPOOut("STOCKNAME")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("QUANTITYORDER")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("QUANTITYRECEIVE")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("QUANTITYSISA")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("LAMA")),2)%></td>
			<td align="left"><%=rsPOOut("REMARKS")%></td>
		  </tr>	
		      
<%
			
			rsPOOut.MoveNext
		wend
	end if
%>

</table>
<!--#include file="bottom.Asp"-->