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
	
	
	strStockOnHand		=	" SELECT * FROM(" & _
							" SELECT" & _
							" SOH.FCBA FCBA," & _
							" STR.FCNAME STORENAME," & _
							" SOH.STOCKCODE STOCKCODE," & _
							" STO.FCNAME STOCKNAME," & _
							" SOH.QUANTITY QUANTITY," & _
							" STO.UNITOFMEASUREMENT UOM" & _
							" FROM STOCKONHAND SOH, STOCK STO, STORE STR" & _
							" WHERE" & _
							" (SOH.STORECODE = STR.FCCODE AND SOH.FCBA = STR.FCBA)" & _
							" AND" & _
							" (SOH.STOCKCODE = STO.FCCODE AND SOH.FCBA = STO.FCBA)" & _
							" ORDER BY SOH.FCBA, SOH.STORECODE, SOH.STOCKCODE ASC)" & _
							" WHERE FCBA = '" & Session("Fcba") & "'"


if  len(trim(v_txtStockCode)) > 0 then
    strStockOnHand	=	strStockOnHand & "AND STOCKCODE = '"& v_txtStockCode &"'"
end if

	
	Set rsStockOnHand = Server.CreateObject("ADODB.Recordset")
	rsStockOnHand.Open strStockOnHand, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Stock On Hand</title>
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
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN STOCK ON HAND</h3>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>Kode Stock</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post">
		<td align="center">
			<input type="text" maxlength="11" name="txtStockCode" size="11" value="<%=v_txtStockCode%>" title="Input Stock code">
		</td>
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th>NO</th>
		<th>STORE NAME</th>
		<th>STOCK CODE</th>
		<th>STOCK NAME</th>
		<th>SATUAN</th>
		<th>QUANTITY</th>
	</tr>

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
				<td align="center"><%=NO%></td>
				<td align="left"><%=rsStockOnHand("STORENAME")%></td>
				<td align="left"><%=rsStockOnHand("STOCKCODE")%></td>
				<td align="left"><%=rsStockOnHand("STOCKNAME")%></td>
				<td align="left"><%=rsStockOnHand("UOM")%></td>
				<td align="right"><%=rsStockOnHand("QUANTITY")%></td>
				
			</tr>
<%
			NO = NO + 1
			nTotal = nTotal + Cdbl(rsStockOnHand("QUANTITY"))
			rsStockOnHand.MoveNext
		Wend
	rsStockOnHand.Close
	set rsStockOnHand = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="5">TOTAL</td>
				<td align="right"><%=nTotal%></td>
			</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->
