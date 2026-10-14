<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, v_txtContractorCode, v_txtNoSPK
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
v_txtContractorCode=	trim(request("txtContractorCode"))
v_txtNoSPK=	trim(request("txtNoSPK"))
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Outstanding Kontrak</title>
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
body{font-family:'Arial Narrow';font-size:12px;}
.printhide{display:none;}
}
</style>

</head>

<body>
	<div id="wrapper">
	<div id="header-wrapper">
	<div class="printhide">
	<!--#include file="menutop.asp"-->
	<table id="form-table" class="centered-margin">
		<form name="frmBTL" method="post" action="<%=ScriptName%>">
			<tr>
				<td><font face="verdana" size="1">TAHUN</font></td>
				<td>
<% 
    MM=Request.Form("vMonth")
    yyyy=Request.Form("vYear")

    mon = Month(Session("v_Costbook"))
    If len(MM)=0 then MM=mon

    Yr = Year(Session("v_Costbook"))
    MinYr = Yr - 2
    MaxYr = Yr  + 2
    If len(yyyy)= 0 then yyyy = Yr
%>
				<select name="vYear">
<%
    For i = MinYr to MaxYr
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
<%
    Next
%>
				</select>
				</td>
				<td><font face="verdana" size="1">BULAN</font></td>
				<td>
					<select name="vMonth">
<%
    For i = 1 to 12
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=Bulan(i)%></option>
<%
    Next
%>
					</select>
				</td>

				<td><font face="verdana" size="1">CUSTOMER</font></td><td><select name="txtContractorCode" ><option value=""></option>
				<%
				Dim rsCustomer, strCustomer
				strCustomer	=	" SELECT CUSTOMERCODE, DESCRIPTION FROM CUSTOMER WHERE SUPPLIER = 'N' AND FCBA = '" & Session("FCBA") & "' ORDER BY CUSTOMERCODE ASC"
				
				Set rsCustomer = Server.CreateObject("ADODB.Recordset")
				rsCustomer.Open strCustomer, connstr, adOpenStatic 
				
				if not rsCustomer.eof then
				Do While Not rsCustomer.eof
				if rsCustomer("CUSTOMERCODE") = v_txtContractorCode then%>
						<option value="<%=rsCustomer("CUSTOMERCODE")%>" Selected><%=rsCustomer("CUSTOMERCODE")%> - <%=rsCustomer("DESCRIPTION")%></option> 
					 <%Else%>
						<option value="<%=rsCustomer("CUSTOMERCODE")%>"><%=rsCustomer("CUSTOMERCODE")%> - <%=rsCustomer("DESCRIPTION")%></option>  
					<%End if
					rsCustomer.movenext
				loop
				end if
				rsCustomer.Close
				set rsCustomer = Nothing
				%>
				</select></td>
				<td><font face="verdana" size="1">NO. SPK</font></td>
				<td><input type="text" maxlength="50" name="txtNoSPK" size="30" value="<%=v_txtNoSPK%>" title="Input No SPK"></td>
				<td><input type="submit" value="View" /></td>
			</tr>
		</form>
	</table>
</div>
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h3>LAPORAN OUTSTANDING KONTRAK</h3>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(mm)%>&nbsp;<%=yyyy%></h4>
	</div>
	</div>
<table border="0" width="100%" id="product-table">
	<tr>
		<th style="color:#FFFFFF" rowspan="3">NO. SPK</th>
		<th style="color:#FFFFFF" rowspan="3">KODE PEMBELI</th>
		<th style="color:#FFFFFF" rowspan="3">NAMA PEMBELI</th>
		<th style="color:#FFFFFF" rowspan="2" colspan="2">JANGKA WAKTU KONTRAK</th>
		<th style="color:#FFFFFF" rowspan="3">NAMA PRODUK</th>
		<th style="color:#FFFFFF" colspan="4">QUANTITY</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF" rowspan="2">KONTRAK</th>
		<th style="color:#FFFFFF" colspan="2">KIRIM</th>
		<th style="color:#FFFFFF" rowspan="2">SISA</th>		
	</tr>
	<tr>
		<th style="color:#FFFFFF">MULAI</th>
		<th style="color:#FFFFFF">BERAKHIR</th>	
		<th style="color:#FFFFFF">BI</th>
		<th style="color:#FFFFFF">SBI</th>	
	</tr>
<%
Dim rsOutstandingDeliveryContract, strOutstandingDeliveryContract, rowno
rowno		=	0
strOutstandingDeliveryContract		=	" SELECT theData.contractcode," & _
										" theMaster.cust_customercode," & _
										" customer.description customername," & _
										" theMaster.STARTDATE," & _
										" theMaster.ENDDATE," & _
										" theMaster.Product_ProductCode," & _
										" theData.contract_qty," & _
										" theData.bi_qty," & _
										" theData.sbi_qty," & _
										" (theData.contract_qty - theData.sbi_qty) remain_qty" & _
										" FROM ( SELECT contractcode," & _
										" SUM (contract_qty) contract_qty," & _
										" SUM (bi_qty) bi_qty," & _
										" SUM (sbi_qty) sbi_qty," & _
										" fcba" & _
										" FROM (SELECT a.contractcode," & _
										" a.totalweight contract_qty," & _
										" 0 bi_qty," & _
										" 0 sbi_qty," & _
										" a.fcba" & _
										" FROM DeliveryContract a" & _
										" WHERE a.fcba = '" & Session("FCBA") & "'" & _
										" UNION" & _
										" SELECT a.contractcode," & _
										" 0 contract_qty," & _
										" SUM (a.quantity) bi_qty," & _
										" 0 sbi_qty," & _
										" a.fcba" & _
										" FROM DeliveryNote a" & _
										" WHERE a.fcba = '" & Session("FCBA") & "'" & _
										" AND TO_NUMBER (TO_CHAR (dnote_date, 'YYYY')) =" & yyyy & _
										" AND TO_NUMBER (TO_CHAR (dnote_date, 'MM')) =" & MM & _
										" GROUP BY a.contractcode, a.fcba" & _
										" UNION" & _
										" SELECT a.contractcode," & _
										" 0 contract_qty," & _
										" 0 bi_qty," & _
										" SUM (a.quantity) sbi_qty," & _
										" a.fcba" & _
										" FROM DeliveryNote a" & _
										" WHERE a.fcba = '" & Session("FCBA") & "'" & _
										" AND TO_NUMBER (TO_CHAR (dnote_date, 'YYYY')) <=" & yyyy & _
										" AND TO_NUMBER (TO_CHAR (dnote_date, 'MM')) <=" & MM & _
										" GROUP BY a.contractcode, a.fcba)" & _
										" GROUP BY contractcode, fcba) theData," & _
										" DeliveryContract theMaster," & _
										" customer" & _
										" WHERE theData.contractcode = theMaster.contractcode" & _
										" AND theMaster.cust_customercode = customer.CustomerCode" & _
										" AND theData.fcba = theMaster.fcba" & _
										" AND theMaster.fcba = customer.fcba"


if  LEN(TRIM(v_txtContractorCode)) > 0 then
	strOutstandingDeliveryContract	=	strOutstandingDeliveryContract & " AND theMaster.cust_customercode  = '" & v_txtContractorCode & "'"
end if
if  LEN(TRIM(v_txtNoSPK)) > 0 then
	strOutstandingDeliveryContract	=	strOutstandingDeliveryContract & " AND theData.contractcode  = '" & v_txtNoSPK & "'"
end if

Set rsOutstandingDeliveryContract = Server.CreateObject("ADODB.Recordset")
rsOutstandingDeliveryContract.Open strOutstandingDeliveryContract, connstr, adOpenStatic
if Not rsOutstandingDeliveryContract.eof then

	While Not rsOutstandingDeliveryContract.eof
			rowno	=	rowno	+	1
%>
	<tr>
		<td align="left"><%=rsOutstandingDeliveryContract("contractcode")%></td>
		<td align="center"><%=rsOutstandingDeliveryContract("cust_customercode")%></td>
		<td align="left"><%=rsOutstandingDeliveryContract("customername")%></td>
		<td align="right"><%=rsOutstandingDeliveryContract("STARTDATE")%></td>
		<td align="right"><%=rsOutstandingDeliveryContract("ENDDATE")%></td>
		<td align="left"><%=rsOutstandingDeliveryContract("Product_ProductCode")%></td>
		<td align="right"><%=FormatNumber(rsOutstandingDeliveryContract("contract_qty"),0)%></td>
		<td align="right"><%=FormatNumber(rsOutstandingDeliveryContract("bi_qty"),0)%></td>
		<td align="right"><%=FormatNumber(rsOutstandingDeliveryContract("sbi_qty"),0)%></td>
		<td align="right"><%=FormatNumber(rsOutstandingDeliveryContract("remain_qty"),0)%></td>
	</tr>
<%

		rsOutstandingDeliveryContract.MoveNext
	Wend
end if
%>
	<!--
	<tr bgcolor="#F4F4A0">
		<td colspan="2" align="left"><strong>Total</strong></td>
		<td align="right"><strong></strong></td>		
		<td align="right"><strong></strong></td>	
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong></strong></td>		
		<td align="right"><strong></strong></td>	
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
	</tr>
	-->

</table>
   
</div>
</body>
</html>