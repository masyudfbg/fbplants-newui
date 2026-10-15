<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtDATECREATED_From, v_txtDATECREATED_To
Dim ScriptName, NO, cClass
Dim nSALDOAWAL, nDEBET, nCREDIT, nSALDOAKHIR


v_txtDATECREATED_From	=	Trim(Request("Datecreated_From"))
v_txtDATECREATED_To		=	Trim(Request("Datecreated_To"))

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

Dim rsTerimaTBS, strTerimaTBS

strTerimaTBS	=	" select tanggal, no_kontrak, nvl(qty_contract,0) qty_contract , no_kendaraan, DELIVERYNOTECODE, nvl(quantity,0) quantity   from ( select a.dnote_date tanggal, a.contractcode no_kontrak,  " & _
					 " (select totalweight from deliverycontracT where fcba = a.fcba and contractcode = a.contractcode) qty_contract, " & _
					 " (  select distinct vehicleno from weighbridgeticket where fcba = a.fcba and chitno = a.wbt_chitno and product_productcode ='CPO' ) no_kendaraan , " & _
					 "  a.DELIVERYNOTECODE, a.quantity from deliverynote a )   "

if  len(trim(v_txtDATECREATED_From)) > 0 then
	strTerimaTBS = strTerimaTBS & " where TANGGAL >=	to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
	strTerimaTBS = strTerimaTBS &  " AND TANGGAL <= to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
end if
strTerimaTBS = strTerimaTBS & " order by no_kontrak, tanggal "


Set rsTerimaTBS = Server.CreateObject("ADODB.Recordset")
rsTerimaTBS.Open strTerimaTBS, connstr, adOpenStatic

%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - DELIVERY NOTE</title>
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
	<div id="wrapper" style="width:900px;margin:0 0;">
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN DELIVERY NOTE</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;</h4>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">From</font></td>
		<td><input type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=TextDate(date())%>" title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="/Images/show-calendar.gif" width="24" height="22" border="0" align="top"></a><br></td>
		<td>To</td>
	<td><input type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=TextDate(date())%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="/Images/show-calendar.gif" width="24" height="22" border="0" align="top"></a></td>
		</td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>

<table border="0" width="900px" id="product-table">
	<tr>
	    <th style="color:#FFFFFF" >No</th>
		<th style="color:#FFFFFF" >Tanggal</th>
		<th style="color:#FFFFFF" >No Kontrak</th>
		
		<th style="color:#FFFFFF" >Qty Kontrak</th>
		<th style="color:#FFFFFF" >No Kendaraan</th>
		
		<th style="color:#FFFFFF" >Delivery Note No</th>
		<th style="color:#FFFFFF" >Qty Delivery</th>
		
		<th style="color:#FFFFFF" >Sisa</th>
	</tr>

<%  dim contrak,qtycontract,sisa
if rsTerimaTBS.bof and rsTerimaTBS.eof then
%>
	<tr><td colspan="8">NO DATA AVAILABLE</td></tr>
<%end if
		NO = 1
		contrak = "XXXX"
		while not rsTerimaTBS.eof
			if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=NO%></td>
		<td align="center"><%=rsTerimaTBS("TANGGAL")%></td>
		<td align="left"><%=rsTerimaTBS("no_kontrak")%></td>
		<td align="center"><%=rsTerimaTBS("qty_contract")%></td>
		<td align="left"><%=rsTerimaTBS("NO_KENDARAAN")%></td>
		<td align="right"><%=rsTerimaTBS("DELIVERYNOTECODE")%></td>
		<td align="right"><%=FbpNum2(rsTerimaTBS("quantity"))%></td>
		<% if contrak <>  rsTerimaTBS("no_kontrak")  then
		      qtycontract = cdbl(rsTerimaTBS("qty_contract"))
			  sisa = 0
			  sisa = cdbl(qtycontract) - (sisa + cdbl(rsTerimaTBS("quantity")))
 			  
		   else 
		      
			  sisa = sisa - cdbl(rsTerimaTBS("quantity"))
		   end if  %>
		<td align="right"><%=FbpNum2(sisa)%></td>
	</tr>
<%
			NO = NO + 1
			' nSALDOAWAL =  nSALDOAWAL + Cdbl(rsTerimaTBS("SALDOAWAL"))
			' nDEBET =  nDEBET + Cdbl(rsTerimaTBS("DEBET"))
			' nCREDIT =  nCREDIT + Cdbl(rsTerimaTBS("CREDIT"))
			' nSALDOAKHIR =  nSALDOAKHIR+ Cdbl(rsTerimaTBS("SALDOAKHIR"))
			contrak = rsTerimaTBS("no_kontrak")
			rsTerimaTBS.MoveNext
		Wend
rsTerimaTBS.Close
set rsTerimaTBS = Nothing
%>

</table>

   <table width="850px" class="data">
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
