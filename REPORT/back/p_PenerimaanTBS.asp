<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtDateFrom, v_txtDateTo,v_txtSupplier
Dim ScriptName, NO, cClass
Dim nBRUTO, nSORTASI, nNETTO


v_txtDateFrom	=	Trim(Request("txtDateFrom"))
v_txtDateTo		=	Trim(Request("txtDateTo"))
v_txtSupplier			=	Trim(Request("txtSupplier"))

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

Dim rsTerimaTBS, strTerimaTBS

strTerimaTBS	=	" SELECT FCBA," & _
					"        TANGGAL," & _
					"        KODESUPPLIER," & _
					"        NAMASUPPLIER," & _
					"        NOKENDARAAN," & _
					"        BRUTO," & _
					"        SORTASI," & _
					"        NETTO" & _
					"   FROM (  SELECT FCBA," & _
					"                  TANGGAL," & _
					"                  KODESUPPLIER," & _
					"                  NAMASUPPLIER," & _
					"                  NOKENDARAAN," & _
					"                  SUM (WEIGHTIN) - SUM (WEIGHTOUT) BRUTO," & _
					"                  SUM (SORTASI) SORTASI," & _
					"                  (SUM (WEIGHTIN) - SUM (WEIGHTOUT)) - SUM (SORTASI) NETTO" & _
					"             FROM (SELECT WBT.FCBA," & _
					"                          WBT.DATEIN TANGGAL," & _
					"                          WBT.CUST_CUSTOMERCODE KODESUPPLIER," & _
					"                          CS.DESCRIPTION NAMASUPPLIER," & _
					"                          WBT.VEHICLENO NOKENDARAAN," & _
					"                          NVL (WBT.WEIGHTIN, 0) WEIGHTIN," & _
					"                          NVL (WBT.WEIGHTOUT, 0) WEIGHTOUT," & _
					"                          NVL (GR.WEIGHT, 0) SORTASI" & _
					"                     FROM WEIGHBRIDGETICKET WBT" & _
					"                          LEFT JOIN GRADINGRECORD GR" & _
					"                             ON (WBT.CHITNO = GR.WBT_CHITNO" & _
					"                                 AND WBT.FCBA = GR.FCBA)" & _
					"                          INNER JOIN CUSTOMER CS" & _
					"                             ON (WBT.FCBA = CS.FCBA" & _
					"                                 AND WBT.CUST_CUSTOMERCODE = CS.CUSTOMERCODE)" & _
					"                    WHERE PRODUCT_PRODUCTCODE = 'FFB')" & _
					"         GROUP BY FCBA," & _
					"                  TANGGAL," & _
					"                  KODESUPPLIER," & _
					"                  NAMASUPPLIER," & _
					"                  NOKENDARAAN)" & _
					" WHERE FCBA = '" & Session("FCBA") & "'"

if  len(trim(v_txtDateFrom)) > 0 then
	strTerimaTBS = strTerimaTBS & " AND TANGGAL >=	to_Date('" & v_txtDateFrom & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDateTo)) > 0 then
	strTerimaTBS = strTerimaTBS &  " AND TANGGAL <= to_Date('" & v_txtDateTo & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtSupplier)) > 0 then
	strTerimaTBS = strTerimaTBS &  " AND KODESUPPLIER = '" & v_txtSupplier & "'"
end if


Set rsTerimaTBS = Server.CreateObject("ADODB.Recordset")
rsTerimaTBS.Open strTerimaTBS, connstr, adOpenStatic

%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Penerimaan TBS</title>
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
		<h2>LAPORAN PENERIMAAN TBS</h3>
		<br/>
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">From</font></td>
		<td align="center">
		<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
			<input type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtDateFrom');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td><font face="verdana" size="1">To</font></td>
		<td align="center">
		<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
			<input type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtDateTo');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td><font face="verdana" size="1">Supplier</font></td>
		<td><select name="txtSupplier" ><option value=""></option>
		<%
		Dim rsSupplier, strSupplier
		strSupplier	=	" SELECT CUSTOMERCODE, DESCRIPTION FROM CUSTOMER WHERE SUPPLIER = 'Y' AND FCBA = '" & Session("FCBA") & "'"
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

<table border="0" width="100%" id="product-table">
	<tr>
		<th style="color:#FFFFFF" rowspan="2">NO</th>
		<th style="color:#FFFFFF" rowspan="2">TANGGAL</th>
		<th style="color:#FFFFFF" rowspan="2">KODE SUPPLIER</th>
		<th style="color:#FFFFFF" rowspan="2">NAMA SUPPLIER</th>
		<th style="color:#FFFFFF" rowspan="2">NO KENDARAAN</th>
		<th style="color:#FFFFFF" colspan="3">QUANTITY (KG)</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF" >BRUTO</th>
		<th style="color:#FFFFFF" >SORTASI</th>
		<th style="color:#FFFFFF" >NETTO</th>
	</tr>
<%
if rsTerimaTBS.bof and rsTerimaTBS.eof then
%>
	<tr><td colspan="8">NO DATA AVAILABLE</td></tr>
<%else
		NO = 1
			nBRUTO =  0
			nSORTASI =  0
			nNETTO =  0
		Do while not rsTerimaTBS.eof
			if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=NO%></td>
		<td align="right"><%=FbpNum2(rsTerimaTBS("TANGGAL"))%></td>
		<td align="left"><%=rsTerimaTBS("KODESUPPLIER")%></td>
		<td align="left"><%=rsTerimaTBS("NAMASUPPLIER")%></td>
		<td align="left"><%=rsTerimaTBS("NOKENDARAAN")%></td>
		<td align="right"><%=FormatNumber(rsTerimaTBS("BRUTO"),2)%></td>
		<td align="right"><%=FormatNumber(rsTerimaTBS("SORTASI"),2)%></td>
		<td align="right"><%=FormatNumber(rsTerimaTBS("NETTO"),2)%></td>
	</tr>
<%
			NO = NO + 1
			nBRUTO =  nBRUTO + Cdbl(rsTerimaTBS("BRUTO"))
			nSORTASI =  nSORTASI + Cdbl(rsTerimaTBS("SORTASI"))
			nNETTO =  nNETTO + Cdbl(rsTerimaTBS("NETTO"))

			rsTerimaTBS.MoveNext
		loop
end if
rsTerimaTBS.Close
set rsTerimaTBS = Nothing
%>

			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="5"><strong>TOTAL</td>
				<td align="right"><strong><%=FormatNumber(nBRUTO,2)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nSORTASI,2)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nNETTO,2)%></strong></td>
			</tr>
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
