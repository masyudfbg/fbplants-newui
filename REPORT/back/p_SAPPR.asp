<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtDateFrom, v_txtDateTo

v_txtDateFrom		=	trim(Request("txtDateFrom"))
v_txtDateTo			=	trim(Request("txtDateTo"))

ScriptName			=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - PR SAP</title>
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
		<table id="form-table" class="centered-margin">
			<tr>
				<form name="frmEntry" method="post" action="<%=ScriptName%>">
				<td><font face="verdana" size="1">Dari</font></td>
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
				<td><font face="verdana" size="1">Sampai</font></td>
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
				<td><input type="submit" value="View" /></td>
				</form>
			</tr>
		</table>
		<div id="header">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered">
			<h2>LAPORAN PR SAP</h2>			 
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>	
		</div>
	</div>

<table border="0" width="100%" id="product-table">
     <tr>
		<th style="color:#fff;">DOCUMENT TYPE</th>		
		<th style="color:#fff;">L</th>
		<th style="color:#fff;">ITEMNUMBER</th>
		<th style="color:#fff;">MATERIAL CODE</th>
		<th style="color:#fff;">PR QUANTITY</th>
		<th style="color:#fff;">MATERIAL UOM</th>
		<th style="color:#fff;">PLANT</th>
		<th style="color:#fff;">STORAGE LOCATION</th>
		<th style="color:#fff;">MATERIAL DESCRIPTION</th>
		<th style="color:#fff;">DELIVERY DATE</th>
		<th style="color:#fff;">REQUISITIONE</th>
		<th style="color:#fff;">PURCHASING GROUP</th>
		<th style="color:#fff;">PR DATE</th>
		<th style="color:#fff;">VALUATION PRICE</th>
		<th style="color:#fff;">CURRENCY</th>
		<th style="color:#fff;">PRICE UNIT</th>
		<th style="color:#fff;">PURCHASING ORGANIZATION</th>
		<th style="color:#fff;">ITEM TEXT 1</th>
		<th style="color:#fff;">ITEM TEXT 2</th>
		<th style="color:#fff;">ITEM TEXT 3</th>
		<th style="color:#fff;">ITEM TEXT 4</th>
		<th style="color:#fff;">ITEM TEXT 5</th>
		<th style="color:#fff;">ITEM NOTE 1</th>
		<th style="color:#fff;">ITEM NOTE 2</th>
		<th style="color:#fff;">ITEM NOTE 3</th>
		<th style="color:#fff;">ITEM NOTE 4</th>
		<th style="color:#fff;">ITEM NOTE 5</th>
		<th style="color:#fff;">ITEM DELIVERY 1</th>
		<th style="color:#fff;">ITEM DELIVERY 2</th>
		<th style="color:#fff;">ITEM DELIVERY 3</th>
		<th style="color:#fff;">ITEM DELIVERY 4</th>
		<th style="color:#fff;">ITEM DELIVERY 5</th>
		<th style="color:#fff;">ITEM DELIVERY 6</th>
		<th style="color:#fff;">NO AFCE</th>
		<th style="color:#fff;">PROJECTCODE</th>
	</tr>
	<%
	Dim rsSAPPR, strSAPPR , noline
	
	strSAPPR	=	" SELECT Purchase_Report_Api.GetDocumentType (a.fcba) document_type," & _
					"        b.theNumber," & _
					"        a.lineid," & _
					"        a.itemcode," & _
					"        a.approvedqty," & _
					"        a.unitofmeasurement," & _
					"        Purchase_Report_Api.GetPlant (a.fcba, b.managedid) Plant," & _
					"        Purchase_Report_Api.GetStorageLocation (a.fcba) Storage_Location," & _
					"        TO_CHAR (a.wanteddeliverydate, 'DD.MM.YYYY') delivery_date," & _
					"        B.REQUESTORID," & _
					"        B.PURCHASEGROUPID," & _
					"        TO_CHAR (b.prdate, 'DD.MM.YYYY') prdate," & _
					"        Purchase_Report_Api.GetPurchasingOrganization (a.fcba, b.managedid)" & _
					"           Purchasing_Organization," & _
					"        a.description," & _
					"        a.prcode" & _
					"   FROM purchaserequestdetails a," & _
					"        (SELECT ROWNUM theNumber," & _
					"                prcode," & _
					"                fcba," & _
					"                prdate," & _
					"                managedid," & _
					"                requestorid," & _
					"                purchasegroupid," & _
					"                rowstateid" & _
					"           FROM purchaserequest" & _
					"          WHERE rowstateid = 'Approved') b" & _
					"  WHERE a.prcode = b.prcode AND a.fcba = b.fcba" & _
					"  AND a.FCBA = '" & Session("FCBA") & "'"
if  len(trim(v_txtDateFrom)) > 0 then
	strSAPPR = strSAPPR & " AND b.prdate >=	to_Date('" & v_txtDateFrom & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDateTo)) > 0 then
	strSAPPR = strSAPPR &  " AND b.prdate <= to_Date('" & v_txtDateTo & "','MM/DD/YYYY')"
end if
	 		
	Set rsSAPPR = Server.CreateObject("ADODB.Recordset")
	rsSAPPR.Open strSAPPR, connstr, adOpenStatic
    
	if rsSAPPR.bof and rsSAPPR.eof then
	%>
	<tr>
			<td align="left" colspan="35"><strong>DATA TIDAK DITEMUKAN</strong></td>
	</tr>	
	<%
	elseif Not rsSAPPR.eof then		
		while Not rsSAPPR.eof  %>
		           
		  <tr>
			<td align="left"><%=rsSAPPR("document_type")%></td>
			<td align="right"><%=rsSAPPR("theNumber")%></td>
			<td align="right"><%=rsSAPPR("lineid")%></td>
			<td align="left"><%=rsSAPPR("itemcode")%></td>
			<td align="right"><%=formatnumber( cdbl(rsSAPPR("approvedqty")),0)%></td>
			<td align="left"><%=rsSAPPR("unitofmeasurement")%></td>
			<td align="left"><%=rsSAPPR("Plant")%></td>
			<td align="left"><%=rsSAPPR("Storage_Location")%></td>
			<td align="left">&nbsp;</td>
			<td align="right"><%=rsSAPPR("delivery_date")%></td>
			<td align="left"><%=rsSAPPR("REQUESTORID")%></td>
			<td align="left"><%=rsSAPPR("PURCHASEGROUPID")%></td>
			<td align="right"><%=rsSAPPR("prdate")%></td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left"><%=rsSAPPR("Purchasing_Organization")%></td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left"><%=rsSAPPR("prcode")%></td>
			<td align="left"><%=rsSAPPR("description")%></td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
			<td align="left">&nbsp;</td>
		  </tr>	
		      
<%
			
			rsSAPPR.MoveNext
		wend
	end if
rsSAPPR.Close
set rsSAPPR = Nothing
%>

</table>
<!--#include file="bottom.Asp"-->