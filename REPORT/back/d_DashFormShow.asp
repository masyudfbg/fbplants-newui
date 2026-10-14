<%
	response.buffer = True
	validatesession
%>
<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<html>

<head>
<title>iPS Reporting</title>
<script type="text/javascript" language="JavaScript1.2" src="../stlib.js" ></script>
<script language="JavaScript" src="../date-picker.js"></script> 
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

<table>
<tr>
	<td>
		<a href="Dashboard.Asp" title="Back to Report List"><img src="../images/latar.jpg" border="0"></a>
	</td>
</tr>
<tr><td>&nbsp;</td></tr>
<tr><td><p><font face="Verdana" size="1">Select Data Here :</font></p></td></tr>
</table>
<%if Request("txtKind") ="ReceivedVoucher" then%>
    <form action="d_ReceivedVoucher.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
     <table width="400">
    <tr><td colspan="2"> </td></tr>
	<tr><td width="100"><font face="Verdana" size="1">From </td>
	<td width="300"><input type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=TextDate(date())%>" title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><br></td>
	<tr><td>To</td>
	<td><input type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=TextDate(date())%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
	<tr><td><font face="Verdana" size="1">Voucher Code</font></td>
	<td><input type="text" size="15" maxlength="30" name="VOUCHERCODE"  style="background:#99FF99"></td>
	<tr><td><font face="Verdana" size="1">Bank</font></td>
	<td><select name="BANKCODE" >
		<%cmdCallCombo "Bank","","",""%>
	</select></td>
    <tr><td><font face="Verdana" size="1">TT</font></td>
	<td><select name="LOCATIONTYPE" >
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select></td>
	<tr><td><font face="Verdana" size="1">TC</font></td>
	<td><input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#99FF99"></td>
	<tr><td colspan="2"><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td></tr></table>
<%Elseif Request("txtKind") ="PaymentVoucher" then%>
    <form action="d_PaymentVoucher.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table width="400">
    <tr><td colspan="2"> </td></tr>
	<tr><td width="100"><font face="Verdana" size="1">From </td>
	<td width="300"><input type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=TextDate(date())%>" title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><br></td>
	<tr><td>To</td>
	<td><input type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=TextDate(date())%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
	<tr><td><font face="Verdana" size="1">Voucher Code</font></td>
	<td><input type="text" size="15" maxlength="30" name="VOUCHERCODE"  style="background:#99FF99"></td>
	<tr><td><font face="Verdana" size="1">Bank</font></td>
	<td><select name="BANKCODE" >
		<%cmdCallCombo "Bank","","",""%>
	</select></td>
	<!--'MODIFIED BY UKIE 07 NOV 2012 BY ADDING FILTER BY JOBCODE-->
	<tr>
		<td><font face="Verdana" size="1">Job</font></td>
		<td><select name="JOBCODE" ><%cmdCallCombo "Job","","",""%></select></td>
	</tr>	
    <tr><td><font face="Verdana" size="1">TT</font></td>
	<td><select name="LOCATIONTYPE" >
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select></td>
	<tr><td><font face="Verdana" size="1">TC</font></td>
	<td><input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#99FF99"></td>
	<tr><td colspan="2"><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td></tr></table>
<%Elseif Request("txtKind") ="ArusCash" then%>
    <form action="d_ArusCash.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table width="400">
    <tr><td colspan="2"> </td></tr>
	<tr><td width="100"><font face="Verdana" size="1">From </td>
	<td width="300"><input type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=TextDate(date())%>" title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><br></td>
	<tr><td>To</td>
	<td><input type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=TextDate(date())%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
	<tr><td><font face="Verdana" size="1">Bank</font></td>
	<td><select name="BANKCODE" >
		<%cmdCallCombo "Bank","","",""%>
	</select></td>
	<tr><td colspan="2"><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td></tr></table>
<%Elseif Request("txtKind") ="PayslipSummary" then%>
    <form action="d_PayslipSummary.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table width="400">
    <tr><td colspan="2"> </td></tr>
	<tr><td width="100"><font face="Verdana" size="1">From </td>
	<td width="300"><input type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=TextDate(date())%>" title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><br></td>
	<tr><td>To</td>
	<td><input type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=TextDate(date())%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
	<tr><td><font face="Verdana" size="1">Employee Code</font></td>
	<td><input type="text" size="15" maxlength="30" name="Empcode"  style="background:#99FF99"></td>
	<tr><td><font face="Verdana" size="1">Gang</font></td>
	<td><select name="GANGCODE" >
		<%cmdCallCombo "GANG","","",""%>
	</select></td>
	<tr><td colspan="2"><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td></tr></table>
<%Elseif Request("txtKind") ="Payslip" then%>
    <form action="d_PayslipDetails.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table width="600">
    <tr><td colspan="2"> </td></tr>
	<tr><td width="500" colspan="2">
      <font face="Verdana" size="1">&nbsp;</Font>
     <input name="cField" type="checkbox" value="emp.adhocdate" /><font face="Verdana" size="1">Date</font>
     <input name="cField" type="checkbox" value="emp.empcode" /><font face="Verdana" size="1">Employee Code</Font>
	  <input name="cField1" type="checkbox" value="employee.fcname as empname"/> <font face="Verdana" size="1">Employee Name</Font> 
	  <input name="cField" type="checkbox" value="employee.gangcode" /> <font face="Verdana" size="1">Gang. Description</Font>
	   <input name="cField" type="checkbox" value="allowded.fcname"/><font face="Verdana" size="1">Allowded</Font></td>
	
	</td></tr>
	<tr><td width="100"><font face="Verdana" size="1">From </td>
	<td width="300"><input type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=TextDate(date())%>" title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.FDDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><br></td>
	<tr><td>To</td>
	<td><input type="text" maxlength="10" name="FDDATE_to" size="10" value="<%=TextDate(date())%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.FDDATE_to');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
	<tr><td><font face="Verdana" size="1">Employee Code</font></td>
	<td><input type="text" size="15" maxlength="30" name="Empcode"  style="background:#99FF99"></td>
	<tr><td><font face="Verdana" size="1">Gang</font></td>
	<td><select name="GANGCODE" >
		<%cmdCallCombo "GANG","","",""%>
	</select></td>

	<tr><td><font face="Verdana" size="1">Allowance</font></td><td>
	<select name="Allowded" >
		<%cmdCallCombo "Allowded","","",""%>
	</select>
    <tr><td><font face="Verdana" size="1">Employee</font></td><td>
	<select name="EMPLOYEECODE" >
		<%cmdCallCombo "EMPLOYEE","","",""%>
	</select></td></tr>
	<tr><td colspan="2"><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td></tr></table>

<%Elseif Request("txtKind") ="StockItem" or Request("txtKind") ="Stock" then%>
    	<form action="d_PrintData.asp?txtTable=Stock" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
	<select name="txtField" title="Case is not Sensitive">
	<Option value=""></Option>
	<Option value="FCCODE">Code</Option>
	<Option value="FCNAME">Description</Option>
	<Option value="ITEMTYPE">Item Type</Option>
	<Option value="MANUFACTUREPARTNUMBER">Manufacture Part Number</Option>
	<Option value="MANUFACTURECODE">Manufacture Code</Option>
	<Option value="SUPPLIERPARTNUMBER">Supplier Part Number</Option>
	<Option value="SUPPLIERCODE">Supplier Code</Option>
	<Option value="UNITOFMEASUREMENT">UOM</Option>
	<Option value="STOCKGROUPCODE">Stock<br>Group</Option>
	</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
<%elseif Request("txtKind") ="PurchaseOrderList" or Request("txtKind")="POOutstanding" then%>
    <%if Request("txtKind")="POOutstanding" then%>
        <form action="../p_PurchaseOrderDetails.asp?txtKind=POOutstanding" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <%Else%>
        <form action="../p_PurchaseOrderDetails.asp?txtKind=PurchaseOrderList" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%End if%>
	
     
	<input type="text" maxlength="10" name="PODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.PODATE_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	To
	<input type="text" maxlength="10" name="PODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.PODATE_To');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">PO Number</font>
	<input type="text" size="15" maxlength="30"  name="POCODE"  style="background:#99FF99">
	<font face="Verdana" size="1">Item Code</font>
	<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#99FF99">
	Supplier</font>
	<select name="SUPPLIERCODE">
		<%cmdCallCombo "SUPPLIER","","",""%>
	</select>
    <input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%elseif Request("txtKind") ="ReceiveNote" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/GRN.bmp" border="0" >
	<b>Maintenance Receive Note</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="RNDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.RNDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	To
	<input type="text" maxlength="10" name="RNDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.RNDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">RN Number</font> <input type="text" size="15" maxlength="30"  name="RECEIVENOTECODE"  style="background:#99FF99">
	<font face="Verdana" size="1">Item Code</font>	<input type="text" size="15" maxlength="30" name="PURCHASEITEMCODE"  style="background:#99FF99">
    <font face="Verdana" size="1">	Supplier</font>
	<select name="SUPPLIERCODE">
		<%cmdCallCombo "SUPPLIER","","",""%>
	</select>
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%elseif Request("txtKind") ="Invoice" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/Invoice.jpg" border="0" >
	<b>Maintenance Invoice</b></a>
    <br>
     
	<font face="Verdana" size="1">  
	<input type="text" maxlength="10" name="INVOICEDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.INVOICEDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	To
	<input type="text" maxlength="10" name="INVOICEDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.INVOICEDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">Supplier</font>
	<select name="SUPPLIERCODE">
		<%cmdCallCombo "SUPPLIER","","",""%>
	</select>
	<font face="Verdana" size="1">Invoice Number</font> <input type="text" size="15" maxlength="30"  name="INVOICECODE"  style="background:#99FF99">
	<font face="Verdana" size="1">Item Code</font>	<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Purchasing_Closing" then%>
    <b>Purchasing Close</b></a><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%elseif Request("txtKind") ="IssueDetails" or Request("txtKind")="StockIssue" then%>
	<%if Request("txtKind")="IssueDetails" then%>
        <form action="../p_IssueDetails.asp?txtKind=IssueDetails" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <%Else%>
        <form action="/p_IssueDetails.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%End if%>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="SIVDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.SIVDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="SIVDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.SIVDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">Issue No</font><input type="text" size="15" maxlength="30"  name="SIVCODE"  style="background:#99FF99">
	<font face="Verdana" size="1">Stock Code</font><input type="text" size="15" maxlength="30" name="STOCKCODE"  style="background:#99FF99">
	<!--'MODIFIED BY UKIE 07 NOV 2012 BY ADDING FILTER BY JOBCODE-->
	<font face="Verdana" size="1">Job</font><select name="JOBCODE" ><%cmdCallCombo "Job","","",""%></select>
    <font face="Verdana" size="1">TT</font>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<font face="Verdana" size="1">TC</font><input type="text" size="5" maxlength="30" name="LocationCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%elseif Request("txtKind") ="WriteOff" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><img src="ICON/WriteOff.jpg" border="0" >
	<b>Maintenance Write Off</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="WODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.WODATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="WODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.WODATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">Writeoff Code</font><input type="text" size="15" maxlength="30"  name="WriteOffCode"  style="background:#99FF99">
	<font face="Verdana" size="1">Item Code</font><input type="text" size="15" maxlength="30" name="STOCKCODE"  style="background:#99FF99">
    <font face="Verdana" size="1">TT</font>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<font face="Verdana" size="1">TC</font><input type="text" size="5" maxlength="30" name="LocationCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="StockLedger" then%>
	<form action="../p_stockledger.asp?txtKind=StockLedger" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear value="<%=txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth value="<%=txtMonth%>"></td>
		<td><font face="verdana" size="1">Store Code</font></td><td><input type=text name=txtstorecode></td>
		<td><font face="verdana" size="1">STOCKGROUP</font></td><td><input type=text name=txtstockgroup></td>
		<td><font face="verdana" size="1">STOCK</font></td><td><input type=text name=txtstock></td>
		<td><input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td>
	</tr>
	</table>
	</form>

	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
<%Elseif Request("txtKind") ="Store_Closing" then%>
    <b>Store Close</b></a><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="EmployeeList" then%>
    <form action="../PrintData.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
	<tr>
		<td><font face="verdana" size="1">CODE</font></td><td><input type=text name=txtcode value="<%=txtcode%>"></td>
		<td><font face="verdana" size="1">EmployeeName</font></td><td><input type=text name=txtemployeename value="<%=txtemployeename%>"></td>
		<td><font face="verdana" size="1">Gang</font></td><td><input type=text name=txtGangCode value="<%=txtGangCode%>"></td>
		<td><input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td>
	</tr>
	</table>
    <input  type="hidden" value="Employee" name="txtTable">
	<input  type="hidden" value="Data Master" name="txtModule">
		
<%elseif Request("txtKind") ="AttendanceList" then%>
	 <form action="/p_Attendance_GAD.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    
<table>
  <tr> 
    <Td colspan="3"><br/><br/>
      <font face="Verdana" size="1">Select Field :</Font>&nbsp;&nbsp;
	 <input name="cField" type="checkbox" value="gad.fddate" /><font face="Verdana" size="1">Date</font>
     <input name="cField" type="checkbox" value="gad.employeecode" /><font face="Verdana" size="1">Employee Code</Font>
	  <input name="cField" type="checkbox" value="employee.fcname"/> <font face="Verdana" size="1">Employee Name</Font>
	  <input name="cField" type="checkbox" value="gad.jobcode"/> <font face="Verdana" size="1">Jobcode</Font> 
	  <input name="cField1" type="checkbox" value="job.fcname as jobname" /> <font face="Verdana" size="1">Job. Description</Font>
	   <input name="cField" type="checkbox" value="gad.locationtype"/><font face="Verdana" size="1">Loc.Type</Font>
	   <input name="cField" type="checkbox" value="gad.locationcode"/><font face="Verdana" size="1">Loc.Code</Font></td>
  </tr>
  <tr> 
    <TD colspan="3"><br></font><br> 
	<input type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" > 
      <a href="javascript:show_calendar('frmEntry.FDDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"> 
      <img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
    <font face="Verdana" size="1">To</font> <input type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" > 
      <a href="javascript:show_calendar('frmEntry.FDDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;"> 
      <img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	  <font face="Verdana" size="1">Gang</font> <select name="GANGCODE">
        <%cmdCallCombo "Gang","","",""%>
      </select>
	  </TD>
  </TR>
  <TR> 
    <TD colspan="3"><font face="Verdana" size="1">Job</font> <select name="JOBCODE" >
        <%cmdCallCombo "Job","","",""%>
      </select>
	  <font face="Verdana" size="1">Employee</font> <select name="EMPLOYEECODE" >
        <%cmdCallCombo "EMPLOYEE","","",""%>
      </select>
	  <font face="Verdana" size="1">TT</font> <select name="LOCATIONTYPE">
        <%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
      </select></TD>
  </TR>
  <TR> 
    <TD colspan="3"><font face="Verdana" size="1">TC</font> <input type="text" size="15" maxlength="30" name="LOCATIONCODE"  style="background:#99FF99">
    <font face="Verdana" size="1">Report</font> <select name="Report" >
        <option value="ATTENDANCE">ATTENDANCE</option>
      </select><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></TD>
  </TR>
  
</table>
<%Elseif Request ("txtkind") = "SupplierBalance" then%>
	<form action="../p_Invoice.asp?txtKind=SupplierBalance" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
<table>		
	<br>
    
	<font face="Verdana" size="1">  
	<input type="text" maxlength="10" name="INVOICEDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.INVOICEDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	To
	<input type="text" maxlength="10" name="INVOICEDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.INVOICEDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">Supplier</font>
	<select name="SUPPLIERCODE">
		<%cmdCallCombo "SUPPLIER","","",""%>
	</select>
	<font face="Verdana" size="1">Invoice Number</font> <input type="text" size="15" maxlength="30"  name="INVOICECODE"  style="background:#99FF99">
	<font face="Verdana" size="1">Item Code</font>	<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
</table>

<%Elseif Request("txtKind") ="EmployeeActivity" then%>
    <form action="/p_Attendance_GAD.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
  <tbody><tr> 
    <td colspan="3"><br/><br/>
      <font size="1" face="Verdana">Select Field :</font>&nbsp;&nbsp;
	 <input type="checkbox" value="gad.fddate" name="cField"><font size="1" face="Verdana">Date</font>
     <input type="checkbox" value="gad.employeecode" name="cField"><font size="1" face="Verdana">Employee Code</font>
	  <input type="checkbox" value="employee.fcname" name="cField"> <font size="1" face="Verdana">Employee Name</font>
	  <input type="checkbox" value="gad.jobcode" name="cField"> <font size="1" face="Verdana">Jobcode</font> 
	  <input type="checkbox" value="job.fcname as jobname" name="cField1"> <font size="1" face="Verdana">Job. Description</font>
	   <input type="checkbox" value="gad.locationtype" name="cField"><font size="1" face="Verdana">Loc.Type</font>
	   <input type="checkbox" value="gad.locationcode" name="cField"><font size="1" face="Verdana">Loc.Code</font></td>
  </tr>
  <tr> 
    <td colspan="3"><br><font size="1" face="Verdana"> :</font><br> 
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="FDDATE_From" maxlength="10"> 
      <a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.FDDATE_From');"> 
      <img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
    <font size="1" face="Verdana">To</font> <input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="FDDATE_To" maxlength="10"> 
      <a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.FDDATE_To');"> 
      <img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
	  <font size="1" face="Verdana">Gang</font> <select name="GANGCODE">
        
	 <%cmdCallCombo "GANG","","",""%>
		
      </select>
	  </td>
  </tr>
  <tr> 
    <td colspan="3"><font size="1" face="Verdana">Job</font> <select name="JOBCODE">
        
	 <%cmdCallCombo "JOB","","",""%>
		
      </select>
	  <font size="1" face="Verdana">Employee</font> <select name="EMPLOYEECODE">
        
	 <%cmdCallCombo "EMPLOYEE","","",""%>
		
      </select>
	  <font size="1" face="Verdana">TT</font> <select name="LOCATIONTYPE">
        
	 <option selected="" value=""></option>
		        <option value="AP">AP</option>  
		    
		        <option value="BM">BM</option>  
		    
		        <option value="CA">CA</option>  
		    
		        <option value="CB">CB</option>  
		    
		        <option value="BA">BA</option>  
		    
		        <option value="EM">EM</option>  
		    
		        <option value="ES">ES</option>  
		    
		        <option value="FF">FF</option>  
		    
		        <option value="GC">GC</option>  
		    
		        <option value="MA">MA</option>  
		    
		        <option value="MN">MN</option>  
		    
		        <option value="PN">PN</option>  
		    
		        <option value="PR">PR</option>  
		    
		        <option value="SA">SA</option>  
		    
		        <option value="ST">ST</option>  
		    
		        <option value="VH">VH</option>  
		    
		        <option value="WC">WC</option>  
		    
		        <option value="WS">WS</option>  
		    
		        <option value="CP">CP</option>  
		    
		        <option value="PJ">PJ</option>  
		    
		        <option value="LOC">LOC</option>  
		    
        </select></td>
       </tr>
     <tr> 
    <td colspan="3"><font size="1" face="Verdana">TC</font> <input type="text" style="background:#99FF99" name="LOCATIONCODE" maxlength="30" size="15">
    <font size="1" face="Verdana">Report</font> <select name="Report">
        <option value="ACTIVITY">ACTIVITY</option>\
	</select><input type="submit" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" value="Enter"></td>
  </tr>
  
</tbody>
</table>
</form>
	
<%elseif Request("txtKind") ="EmpAllded" then%>
<form action="../p_EmpAllded.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    
	<Table>
  <tr> 
    <Td colspan="3">
      <font face="Verdana" size="1">Select Field :</Font><br>
     <input name="cField" type="checkbox" value="emp.adhocdate" /><font face="Verdana" size="1">Date</font>
     <input name="cField" type="checkbox" value="emp.empcode" /><font face="Verdana" size="1">Employee Code</Font>
	  <input name="cField1" type="checkbox" value="employee.fcname as empname"/> <font face="Verdana" size="1">Employee Name</Font> 
	  <input name="cField" type="checkbox" value="employee.gangcode" /> <font face="Verdana" size="1">Gang. Description</Font>
	   <input name="cField" type="checkbox" value="allowded.fcname"/><font face="Verdana" size="1">Allowded</Font></td>
  </tr>	
	<tr>
	<Td colspan="3"><br></font><br> 
	<input type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.FDDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.FDDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
    <font face="Verdana" size="1">Gang</font>
	<select name="GANGCODE">
		<%cmdCallCombo "Gang","","",""%>
	</select></td>
	</tr>
	<tr>
	<Td colspan="3">
    <font face="Verdana" size="1">Allowance</font>
	<select name="Allowded" >
		<%cmdCallCombo "Allowded","","",""%>
	</select>
    <font face="Verdana" size="1">Employee</font>
	<select name="EMPLOYEECODE" >
		<%cmdCallCombo "EMPLOYEE","","",""%>
	</select>
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</td>
	</tr>
	</Table>
<%Elseif Request("txtKind") ="Calculate" then%>
    <b>Calculate Payroll</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Employee_Closing" then%>
    <b>Employee Close</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%elseif Request("txtKind") ="HarvestingQuality" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="ICON/HarvestingQuality.jpg" border="0" >
	<b>Maintenance Harvesting Quality</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.FDDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.FDDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
    <font face="Verdana" size="1">Employee</font>
	<select name="EMPLOYEECODE" >
		<%cmdCallCombo "EMPLOYEE","","",""%>
	</select>
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	
	
<%elseif Request("txtKind") ="WorkshopList" then%>
	<form action="../p_WorkshopList.asp?txtKind=WorkshopList" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
	<tr>
	<br>
	<font size="1" face="Verdana"> 
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_From" maxlength="10">
	<a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
	<font size="1" face="Verdana">To</font>
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_To" maxlength="10">
	<a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
    <font size="1" face="Verdana">Activity</font>
	<select name="JOB"><%cmdCallCombo "JOB","","",""%>
		
	</select>
    <font size="1" face="Verdana">TT</font>
	<select name="LOCATIONTYPE">
		
	 <option selected="" value=""></option>
		        <option value="AP">AP</option>  
		    
		        <option value="BM">BM</option>  
		    
		        <option value="CA">CA</option>  
		    
		        <option value="CB">CB</option>  
		    
		        <option value="BA">BA</option>  
		    
		        <option value="EM">EM</option>  
		    
		        <option value="ES">ES</option>  
		    
		        <option value="FF">FF</option>  
		    
		        <option value="GC">GC</option>  
		    
		        <option value="MA">MA</option>  
		    
		        <option value="MN">MN</option>  
		    
		        <option value="PN">PN</option>  
		    
		        <option value="PR">PR</option>  
		    
		        <option value="SA">SA</option>  
		    
		        <option value="ST">ST</option>  
		    
		        <option value="VH">VH</option>  
		    
		        <option value="WC">WC</option>  
		    
		        <option value="WS">WS</option>  
		    
		        <option value="CP">CP</option>  
		    
		        <option value="PJ">PJ</option>  
		    
		        <option value="LOC">LOC</option>  
		    
	</select>
	<font size="1" face="Verdana">TC</font><input type="text" style="background:#99FF99" name="LocationCode" maxlength="15" size="5">
	<input type="submit" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" value="Enter">
			
		</font>
	</tr>
	</table>
    <input  type="hidden" value="Workshop" name="txtTable">
	<input  type="hidden" value="Data Master" name="txtModule">
	
<%elseif Request("txtKind") ="WorkshopRunningAccount" then%>
	<form action="../p_WorkshopActivity.asp?txtKind=WorkshopRunningAccount" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
	<tr>
	<br>
	<font size="1" face="Verdana"> 
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_From" maxlength="10">
	<a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
	<font size="1" face="Verdana">To</font>
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_To" maxlength="10">
	<a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
    <font size="1" face="Verdana">Activity</font>
	<select name="JOB">
		
	 <%cmdCallCombo "JOB","","",""%>
		
	</select>
    <font size="1" face="Verdana">TT</font>
	<select name="LOCATIONTYPE">
		
	 <option selected="" value=""></option>
		        <option value="AP">AP</option>  
		    
		        <option value="BM">BM</option>  
		    
		        <option value="CA">CA</option>  
		    
		        <option value="CB">CB</option>  
		    
		        <option value="BA">BA</option>  
		    
		        <option value="EM">EM</option>  
		    
		        <option value="ES">ES</option>  
		    
		        <option value="FF">FF</option>  
		    
		        <option value="GC">GC</option>  
		    
		        <option value="MA">MA</option>  
		    
		        <option value="MN">MN</option>  
		    
		        <option value="PN">PN</option>  
		    
		        <option value="PR">PR</option>  
		    
		        <option value="SA">SA</option>  
		    
		        <option value="ST">ST</option>  
		    
		        <option value="VH">VH</option>  
		    
		        <option value="WC">WC</option>  
		    
		        <option value="WS">WS</option>  
		    
		        <option value="CP">CP</option>  
		    
		        <option value="PJ">PJ</option>  
		    
		        <option value="LOC">LOC</option>  
		    
	</select>
	<font size="1" face="Verdana">TC</font><input type="text" style="background:#99FF99" name="LocationCode" maxlength="15" size="5">
	<input type="submit" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" value="Enter">
			
		</font>
	</tr>
	</table>
    <input  type="hidden" value="Workshop" name="txtTable">
	<input  type="hidden" value="Data Master" name="txtModule">
	
<%Elseif Request("txtKind") ="Workshop_Closing" then%>
    <b>Workshop Close</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%elseif Request("txtKind") ="VehicleList" then%>
	<form action="../PrintData.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <br>
    <font size="1" face="verdana">
	Search :
		<input width="5" type="text" value="" name="txtSearch" title="Case is not Sensitive for date [dd/mm/yyyy]">
			<font size="1" face="verdana">
			& nbsp;from Column
		<select title="Case is not Sensitive" name="txtField">
			<option value=""></option>
			<option value="FCCODE">Code</option>
			<option value="FCNAME">Description</option>
			<option value="VEHICLEGROUPCODE">VehicleGroup Code</option>
			<option value="REGISTRATIONNO">RegistrationNo</option>
			<option value="ACTIVATION">Active</option>
			<option value="OWNERSHIP">Ownership</option>
			<option value="VEHICLESTATUS">VehicleStatus</option>
			<option value="CONTROLALLOCATION">ControlAllocation</option>
			<option value="CONTROLJOB">ControlAccount</option>
		</select>
		<input type="submit" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" value="Enter">
			</font>
	</font>
	
	<input  type="hidden" value="Vehicle" name="txtTable">
	<input  type="hidden" value="Data Master" name="txtModule">

<%Elseif Request("txtKind") ="VehicleRunningAccount" then%>
	<form action="../p_VehicleActivity.asp?txtKind=VehicleRunningAccount" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
	<br>
   <font size="1" face="Verdana"> : 
	<font size="1" face="Verdana"> 
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="VEHDATE_From" maxlength="10">
	<a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.VEHDATE_From');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
	<font size="1" face="Verdana">To</font>
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="VEHDATE_To" maxlength="10">
	<a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.VEHDATE_To');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
    <font size="1" face="Verdana">Vehicle</font>
	<select name="VEHICLECODE"><%cmdCallCombo "VEHICLE","","",""%>
		
	</select>
    <font size="1" face="Verdana">Activity</font>
	<select name="JOBCODE"><%cmdCallCombo "JOB","","",""%>
		
	</select>
    <font size="1" face="Verdana">TT</font>
	<select name="LOCATIONTYPE">
		
	 <option selected="" value=""></option>
		        <option value="AP">AP</option>  
		    
		        <option value="BM">BM</option>  
		    
		        <option value="CA">CA</option>  
		    
		        <option value="CB">CB</option>  
		    
		        <option value="BA">BA</option>  
		    
		        <option value="EM">EM</option>  
		    
		        <option value="ES">ES</option>  
		    
		        <option value="FF">FF</option>  
		    
		        <option value="GC">GC</option>  
		    
		        <option value="MA">MA</option>  
		    
		        <option value="MN">MN</option>  
		    
		        <option value="PN">PN</option>  
		    
		        <option value="PR">PR</option>  
		    
		        <option value="SA">SA</option>  
		    
		        <option value="ST">ST</option>  
		    
		        <option value="VH">VH</option>  
		    
		        <option value="WC">WC</option>  
		    
		        <option value="WS">WS</option>  
		    
		        <option value="CP">CP</option>  
		    
		        <option value="PJ">PJ</option>  
		    
		        <option value="LOC">LOC</option>  
		    
	</select>
	<font size="1" face="Verdana">TC</font><input type="text" style="background:#99FF99" name="LocationCode" maxlength="30" size="15">
    <input type="submit" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" value="Enter">


			<!-- Akhir Tampilan Select form -->
			
		</font></font>
	
<%Elseif Request("txtKind") ="Vehicle_Closing" then%>
    <b>Vehicle Close</b></a><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%elseif request("txtkind") ="ContractorList" then%>
	<form action="../PrintData.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<tr>
	<br>
	<font size="1" face="Verdana"> 
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_From" maxlength="10">
	<a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
	<font size="1" face="Verdana">To</font>
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_To" maxlength="10">
	<a onmouseout="window.status='';return true;" onmouseover="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
    <font size="1" face="Verdana">Code</font><input type="text" style="background:#99FF99" name="" maxlength="30" size="15">
	<input type="submit" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" value="Enter">

	</font></font>
</tr>	
	<input  type="hidden" value="Contractor" name="txtTable">
	<input  type="hidden" value="Data Master" name="txtModule">

<%elseif Request("txtKind") ="ContractAgreement" then%>
	<form action="../p_AGREEMENTDETAIL.asp?" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font size="1" face="Verdana"> 
	<br>
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	To
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	 <font size="1" face="Verdana">Code</font><input type="text" style="background:#99FF99" name="" maxlength="30" size="15">
	 <input type="submit" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" value="Enter">
</font>

	
<%elseif Request("txtKind") ="AGREEMENTDETAIL" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<br>
	<img src="ICON/Contract.jpg" border="0" ><br>
	<b>Maintenance Contract Agreement</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
    <font face="Verdana" size="1">Activity</font>
	<select name="JOB" >
		<%cmdCallCombo "job","","",""%>
	</select>
    <font face="Verdana" size="1">TT</font>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<font face="Verdana" size="1">TC</font><input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%elseif Request("txtKind") ="WorkinProgress" then%>
	<form action="../p_WIP.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <br>
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
    <font face="Verdana" size="1">Activity</font>
	<select name="JOB" >
		<%cmdCallCombo "job","","",""%>
	</select>
    <font face="Verdana" size="1">TT</font>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<font face="Verdana" size="1">TC</font><input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%Elseif Request("txtKind") ="Contract_Closing" then%>
    <b>Contract Close</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%elseif Request("txtKind") ="NurseryList" then%>
	<form action="../p_Nursery.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry"><br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Batch</font><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#99FF99">
 	<font face="Verdana" size="1">Crop</font><input type="text" size="15" maxlength="30" name="Crop"  style="background:#99FF99">
 	<font face="Verdana" size="1">Source</font><input type="text" size="15" maxlength="30" name="Source"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%elseif Request("txtKind") ="NurseryOrder" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Nursery Order</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Batch</font><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%elseif Request("txtKind") ="NurseryCulling" then%>
<form action="../p_NurseryCulling.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Batch</font><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%elseif Request("txtKind") ="NurseryIssue" then%>
<form action="../p_NurseryIssue.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Batch</font><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#99FF99">
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<font face="Verdana" size="1">TC</font><input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%elseif Request("txtKind") ="NurseryLedger" then%>
    <form action="../p_NurseryLedger.asp?txtKind=NurseryList" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry"><br>
    <font face="Verdana" size="1"><br> : 
	<font face="Verdana" size="1"> 
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
		<font face="verdana" size="1">YEAR</font><input type=text name=txtyear value="<%=txtyear%>">
		<font face="verdana" size="1">MONTH</font><input type=text name=txtmonth value="<%=txtmonth%>">
	<font face="Verdana" size="1">Batch</font><input type="text" size="15" maxlength="30" name="txtnurserycode"  style="background:#99FF99">
    <input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%Elseif Request("txtKind") ="Nursery_Closing" then%>
    <b>Nursery Close</b></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%elseif Request("txtKind") ="ProjectList" then%>
	 <form action="../p_Project.asp?txtKind=ProjectList" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	 <br>
	<br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#99FF99">
 	<font face="Verdana" size="1">Crop</font>
		<% 	Dim rsCropproject
		q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
		Set rsCropproject = Server.CreateObject("ADODB.Recordset")
		rsCropproject.CursorType = 3
		rsCropproject.CursorLocation = 3
		rsCropproject.Open q, connstr
		if not rsCropproject.eof then%>
		<select name="Crop">         
		<option value=""></option> 
		<%
		Do While Not rsCropproject.eof
		%>
		<option value="<%=rsCropproject("FCCODE")%>"><%=rsCropproject("FCNAME")%></option> 
		<%rsCropproject.movenext
		loop
		End if
		rsCropproject.Close
		set rsCropproject = Nothing
		q=""
		%>
		</select>				
 	<font face="Verdana" size="1">Source</font><input type="text" size="15" maxlength="30" name="Source"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%elseif Request("txtKind") ="Project_Production" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Production</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%elseif Request("txtKind") ="Project_Adjustment" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Adjusment</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%elseif Request("txtKind") ="Project_Issue" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>"><br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Issue</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#99FF99">
	<font face="Verdana" size="1">LC</font>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<font face="Verdana" size="1">TC</font><input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#99FF99">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%elseif Request("txtKind") ="ProjectLedger" then%>
	 <form action="../p_Project_Ledger.asp?txtKind=ProjectLedger" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
     <font face="Verdana" size="1"><br> : 
	<font face="Verdana" size="1"> 
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
		<font face="verdana" size="1">Year</font><input type=text name=txtyear value="<%=txtyear%>">
		<font face="verdana" size="1">Month</font><input type=text name=txtmonth value="<%=txtmonth%>">
	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="txtprojectcode"  style="background:#99FF99">
    <input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	
<%Elseif Request("txtKind") ="Project_Close" then%>
    <b>Project Close</b></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Building_Closing" then%>
    <b>Building Close</b></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Building_Closing" then%>
    <b>Building Close</b></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="General_Charges_Closing" then%>
    <b>General Charge Close</b></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Harvesting_Closing" then%>
    <b>Close Harvesting</b></a><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Financial_Monthly_Closing" then%>
    <b>Closing This Month</b></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Previous_Month_UnClosing" then%>
    <b>un-Closing This Month</b></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Financial_Year_Closing" then%>
    <b>Close Financial Year</b></a>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="TrialBalance" then%>
	<form action="../p_GLMasterRegisterMonthly.asp?txtKind=TrialBalance" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
	<form action=p_glmasterregistermonthly.asp method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear value="<%=txtyear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth value="<%=txtmonth%>"></td>
		<td><font face="verdana" size="1">JOBGROUP</font></td><td><input type=text name=txtjobgroup></td>
		<td><font face="verdana" size="1">JOB</font></td><td><input type=text name=txtjob></td>
		<td><input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td>
	</tr>
	</table>
	</form>
<%Elseif Request("txtKind") ="Cost_Activity" then
    if request("txtYear")="" then
	   txtYear=Year(Now)
	end if
    if request("txtMonth")="" then
	   txtMonth=Month(Now)
	end if
	%>
	<form action=../p_Cost_Activity.asp method="post" target="_blank">
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type=text name=txtyear width="4" value="<%=txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type=text name=txtMonth width="3" value="<%=txtMonth%>"></td>
		<td><font face="verdana" size="1">JOBGROUP</font></td><td width="4"><input type=text name=txtjobgroup width="6"></td>
		<td><font face="verdana" size="1">JOB</font></td><td width="4"><input type=text name=txtjob width="10"></td></TR><TR>
		<td><font face="verdana" size="1">ST</font></td><td width="4"><input type=text name=txtST width="5"></td>
		<td><font face="verdana" size="1">SC </font></td><td width="4"><input type=text name=txtSC width="10"></td>
		<td><font face="verdana" size="1">TT</font></td><td width="4"><input type=text name=txtTT width="5"></td>
		<td><font face="verdana" size="1">TC</font></td><td width="4"><input type=text name=txtTC width="10"></td>
	</tr>

	<tr>
		<td colspan="3" align="right"><font face="verdana" size="1">Reference</font></td><td width="4"><input type=text name=txtreference width="4" value="<%=txtReference%>"></td>
		<td colspan="5"><input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td>
	</tr>
	</table>
	</form>
<%Elseif Request("txtKind") ="Neraca" then%>
    <b><br><img src="ICON/Neraca.jpg" border="0"><br>Neraca</b></a>
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
	<form action=p_neraca.asp method=post>
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth></td>
		<td><font face="verdana" size="1">JOBGROUP</font></td><td><input type=text name=txtjobgroup></td>
		<td><font face="verdana" size="1">JOB</font></td><td><input type=text name=txtjob></td>
		<td><input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td>
	</tr>
	</table>
	</form>
<%elseif Request("txtKind") ="Harvesting" then%>
<form action="../p_Harvesting.asp" method="post" target="_blank" name="frmEntry">
	<table>
	<tr>
	<td><font face="Verdana" size="1">From</td>
	<td><input type="text" maxlength="10" name="RECEPTIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.RECEPTIONDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
	<td><font face="Verdana" size="1">To</font></td>
	<td><input type="text" maxlength="10" name="RECEPTIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.RECEPTIONDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
 	<td><font face="Verdana" size="1">Chit No</font></td>
	<td><input type="text" size="15" maxlength="30" name="CHITNO"></td>
	</tr>
	<tr>
	<td><font face="Verdana" size="1">Field Code</font></td>
	<td><select name="FIELDCODE">
	<%cmdCallCombo "Field","FIELDCODE","",""%>
	</select></td>
	<td><font face="Verdana" size="1">Cropcode Code</font></td>					
	<% 	Dim rsCrop
	q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
	Set rsCrop = Server.CreateObject("ADODB.Recordset")
	rsCrop.CursorType = 3
	rsCrop.CursorLocation = 3
	rsCrop.Open q, connstr
	if not rsCrop.eof then%>
	<td><select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
	<option value=""></option> 
	<%
	Do While Not rsCrop.eof
	%>
	<option value="<%=rsCrop("FCCODE")%>"><%=rsCrop("FCNAME")%></option> 
	<%rsCrop.movenext
	loop
	End if
	rsCrop.Close
	set rsCrop = Nothing
	q=""
	%></select></td>
	<td><font face="Verdana" size="1">Product Code</font></td>
	<td><select name="ProductCODE">
		<%cmdCallComboSetup "PRODUCT","ProductCODE","",""%>
		</select>
	</td>	
	<td><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td>
	</tr>
	</table>
</form>	
<%elseif Request("txtKind") ="HarvestingRound" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Harvesting Round</b></a>
    <br>
     
	<table>
	<tr>
	<td><font face="verdana" size="1">Year</font></td><td><input type=text size="10" maxlength="4" name=HRNDYEAR></td>
	<td><font face="verdana" size="1">Month</font></td><td><input type=text size="10" maxlength="2" name=HRNDMONTH></td>
	<td><font face="Verdana" size="1">Field Code</font></td>
	<td><select name="FIELDCODE">
	<%cmdCallCombo "Field","FIELDCODE","",""%>
	</select></td>
	<td><font face="Verdana" size="1">Cropcode Code</font></td>					
	<% 	Dim rsCropA
	q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
	Set rsCropA = Server.CreateObject("ADODB.Recordset")
	rsCropA.CursorType = 3
	rsCropA.CursorLocation = 3
	rsCropA.Open q, connstr
	if not rsCropA.eof then%>
	<td><select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
	<option value=""></option> 
	<%
	Do While Not rsCropA.eof
	%>
	<option value="<%=rsCropA("FCCODE")%>"><%=rsCropA("FCNAME")%></option> 
	<%rsCropA.movenext
	loop
	End if
	rsCropA.Close
	set rsCropA = Nothing
	q=""
	%></select>
	</td>
	<td><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	</td>
	</tr>
	</table>
<%elseif Request("txtKind") ="Weather" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Weather</b></a>
    <br>
     
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="WIDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.WIDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="WIDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.WIDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Division</font><input type="text" size="15" maxlength="30" name="WIDIVISION">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	

<%elseif Request("txtKind") ="PRODUCT" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Product</b></a>
    <br>
     
 	<font face="Verdana" size="1">Product Code</font><input type="text" size="15" maxlength="30" name="PRODUCTCODE">
 	<font face="Verdana" size="1">Crop Code</font>	
	<% 	Dim rsCropP
	q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
	Set rsCropP = Server.CreateObject("ADODB.Recordset")
	rsCropP.CursorType = 3
	rsCropP.CursorLocation = 3
	rsCropP.Open q, connstr
	if not rsCropP.eof then%>
	<select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
	<option value=""></option> 
	<%
	Do While Not rsCropP.eof
	%>
	<option value="<%=rsCropP("FCCODE")%>"><%=rsCropP("FCNAME")%></option> 
	<%rsCropP.movenext
	loop
	End if
	rsCropP.Close
	set rsCropP = Nothing
	q=""
	%></select>
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	
<%elseif Request("txtKind") ="WEIGHBRIDGE" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Weightbridge</b></a>
    <br>
     
 	<font face="Verdana" size="1">Weighbridge Code</font><input type="text" size="15" maxlength="30" name="WEIGHBRIDGECODE">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	

<%elseif Request("txtKind") ="CUSTOMER" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Customer</b></a>
    <br>
     
 	<font face="Verdana" size="1">Customer Code</font><input type="text" size="15" maxlength="30" name="CUSTOMERCODE">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	
<%elseif Request("txtKind") ="GRADINGITEM" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Grading Item</b></a>
    <br>
     
 	<font face="Verdana" size="1">Item Code</font><input type="text" size="15" maxlength="30" name="ITEMCODE">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	
<%elseif Request("txtKind") ="OPHarvRate" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance OP HARVESTING RATE</b></a>
    <br>
     
 	<table>
	<tr>
	<td><font face="Verdana" size="1">Crop Code</font></td>	
	<% 	Dim rsCropo
	q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcBA='" &  Session("fcBA") & "' order by P.FCNAME desc"
	Set rsCropo = Server.CreateObject("ADODB.Recordset")
	rsCropo.CursorType = 3
	rsCropo.CursorLocation = 3
	rsCropo.Open q, connstr
	if not rsCropo.eof then%>
	<td><select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
	<option value=""></option> 
	<%
	Do While Not rsCropo.eof
	%>
	<option value="<%=rsCropo("FCCODE")%>"><%=rsCropo("FCNAME")%></option> 
	<%rsCropo.movenext
	loop
	End if
	rsCropo.Close
	set rsCropo = Nothing
	q=""
	%></select></td>
 	<td><font face="Verdana" size="1">Terrain Type</font></td>
	<td><input type="text" size="15" maxlength="30" name="TERRAIN_TYPE"></td>
	<td><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td>
	</tr>
	</table>

<%elseif Request("txtKind") ="WEIGHBRIDGETICKET" then%>
<form action="../p_WEIGHBRIDGETICKET.asp" method="post" target="_blank" name="frmEntry">
	<!--ADDED BY UKIE 14-NOV-2012-->
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="Dateout_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.Dateout_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="/Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="Dateout_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.Dateout_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="/Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<!--END ADDING-->
 	<font face="Verdana" size="1">Product Code</font>
		<select name="PRODUCT_PRODUCTCODE">
		<%cmdCallComboSetup "PRODUCT","ProductCODE","",""%>
		</select>
 	<font face="Verdana" size="1">Chit No</font><input type="text" size="15" maxlength="30" name="CHITNO">
 	<font face="Verdana" size="1">Customer Code</font>
		<select name="CUST_CUSTOMERCODE">
		<%cmdCallComboSetup "CUSTOMER","CUSTOMERCODE","",""%>
		</select>
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	
</form>
	
<%elseif Request("txtKind") ="GRADINGRECORD" then%>
<form action="../p_GRADINGRECORD.asp" method="post" target="_blank" name="frmEntry">
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="DATETIME_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATETIME_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATETIME_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATETIME_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Grading No</font><input type="text" size="15" maxlength="30" name="GRADINGNO">
 	<font face="Verdana" size="1">Vehicle No</font><input type="text" size="15" maxlength="30" name="VEHICLENO">
 	<font face="Verdana" size="1">Chitno</font><input type="text" size="15" maxlength="30" name="WBT_CHITNO">

	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
</form>
<%elseif Request("txtKind") ="SHIFT" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Shift</b></a>
    <br>
     
 	<font face="Verdana" size="1">Shift Code</font><input type="text" size="15" maxlength="30" name="SHIFTCODE">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	
<%elseif Request("txtKind") ="SHIFTREPORT" then%>
<form action="../p_SHIFTREPORT.asp" method="post" target="_blank" name="frmEntry">
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="PRODUCTIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.PRODUCTIONDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="PRODUCTIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.PRODUCTIONDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Shift Code</font><select name="SHIFT_SHIFTCODE"><%cmdCallComboSetup "SHIFT","SHIFT_SHIFTCODE","",""%></select>
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
</form>
<%elseif Request("txtKind") ="PRODUCTSTORAGE" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Product Storage</b></a>
    <br>
     
 	<font face="Verdana" size="1">Storage Code</font><input type="text" size="15" maxlength="30" name="STORAGECODE">
 	<font face="Verdana" size="1">Product Code</font>
		<select name="PRODUCT_PRODUCTCODE">
		<%cmdCallComboSetup "PRODUCT","PRODUCT_PRODUCTCODE","",""%>
		</select>
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	

<%elseif Request("txtKind") ="PRODUCTSTORAGEREADING" then%>
<form action="../p_PRODUCTSTORAGEREADING.asp" method="post" target="_blank" name="frmEntry">
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="READINGDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.READINGDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="READINGDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.READINGDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Storage Code</font><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","",""%></select>	
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
</form>
<%elseif Request("txtKind") ="PRODUCTADJUSTMENT" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Product Adjusment</b></a>
    <br>
     
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="ADJUSTMENTDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.ADJUSTMENTDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="ADJUSTMENTDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.ADJUSTMENTDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Adjusment No</font><input type="text" size="15" maxlength="30" name="ADJUSTMENTNO">
 	<font face="Verdana" size="1">Storage Code</font><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","",""%></select>
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	

<%elseif Request("txtKind") ="PRODUCTSTORAGETRANSFER" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Storage Transfer</b></a>
    <br>
     
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="TRANSFERDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.TRANSFERDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="TRANSFERDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.TRANSFERDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Transfer No</font><input type="text" size="15" maxlength="30" name="TRANSFERNO">
 	<font face="Verdana" size="1">Storage Code</font><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","",""%></select>
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	

<%elseif Request("txtKind") ="DELIVERYCONTRACT" then%>
<form action="../p_DELIVERYCONTRACT.asp" method="post" target="_blank" name="frmEntry">
	<TABLE>
	<TR>
	<TD><font face="Verdana" size="1">From</FONT></TD><TD><input type="text" maxlength="10" name="STARTDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.STARTDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></TD>
	<TD><font face="Verdana" size="1">To</font></TD><TD><input type="text" maxlength="10" name="STARTDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.STARTDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></TD>
 	<TD><font face="Verdana" size="1">Contract Code</font></TD><TD><input type="text" size="15" maxlength="30" name="CONTRACTCODE"></TD>
	<TR>
 	<TD><font face="Verdana" size="1">Customer Code</font></TD><TD><select name="CUST_CUSTOMERCODE"><%cmdCallComboSetup "CUSTOMER","CUSTOMERCODE","",""%></select></TD>
 	<TD><font face="Verdana" size="1">Product Code</font></TD><TD><select name="PRODUCT_PRODUCTCODE"><%cmdCallComboSetup "PRODUCT","ProductCODE","",""%></select></TD>
	<TD></TD><TD><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></TD>	
	</TR>
	</TABLE>
</form>	
<%elseif Request("txtKind") ="DELIVERYNOTE" then%>
<form action="../p_DELIVERYNOTE.asp" method="post" target="_blank" name="frmEntry">
	<TABLE>
	<TR>
	<TD><font face="Verdana" size="1">From</FONT></TD><TD><input type="text" maxlength="10" name="DNOTE_DATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DNOTE_DATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></TD>
	<TD><font face="Verdana" size="1">To</font></TD><TD><input type="text" maxlength="10" name="DNOTE_DATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DNOTE_DATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></TD>
 	<TD><font face="Verdana" size="1">Delivery NOte Code</font></TD><TD><input type="text" size="15" maxlength="30" name="DELIVERYNOTECODE"></TD>
	<TR>
 	<TD><font face="Verdana" size="1">Delivery Note Code</font></TD><TD><input type="text" size="15" maxlength="30" name="DELIVERYNOTECODE"></TD>
 	<TD><font face="Verdana" size="1">Contract Code</font></TD>
	    	<% 	Dim rsContract
			q = "select D.CONTRACTCODE from  DELIVERYCONTRACT D"
			Set rsContract = Server.CreateObject("ADODB.Recordset")
			rsContract.CursorType = 3
			rsContract.CursorLocation = 3
			rsContract.Open q, connstr
			if not rsContract.eof then%>
			<td align="left"> <select name="CONTRACTCODE" title="CONTRACTCODE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%
			Do While Not rsContract.eoF%>
				<option value="<%=rsContract("CONTRACTCODE")%>"><%=rsContract("CONTRACTCODE")%></option> 
			<%rsContract.movenext
			loop
			End if
			rsContract.Close
			set rsContract = Nothing
			q=""
			%>
			</select></td>
	<TD></TD><TD><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></TD>	
	</TR>
	</TABLE>
</form>
<%elseif Request("txtKind") ="DELIVERYRETURNNOTE" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Delivery Return Note</b></a>
    <br>
     
	<table>
	<tr>
	<td><font face="Verdana" size="1">From</font></td>
	<td><input type="text" maxlength="10" name="RETURNDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.RETURNDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
	<td><font face="Verdana" size="1">To</font></td>
	<td><input type="text" maxlength="10" name="RETURNDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.RETURNDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
 	<td><font face="Verdana" size="1">Delivery Note Code</font></td><td><input type="text" size="15" maxlength="30" name="DNOTE_DELIVERYNOTECODE"></td>
	</tr>
	<tr> 
 	<td><font face="Verdana" size="1">WB Chit No</font></td><td><input type="text" size="15" maxlength="30" name="WBT_CHITNO"></td>
 	<td><font face="Verdana" size="1">Storage Code</font></td><td><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","",""%></select></td>
	<td><input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td>	
	</tr>
	</table>
<%elseif Request("txtKind") ="DeliveryQuality" then%>
	<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
	title="New : <%=txtTable%>">
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Delivery Product Quality</b></a>
    <br>
     
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="QUALITYDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.QUALITYDATE_From');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="QUALITYDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.QUALITYDATE_To');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Delivery Note Code</font><input type="text" size="15" maxlength="30" name="DELIVERYNOTE">
	<input type="submit" value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">		
<%Elseif Request("txtKind") ="Stock_Product" then%>
    <b>Stock Product Ledger</b><img src="ICON/StockLedger.jpg" border="0" ></a>
	
	<table>
	<tr>
		<td><font face="verdana" size="1">Year</font></td><td><input type=text name=txtyear></td>
		<td><font face="verdana" size="1">Month</font></td><td><input type=text name=txtmonth></td>
		<td><font face="verdana" size="1">Storage Code</font></td><td><input type=text name=txtSTORAGE_CODE></td>
		<td><font face="verdana" size="1">Product Code</font></td><td><input type=text name=txtPRODUCT_CODE></td>
		<td><input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"></td>
	</tr>
	</table>
	
<%Elseif Request("txtKind") ="Close_Delivery" then%>
    <b>Close Delivery</b></a><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br>
	<input type="submit" value="::CLOSE::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Calculate_Mill_Product" then%>
    <b>Calculate Delivery</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="WeighBridge_Closing" then%>
    <b>WeighBridge Closing </b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Production_Closing" then%>
    <b>Production Closing </b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Delivery_Closing" then%>
    <b>Delivery Closing</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="Clone" then%>
    <b>Clone Data Master to All Business Unit</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Clone::" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
<%Elseif Request("txtKind") ="InvoiceLedger" then%>
<br /><br />
	<form action="p_InvoiceLedger.asp" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type="text" name="txtyear" value="<%=txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type="text" name="txtmonth" value="<%=txtMonth%>"></td>
		<td><font face="verdana" size="1">SUPPLIER CODE</font></td><td><select name="txtsuppliercode" ><%cmdCallCombo "Supplier","","",""%></select></td>
		<td><font face="verdana" size="1">INVOICE CODE</font></td><td><input type="text" name="txtinvoicecode" /></td>
		<td><input type="submit" value="View" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" /></td>
	</tr>
	</table>
	</form>
<%Elseif Request("txtKind") ="ContractorLedger" then%>
	<form action="p_ContractorLedger.asp" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type="text" name="txtyear" value="<%=txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type="text" name="txtmonth" value="<%=txtMonth%>"></td>
		<td><font face="verdana" size="1">CONTRACTOR CODE</font></td><td><input type="text" name="txtcontractorcode" /></td>
		<td><font face="verdana" size="1">AGREEMENT CODE</font></td><td><input type="text" name="txtagreementcode" /></td>
		<td><font face="verdana" size="1">AGREEMENT LINE NO</font></td><td><input type="text" name="txtagreementlineno" /></td>
		<td><input type="submit" value="View" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" /></td>
	</tr>
	</table>
	</form>
<%Elseif Request("txtKind") ="Aging" then%>
	<form action="p_Aging.asp" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
	<tr>
		<td><font face="verdana" size="1">SUPPLIER CODE</font></td><td><select name="txtsuppliercode" ><%cmdCallCombo "Supplier","","",""%></select></td>
		<td><font face="verdana" size="1">INVOICE CODE</font></td><td><input type="text" name="txtinvoicecode" /></td>
		<td><input type="submit" value="View" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana" /></td>
	</tr>
	</table>
	</form>
<%Elseif Request("txtKind") ="PayEmployee" then%>
	<form action="p_PayEmployee.asp" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type="text" name="txtyear" value="<%=txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type="text" name="txtmonth" value="<%=txtMonth%>"></td>
		<td><font face="verdana" size="1">Employee</font></td>
		<td>
		<%
			Dim rsEmployee
			strEmployee = "SELECT * FROM EMPLOYEE WHERE FCBA = '" & Session("fcBA") & "' ORDER BY FCCODE"
			
			Set rsEmployee = Server.CreateObject("ADODB.Recordset")
			rsEmployee.CursorType = 3
			rsEmployee.CursorLocation = 3
			rsEmployee.Open strEmployee, connstr
			
			if not rsEmployee.eof then
		%>
			<select name="txtEmployeeCode" >
		<%
				Do While Not rsEmployee.eof
					if rsEmployee("FCCODE") = txtEmployeeCode then
		%>				
						<option value="<%=rsEmployee("FCCODE")%>" Selected><%=rsEmployee("FCCODE")%> : <%=rsEmployee("FCNAME")%></option>
					<%Else%>
						<option value="<%=rsEmployee("FCCODE")%>"><%=rsEmployee("FCCODE")%> : <%=rsEmployee("FCNAME")%></option>
					<%End if
				rsEmployee.movenext
				loop
			End if
			
			rsEmployee.Close
			set rsEmployee = Nothing
			strEmployee = ""
		%>			
			</select>
		</td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</table>
	</form>
<%Elseif Request("txtKind") ="Absence" then%>
<br /><br />
	<form action="p_Absence.asp" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%if len(trim(request("txtYear"))) > 0 then
	     txtyear=request("txtYear")
	else
	     txtyear=year(Now)	 
	end if%>
	<%if len(trim(request("txtmonth"))) > 0 then
	     txtmonth=request("txtmonth")
	else
	     txtmonth=Month(Now)	 
	end if%>
	<table>
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type="text" name="txtyear" value="<%=txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type="text" name="txtmonth" value="<%=txtMonth%>"></td>
		<td><font face="verdana" size="1">Employee</font></td><td><select name="txtEmployeeCode" ><%cmdCallCombo "Employee","","",""%></select></td>
		<td><input type="submit" value="View" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana"/></td>
	</tr>
	</table>
	</form>	
<%elseif Request("txtKind") ="Employee" then%>
	<form action="d_PrintData.asp?txtTable=Employee" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <font face="Verdana" size="1">Select Data Here : <br><br>	
	<font face="Verdana" size="1">Code</font><input type="text" size="15" maxlength="30" name="FCCODE">
 	<font face="Verdana" size="1">Employee Name</font><input type="text" size="15" maxlength="30" name="FCNAME">
    <font face="Verdana" size="1">Gang</font>
	<select name="GANGCODE">
		<%cmdCallCombo "Gang","","",""%>
	</select>
	<input type="submit" value="Preview" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">	
	</form>
<%elseif Request("txtKind") ="Bank" then%>

	<form action="d_PrintData.asp?txtTable=Bank" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="CONTROLJOB">Control Job</Option>
		<Option value="ACCOUNTCODE">Account Code</Option>
		<Option value="ADDRESS">Address</Option>
		<Option value="CITY">City</Option>
		<Option value="STATE">State</Option>
		</select>
		&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="StockGroup" then%>
	
	<form action="d_PrintData.asp?txtTable=StockGroup" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Supplier" then%>
	
	<form action="d_PrintData.asp?txtTable=Supplier" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		<Option value="ACTIVATION">Active</Option>
		<Option value="ADDRESS">Address</Option>
		<Option value="CITY">City</Option>
		<Option value="STATE">State</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Store" then%>
	
	<form action="d_PrintData.asp?txtTable=Store" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Gang" then%>
	
	<form action="d_PrintData.asp?txtTable=Gang" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Workshop" then%>
	
	<form action="d_PrintData.asp?txtTable=Workshop" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Vehicle" then%>
	
	<form action="d_PrintData.asp?txtTable=Vehicle" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="VEHICLEGROUPCODE">Vehicle<br>Group Code</Option>
		<Option value="REGISTRATIONNO">Registration<br>No</Option>
		<Option value="ACTIVATION">Active</Option>
		<Option value="OWNERSHIP">Ownership</Option>
		<Option value="VEHICLESTATUS">Vehicle<br>Status</Option>
		<Option value="CONTROLALLOCATION">Control<br>Allocation</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Contractor" then%>
	
	<form action="d_PrintData.asp?txtTable=Contractor" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		<Option value="ACTIVATION">Active</Option>
		<Option value="ADDRESS">Address</Option>
		<Option value="CITY">City</Option>
		<Option value="STATE">State</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Building" then%>
	
	<form action="d_PrintData.asp?txtTable=Building" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="BUILDGROUPCODE_DESC">Building<br>Group Code</Option>
		<Option value="ITEM">Item</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Allowded" then%>
	
	<form action="d_PrintData.asp?txtTable=Allowded" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ALLDED_GROUP">Allowded<br>Group</Option>
		<Option value="ALLDED">Allowded</Option>
		<Option value="CROP">Crop</Option>
		<Option value="CONTROL_JOB">Control<br>Account</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Field" then%>
	
	<form action="d_PrintData.asp?txtTable=Field" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="DIVISION">Division</Option>
		<Option value="TERRAINTYPE">Terrain<br>Type</Option>
		<Option value="SOILTYPE">Soil Type</Option>		
		<Option value="OWNERSHIP">Ownership</Option>
		<Option value="ACTIVATION">Active</Option>
		<Option value="CROPTYPE">Crop<br>Type</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Station" then%>
	
	<form action="d_PrintData.asp?txtTable=Station" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="JOBCODE">Control<br>Account</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Machine" then%>
	
	<form action="d_PrintData.asp?txtTable=Machine" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="DESCRIPTION">Description</Option>
		<Option value="FIXEDASSETCODE">Fixex<br>Asset</Option>
		<Option value="OWNERSHIP">Ownership</Option>
		<Option value="STATION_STATIONCODE">Station<br>Code</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		</select>		
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="BusinessUnit" then%>
	
	<form action="d_PrintData.asp?txtTable=BusinessUnit" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="FCCOMPANYCODE">Company Code</Option>
		<Option value="FCCOMPANYNAME">Company Description</Option>
		<Option value="FCAREA">Area</Option>
		<Option value="FCOWNERSHIP">Ownership</Option>
		<Option value="FCMANAGER">Manager</Option>
		<Option value="FCTELEPHONE">Phone</Option>
		<Option value="FCFAX">Fax</Option>
		<Option value="FCEMAIL">e-Mail</Option>
		</select>		
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="UserProfile" then%>
	
	<form action="d_PrintData.asp?txtTable=UserProfile" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FULLNAME">Description</Option>
		<Option value="DEPARTMENT">Department</Option>
		<Option value="ACCESSRIGHTS">Access Right</Option>
		</select>	
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="AccessRights" then%>
	
	<form action="d_PrintData.asp?txtTable=AccessRights" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">User</Option>
		<Option value="MODULECODE">Module</Option>
		<Option value="SUBMODULECODE">Sub Module</Option>
		<Option value="AUTHORIZED">Authorized</Option>
		</select>		
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Trigger_Control" then%>
	
	<form action="d_PrintData.asp?txtTable=Trigger_Control" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="ACTIVITY">Activity</Option>
		<Option value="SOURCE">Source Type</Option>
		<Option value="TARGET">Target Type</Option>
		</select>		
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Elseif Request("txtKind") ="Job" then%>
	
	<form action="d_PrintData.asp?txtTable=Job" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="JOBGROUPCODE">Job Group<br>Code</Option>
		<Option value="JOB_FORCROPTYPE">Job For Type</Option>
		<Option value="JOB_FORFIELDSTATUS">Job For Status</Option>
		<Option value="JOB_CATEGORY">Job Category</Option>
		<Option value="UNITOFMEASUREMENT">Unit of Measurement</Option>
		<Option value="ACCOUNTCODE">Account<br>Code</Option>
		</select>		
			&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>		
	<%Elseif Request("txtKind") ="JobGroup" then%>
	
	<form action="d_PrintData.asp?txtTable=JobGroup" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="PARENTCODE">Parent<br>Code</Option>
		<Option value="PARENTDESC">Parent<br>Description</Option>
		</select>
				&nbsp;&nbsp;
	<input type=submit value="Enter" style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
	</form>
	<%Else%>
		<input type="text" title="Case is not Sensitive" name="txtField" value="<%=txtField%>" width="5">	
	<%End if%>
	
</body>
</html>