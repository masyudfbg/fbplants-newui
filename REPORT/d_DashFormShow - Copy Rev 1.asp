<%
	response.buffer = True
	validatesession
%>
<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<html>

<head>
<title>ePlaS Reporting</title>
<link href="/css/stylereport2.css" rel="stylesheet" >
<link rel="stylesheet" href="/css/stylereport.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<style>th,td { padding-top:5px; padding-bottom:5px;} table.showform{ min-width:500px;} a { font-size:12px;} 
.ckbox_list {
    border: 1px dotted #CCCCCC;
    margin-top: 5px;
    padding: 5px 10px;
}
.ckbox_list span {
    margin-left: 5px;
    margin-right: 5px;
    vertical-align: 3px;
}</style>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--
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
-->
</head>

<body>
<% 
dim xtitle


xtitle = ucase(Request("txtKind"))

 %>
<div id="siteInfo" style=" position:fixed; z-index:1; width:100%;">
    <div class="head" style="padding:5px;">
    	<h1 style="margin-top:15px;"><font color="#FFCC00" style="font-weight:normal;"><%=xtitle%></font></h1>
        <div style="color:#FFFFFF; font-size:10px;"><b><%=session("FULLNAME")%></b> <font size="2" color="#FFCC00">|</font> <b>BA : <%=Session("BA")%> ( <%=session("fcBA")%> )</b> <font size="2"  color="#FFCC00">|</font> Current Period :  <%=FormatDateTime(session("v_Costbook"),1)%>
		</div>
		
        
         <div id="globalNav">
		<%if Session("USERNAME")="" then%>
		<a class="mn_top" href="../logout.asp"><span>log-in</span></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Please login first...</B><a class="mn_top"  href="../login.asp"><i class="home_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Home</span></a></div>
		<%else%>
		<a class="mn_top" href="../logout.asp"><i class="off_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>log-out</span></a><a class="mn_top"  href="../login.Asp"><i class="home_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Home</span></a><!--<a id="button" class="mn_top_full" href="javascript:void(0);" onclick="toggleFull();"><i class="full">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Full Screen</span></a>--></div>
		<%end if%>
    </div>
    
   
   </div>
<br/><br/><br/><br/><br/>
<div id="container2" style="min-height:500px; border-bottom:1px #185C0A solid;">
 <div class="colR" style="height:100%;">
  <div id="logo-scroll">
      <div class="nav">
        <a class="breadcrumb" href="Dashboard.asp" style="display:table">&#139;&nbsp;Menu Laporan</a>
      </div>
      <br/>
  </div>
</div>      
<div class="line"></div>
<div class="colL">
<br/><br/>
<div style="border:1px #ccc solid; padding:10px; margin-right:15px; box-shadow:1px 1px 3px rgba(0,0,0,0.2);">
<div style="background:#185C0A; color:#fff; margin-top:-11px;margin-left:-11px;margin-right:-11px; margin-bottom:10px;padding:5px; padding-left:10px;"><font face="Verdana" size="1">Select Data Here :</font></div>
<%if Request("txtKind") ="ReceivedVoucher" then%>
    <form action="../p_PaymentVoucherDetail.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
     <table class="showform">
    <tr><td colspan="2"> </td></tr>
	<tr><td width="100"><font face="Verdana" size="1">From </td>
	<td width="300"><input class="datepicker" type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=TextDate(date())%>" title="Type Date mm/dd/yyyy" >
	<br></td>
	<tr><td>To</td>
	<td><input class="datepicker" type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=TextDate(date())%>"title="Type Date mm/dd/yyyy" >
	</td>
	<tr><td><font face="Verdana" size="1">Voucher Code</font></td>
	<td><input type="text" size="15" maxlength="30" name="VOUCHERCODE"  style=" background:#FFFFCC;"></td>
	<tr><td><font face="Verdana" size="1">Bank</font></td>
	<td><select name="BANKCODE" >
		<%cmdCallCombo "Bank","","",""%>
	</select></td>
    <tr><td><font face="Verdana" size="1">TT</font></td>
	<td><select name="LOCATIONTYPE" >
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select></td>
	<tr><td><font face="Verdana" size="1">TC</font></td>
	<td><input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC;"></td>
	<tr><td colspan="2"><input type="submit" value="Enter" ></td></tr></table>
<%Elseif Request("txtKind") ="PaymentVoucher" then%>
    <form action="../p_PaymentVoucherDetail.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table class="showform" style="width:100%;">
    <tr>
        <td><font face="Verdana" size="1">From</font></td>
	<td><input class="datepicker" type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<!--<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
	 </td>
    </tr>
    <tr>
        <td><font face="Verdana" size="1">To</font></td>
	<td><input class="datepicker" type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<!--<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
	</td>
    </tr>
    <tr>
        <td><font face="Verdana" size="1">Voucher Code</font></td>
        <td><input class="code" type="text" size="15" maxlength="30" name="VOUCHERCODE" >
	 </td>
    </tr>
    <tr>
        <td><font face="Verdana" size="1">Bank</font></td>
        <td>
	<select name="BANKCODE" >
		<%cmdCallCombo "Bank","","",""%>
	</select>
	</td>
    </tr>
    <tr>
	<!--'MODIFIED BY UKIE 07 NOV 2012 BY ADDING FILTER BY JOBCODE-->
	<td><font face="Verdana" size="1">Job</font></td><td><select name="JOBCODE" ><%cmdCallCombo "Job","","",""%></select>
     </tr>
    <tr>
        <td><font face="Verdana" size="1">Target Type</font></td>
        <td>
	<select name="LOCATIONTYPE" >
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	</td>
    </tr>
    <tr>
        <td><font face="Verdana" size="1">Target Code</font></td>
        <td>
	<input type="text" size="15" maxlength="30" name="ITEMCODE" ></td>
    </tr>
    </table>
    <br/><br/>
	<input type="submit" value="Enter" >
<%Elseif Request("txtKind") ="JournalReclass" then%>
    <form action="../p_JournalReclass.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table class="showform" style="width:100%;">
    <tr>
        <td><font face="Verdana" size="1">From</font></td>
	<td><input class="datepicker" type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<!--<a href="javascript:show_calendar('frmEntry.DATECREATED_From');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
	 </td>
    </tr>
    <tr>
        <td><font face="Verdana" size="1">To</font></td>
	<td><input class="datepicker" type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<!--<a href="javascript:show_calendar('frmEntry.DATECREATED_To');" onMouseOver="window.status='Date Picker';return true;" 
	onmouseout="window.status='';return true;">
	<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
	</td>
    </tr>
    <tr>
        <td><font face="Verdana" size="1">Voucher Code</font></td>
        <td><input class="code" type="text" size="15" maxlength="30" name="VOUCHERCODE" >
	 </td>
    </tr>
    <tr>
        <td><font face="Verdana" size="1">Bank</font></td>
        <td>
	<select name="BANKCODE" >
		<option value="88" <%if request("BANKCODE") = "88" then%> selected <%end if%>>JURNAL</option>
	</select>
	</td>
    </tr>
    <tr>
	<!--'MODIFIED BY UKIE 07 NOV 2012 BY ADDING FILTER BY JOBCODE-->
	<td><font face="Verdana" size="1">Job</font></td><td><select name="JOBCODE" ><%cmdCallCombo "Job","","",""%></select>
     </tr>
    <tr>
        <td><font face="Verdana" size="1">Target Type</font></td>
        <td>
	<select name="LOCATIONTYPE" >
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	</td>
    </tr>
    <tr>
        <td><font face="Verdana" size="1">Target Code</font></td>
        <td>
	<input type="text" size="15" maxlength="30" name="ITEMCODE" ></td>
    </tr>
    </table>
    <br/><br/>
	<input type="submit" value="Enter" >
<%Elseif Request("txtKind") ="ArusCash" then%>
    <form action="d_ArusCash.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table class="showform">
	<tr>
	   <td>
	      <font face="Verdana" size="1">From</font> 
	   </td>
	   <td >
	      <input class="datepicker" type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=TextDate(date())%>" title="Type Date mm/dd/yyyy" >
	   </td>
	</tr>
	<tr>
	   <td>
	       <font face="Verdana" size="1">To</font>
	   </td>
	   <td>
	      <input class="datepicker" type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=TextDate(date())%>"title="Type Date mm/dd/yyyy" >
	   </td>
	</tr>
	 <tr>
        <td >
            <font face="Verdana" size="1">Accounting Period</font>
        </td>
							
			<% 	Dim rsComboA,Data_ACCOUNTINGYEARPERIOD
				q = "select accountingyear||lpad(accountingperiod,2,0) ACCOUNTINGYEARPERIOD,Description from  ACCOUNTINGYEARPERIOD where fcBA='" &  Session("fcBA") & "' order by accountingyear||lpad(accountingperiod,2,0)"

					Set rsComboA = Server.CreateObject("ADODB.Recordset")
					'response.write q
					rsComboA.CursorType = 3
					rsComboA.CursorLocation = 3
					rsComboA.Open q, connstr
					if not rsComboA.eof then%>
					<td align="left"> 
						<select name="ACCOUNTINGYEARPERIOD">    
										<% Do While Not rsComboA.eof %>
										<% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then 
											Data_ACCOUNTINGYEARPERIOD = "true"
										else
											Data_ACCOUNTINGYEARPERIOD = "false"
										End if%>
										<option value="<%=rsComboA("ACCOUNTINGYEARPERIOD")%>" <% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then %> Selected <%End if%>><%=rsComboA("Description")%></option> 
										<%  rsComboA.movenext
										   loop
									if Data_ACCOUNTINGYEARPERIOD = "false" then %>
										<option value="<%=ACCOUNTINGYEAR&ACCOUNTINGPERIOD%>" Selected>Tahun <%=ACCOUNTINGYEAR%> Bulan <%=ACCOUNTINGPERIOD%></option>
									<% end if %>
									
									<%
										End if
										rsComboA.Close
										set rsComboA = Nothing
										q=""
										%>
						</select> <font color="#FF0000">*</font>
					</td>
    </tr>
	<tr>
		<td>
		<font face="Verdana" size="1">Bank</font>
		</td>
		<td>
		  <select name="BANKCODE" >
			<%cmdCallCombo "Bank","","",""%>
		  </select>
		</td>
	</tr>
	<tr>
	    <td colspan="2"><input type="submit" value="Enter" ></td>
	</tr>
  </table>
<%Elseif Request("txtKind") ="PayslipSummary" then%>
    <form action="d_PayslipSummary.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table width="400">
    <tr><td colspan="2"> </td></tr>
	<tr><td width="100"><font face="Verdana" size="1">From </td>
	<td width="300"><input class="datepicker" type="text" maxlength="10" name="DATECREATED_From" size="10" value="<%=TextDate(date())%>" title="Type Date mm/dd/yyyy" >
	<br></td>
	<tr><td>To</td>
	<td><input class="datepicker" type="text" maxlength="10" name="DATECREATED_To" size="10" value="<%=TextDate(date())%>"title="Type Date mm/dd/yyyy" >
	</td>
	<tr><td><font face="Verdana" size="1">Employee Code</font></td>
	<td><input type="text" size="15" maxlength="30" name="Empcode"  style="background:#FFFFCC;"></td>
	<tr><td><font face="Verdana" size="1">Gang</font></td>
	<td><select name="GANGCODE" >
		<%cmdCallCombo "GANG","","",""%>
	</select></td>
	<tr><td colspan="2"><input type="submit" value="Enter" ></td></tr></table>
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
	<td><input type="text" size="15" maxlength="30" name="Empcode"  style="background:#FFFFCC;"></td>
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
	<tr><td colspan="2"><input type="submit" value="Enter" ></td></tr></table>

<%Elseif Request("txtKind") ="StockItem" or Request("txtKind") ="Stock" then%>
   <form action="../PrintData.asp?txtTable=Stock" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="verdana" size="1">Search :</font></td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr><tr><td><font face="verdana" size="1">from Column : </font></td><td>
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
	</td></tr><tr><td colspan="2">
	<input type=submit value="Enter" >
    </td></tr></table>
	</form>
<%elseif Request("txtKind") ="PurchaseRequest" then%>
	<form action="../p_PurchaseRequestDetails.asp" method="post" target="_blank" name="frmEntry">
	<table class="showform">
		<tr>
		    <td>
				<font face="Verdana" size="1">PR Date</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_From');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="PRDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
				<!--<a href="javascript:show_calendar('frmEntry.PRDATE_To');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">PR Number</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30"  name="PRCODE"  style="background:#FFFFCC;">
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">Item Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC;">
			</td>
		</tr>
		
		<tr>
			<td>
				<font face="Verdana" size="1">Requestor</font>
			</td>
			<td>
				<select name="REQUESTORID">
					<%cmdCallCombo "REQUESTOR","","",""%>
				</select>
			</td>
		</tr>

	
		<tr>
			<td>
				<font face="Verdana" size="1">Status</font>
			</td>
			<td>				
				<select name="ROWSTATEID">
					<%cmdCallCombo "rowstatePurchaseRequest", "" ,"",""%>
				</select>
			</td>
		</tr>			
	</table>	
	
	
	
    <input type="submit" value="Preview" >

<%elseif Request("txtKind") ="PurchaseOrderList" or Request("txtKind")="POOutstanding" then%>
    <%if Request("txtKind")="POOutstanding" then%>
        <form action="../p_PurchaseOrderDetails.asp?txtKind=POOutstanding" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <%Else%>
        <form action="../p_PurchaseOrderDetails.asp?txtKind=PurchaseOrderList" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%End if%>
	<table class="showform"><tr><td>
    <font face="verdana" size="1">From</font>
     </td><td>
	<input class="datepicker" type="text" maxlength="10" name="PODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
    </td></tr><tr><td>
	<font face="verdana" size="1">To</font>
    </td><td>
	<input class="datepicker" type="text" maxlength="10" name="PODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
    </td></tr><tr><td>
	<font face="Verdana" size="1">PO Number</font>
    </td><td>
	<input type="text" size="15" maxlength="30"  name="POCODE"  style="background:#FFFFCC;">
     </td></tr><tr><td>
	<font face="Verdana" size="1">Item Code</font>
    </td><td>
	<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC;">
    </td></tr><tr><td>
	<font face="Verdana" size="1">Supplier</font></td><td>
	<select name="SUPPLIERCODE">
		<%cmdCallCombo "SUPPLIER","","",""%>
	</select>
    </td></tr><tr><td colspan="2">
    <input type="submit" value="Enter" >
    </td></tr></table>
    
<%elseif Request("txtKind") ="ReceiveNote" then%>
	
	<form action="../p_ReceiveNote.asp" method="post" target="_blank" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="Verdana" size="1">From</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="RNDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	 </td></tr><tr><td>
	<font face="verdana" size="1">To</font>
    </td><td>
	<input class="datepicker"  type="text" maxlength="10" name="RNDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
    </td></tr><tr><td>
	<font face="Verdana" size="1">RN Number</font> </td><td><input type="text" size="15" maxlength="30"  name="RECEIVENOTECODE"  style="background:#FFFFCC;">
    </td></tr><tr><td>
	<font face="Verdana" size="1">Item Code</font>	</td><td><input type="text" size="15" maxlength="30" name="PURCHASEITEMCODE"  style="background:#FFFFCC;">
    </td></tr><tr><td>
    <font face="Verdana" size="1">	Supplier</font></td><td>
	<select name="SUPPLIERCODE">
		<%cmdCallCombo "SUPPLIER","","",""%>
	</select>
    </td></tr>
	
	<tr>
			<td>
				<font face="Verdana" size="1">Jobcode</font>
			</td>
			<td>
				<select name="jobcode">
					<%cmdCallCombo "JOB",Request("jobcode"),"",""%>
				</select>
			</td>
		</tr>
	
	<tr><td colspan="2">
	<input type="submit" value="Enter" >
    </td></tr></table>
<%elseif Request("txtKind") ="Invoice" then%>
	
	<form action="../p_Invoice.asp" method="post" target="_blank" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="Verdana" size="1">From</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="INVOICEDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	<font face="verdana" size="1">To</font>
    </td><td>
	<input class="datepicker" type="text" maxlength="10" name="INVOICEDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	<font face="Verdana" size="1">Supplier</font>
    </td><td>
	<select name="SUPPLIERCODE">
		<%cmdCallCombo "SUPPLIER","","",""%>
	</select>
    </td></tr><tr><td>
	<font face="Verdana" size="1">Invoice Number</font> </td><td><input type="text" size="15" maxlength="30"  name="INVOICECODE"  style="background:#FFFFCC;">
    </td></tr><tr><td>
	<font face="Verdana" size="1">Item Code</font>	</td><td><input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC;">
    </td></tr><tr><td colspan="2">
	<input type="submit" value="Enter" >
    </td></tr></table>
    
<%Elseif Request("txtKind") ="Purchasing_Closing" then%>
    <b>Purchasing Close</b></a><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br>
	<input type="submit" value="::CLOSE::" >
<%elseif Request("txtKind") ="IssueDetails" or Request("txtKind")="StockIssue" then%>
	<%if Request("txtKind")="IssueDetails" then%>
        <form action="../p_IssueDetails.asp?txtKind=IssueDetails" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <%Else%>
        <form action="/p_IssueDetails.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%End if%>

    <table class="showform"><tr><td>
	<font face="Verdana" size="1">From</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="SIVDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	<font face="Verdana" size="1">To</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="SIVDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	<font face="Verdana" size="1">Issue No</font></td><td><input type="text" size="15" maxlength="30"  name="SIVCODE"  style="background:#FFFFCC;">
    </td></tr><tr><td>
	<font face="Verdana" size="1">Stock Code</font></td><td><input type="text" size="15" maxlength="30" name="STOCKCODE"  style="background:#FFFFCC;">
    </td></tr><tr><td>
	<!--'MODIFIED BY UKIE 07 NOV 2012 BY ADDING FILTER BY JOBCODE-->
	<font face="Verdana" size="1">Job</font></td><td><select name="JOBCODE" ><%cmdCallCombo "Job","","",""%></select>
    </td></tr><tr><td>
    <font face="Verdana" size="1">TT</font></td><td>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
    </td></tr><tr><td>
	<font face="Verdana" size="1">TC</font></td><td><input type="text" size="5" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
    </td></tr><tr><td colspan="2">
	<input type="submit" value="Enter" >
    </td></tr></table>
<%elseif Request("txtKind") ="WriteOff" then%>
	<form action="../p_WriteOff.asp" method="post" target="_blank" name="frmEntry">
		<table class="showform">
			<tr>
				<td><font face="Verdana" size="1">From</font></td>
				<td><input class="datepicker" type="text" maxlength="10" name="WODATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" ></td>
			</tr>
			<tr>
				<td><font face="Verdana" size="1">To</font></td>
				<td><input class="datepicker" type="text" maxlength="10" name="WODATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" ></td>
			</tr>
			<tr>
				<td><font face="Verdana" size="1">Writeoff Code</font></td>
				<td><input type="text" size="15" maxlength="30"  name="WriteOffCode"  style="background:#FFFFCC;"></td>
			</tr>
			<tr>
				<td><font face="Verdana" size="1">Item Code</font></td>
				<td><input type="text" size="15" maxlength="30" name="STOCKCODE"  style="background:#FFFFCC;"></td>
			</tr>
			<tr>
				<td><font face="Verdana" size="1">TT</font></td>
				<td>
					<select name="LOCATIONTYPE">
						<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
					</select>
				</td>
			</tr>
			<tr>
				<td><font face="Verdana" size="1">TC</font></td>
				<td><input type="text" size="5" maxlength="30" name="LocationCode"  style="background:#FFFFCC;"></td>
			</tr>
			<tr>
				<td colspan="2"><input type="submit" value="Enter" ></td>
			</tr>
		</table>
<%Elseif Request("txtKind") ="StockLedger" then%>
	<form action="../p_stockledger.asp?txtKind=StockLedger" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table class="showform">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear value="<%=txtYear%>"></td>
    </tr>
    <tr>   
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth value="<%=txtMonth%>"></td>
    </tr>
    <tr>
		<td><font face="verdana" size="1">Store Code</font></td><td><input type=text name=txtstorecode></td>
    </tr>
    <tr>
		<td><font face="verdana" size="1">STOCKGROUP</font></td><td><input type=text name=txtstockgroup></td>
    </tr>
    <tr>
		<td><font face="verdana" size="1">STOCK</font></td><td><input type=text name=txtstock></td>
    </tr>
    <tr>
		<td colspan="2"><input type=submit value="Enter" ></td>
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
	<input type="submit" value="::CLOSE::" >
<%Elseif Request("txtKind") ="EmployeeList" then%>
    <form action="../PrintData.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
	<tr>
		<td><font face="verdana" size="1">CODE</font></td><td><input type=text name=txtcode value="<%=txtcode%>"></td>
		<td><font face="verdana" size="1">EmployeeName</font></td><td><input type=text name=txtemployeename value="<%=txtemployeename%>"></td>
		<td><font face="verdana" size="1">Gang</font></td><td><input type=text name=txtGangCode value="<%=txtGangCode%>"></td>
		<td><input type=submit value="Enter" ></td>
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
      <a href="javascript:show_calendar('frmEntry.FDDATE_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;"> 
      <img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
    <font face="Verdana" size="1">To</font> <input type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" > 
      <a href="javascript:show_calendar('frmEntry.FDDATE_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;"> 
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
    <TD colspan="3"><font face="Verdana" size="1">TC</font> <input type="text" size="15" maxlength="30" name="LOCATIONCODE"  style="background:#FFFFCC;">
    <font face="Verdana" size="1">Report</font> <select name="Report" >
        <option value="ATTENDANCE">ATTENDANCE</option>
      </select><input type="submit" value="Enter" ></TD>
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
	<a href="javascript:show_calendar('frmEntry.INVOICEDATE_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">Supplier</font>
	<select name="SUPPLIERCODE">
		<%cmdCallCombo "SUPPLIER","","",""%>
	</select>
	<font face="Verdana" size="1">Invoice Number</font> <input type="text" size="15" maxlength="30"  name="INVOICECODE"  style="background:#FFFFCC;">
	<font face="Verdana" size="1">Item Code</font>	<input type="text" size="15" maxlength="30" name="ITEMCODE"  style="background:#FFFFCC;">
	<input type="submit" value="Enter" >
</table>

<%Elseif Request("txtKind") ="EmployeeActivity" then%>
    <form action="../p_Attendance_GAD.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
<script>
	function checkAll() {
		var checkbox_options = document.getElementById("checkbox_options");
		var checkbox_all = document.getElementById("checkbox_all");

		var checkboxes = checkbox_options.getElementsByTagName("input");        
		for (i = 0; i < checkboxes.length; i++) {
			checkboxes[i].checked = checkbox_all.checked;
		}
	}
</script>
	<table>
  <tbody><tr> 
    <td colspan="2">
		<div class="ckbox_list">
			<table>
				<tr>
					<td><input id="checkbox_all" type="checkbox" onclick="checkAll()" /><span><font face="Verdana" size="1">Select All</Font></span></td>
				</tr>
			</table>
			<div id = "checkbox_options" >
				<table>
					<tr>
						<td><input name="cField" type="checkbox" value="fddate" /><span><font face="Verdana" size="1">Date</font></span></td>
						<td><input name="cField" type="checkbox" value="gangcode" /><span><font face="Verdana" size="1">Gang Code</font></span></td>
						<td><input name="cField" type="checkbox" value="supervision_1" /><span><font face="Verdana" size="1">Supervision 1</Font></span></td>
						<td><input name="cField" type="checkbox" value="supervision_2" /><span><font face="Verdana" size="1">Supervision 2</Font></span></td>
						<td><input name="cField" type="checkbox" value="supervision_3" /><span><font face="Verdana" size="1">Supervision 3</Font></span></td>
						<td><input name="cField" type="checkbox" value="supervision_4" /><span><font face="Verdana" size="1">Supervision 4</Font></span></td>
					</tr>
					<tr>
						<td><input name="cField" type="checkbox" value="sectionname" /><span><font face="Verdana" size="1">Section Name</font></span></td>
						<td><input name="cField" type="checkbox" value="employeecode" /><span><font face="Verdana" size="1">Employee Code</Font></span></td>
						<td><input name="cField" type="checkbox" value="employeename"/><span><font face="Verdana" size="1">Employee Name</Font></span></td>
						<td><input name="cField" type="checkbox" value="attendance" /><span><font face="Verdana" size="1">Attendance</Font></span></td>
						<td><input name="cField" type="checkbox" value="jobcode"/><span><font face="Verdana" size="1">Job Code</Font></span></td>
						<td><input name="cField" type="checkbox" value="jobname" /><span><font face="Verdana" size="1">Job Description</Font></span></td>
					</tr>
					<tr>
						<td><input name="cField" type="checkbox" value="locationtype"/><span><font face="Verdana" size="1">Target Type</Font></span></td>
						<td><input name="cField" type="checkbox" value="locationcode"/><span><font face="Verdana" size="1">Target Code</Font></span></td>
						<td><input name="cField" type="checkbox" value="reference" /><span><font face="Verdana" size="1">Reference</Font></span></td>
						<td><input name="cField" type="checkbox" value="remarks"/><span><font face="Verdana" size="1">Remarks</Font></span></td>
					</tr>
				</table>
			</div>
		</div>
       </td>
  </tr>
  <tr> 
    <td ><br><font size="1" face="Verdana">From </font></td><td>
	<input class="datepicker" type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="FDDATE_From" maxlength="10"> 
   </td></tr>
   <tr><td>  
    <font size="1" face="Verdana">To</font></td><td>
     <input class="datepicker" type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="FDDATE_To" maxlength="10"> 
      </td></tr>
   <tr><td> 
	  <font size="1" face="Verdana">Gang</font></td><td>
      <select name="GANGCODE">  
	 <%cmdCallCombo "GANG","","",""%>
      </select>
	  </td>
  </tr>
	<tr>
		<td><font face="Verdana" size="1">Section Name</font></td>
		<td>
			<select name="txtsectionname">
				<option value=""></option>
				<%
				Dim rssectionname, strsectionname
				strsectionname	=	" select fccode, fcname from parameter where parcode = 'E0002' and fcba = '" & Session("fcba") & "' order by fccode asc"
				
				Set rssectionname = Server.CreateObject("ADODB.Recordset")
				rssectionname.Open strsectionname, connstr, adOpenStatic 
				
				if not rssectionname.eof then
				Do While Not rssectionname.eof
				if rssectionname("fccode") = v_txtsectionname then%>
						<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
					 <%Else%>
						<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
					<%End if
					rssectionname.movenext
				loop
				end if
				rssectionname.Close
				set rssectionname = Nothing
				%>
			</select>
		</td>
	</tr>
  <tr> 
    <td><font size="1" face="Verdana">Job</font></td><td>
     <select name="JOBCODE">
        
	 <%cmdCallCombo "JOB","","",""%>
		
      </select>
      </td></tr>
   <tr><td> 
	  <font size="1" face="Verdana">Employee</font></td><td> <select name="EMPLOYEECODE">
        
	 <%cmdCallCombo "EMPLOYEE","","",""%>
		
      </select>
      </td></tr>
   <tr><td> 
	  <font size="1" face="Verdana">TT</font> </td><td><select name="LOCATIONTYPE">
        
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
    <td ><font size="1" face="Verdana">TC</font> </td><td><input type="text" style="background:#FFFFCC;" name="LOCATIONCODE" maxlength="30" size="15">
    </td></tr>
   <tr><td> 
    <font size="1" face="Verdana">Report</font> </td><td><select name="Report">
        <option value="ACTIVITY">ACTIVITY</option>\
	</select>
    </td></tr>
   <tr><td colspan="2"> 
    <input type="submit"  value="Enter"></td>
  </tr>
</tbody>
</table>
</form>
	
<%elseif Request("txtKind") ="EmpAllded" then%>
<form action="../p_EmpAllded.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<script>
		function checkAll() {
			var checkbox_options = document.getElementById("checkbox_options");
			var checkbox_all = document.getElementById("checkbox_all");

			var checkboxes = checkbox_options.getElementsByTagName("input");        
			for (i = 0; i < checkboxes.length; i++) {
				checkboxes[i].checked = checkbox_all.checked;
			}
		}
	</script>
<table width="100%">
  <tbody><tr> 
    <td colspan="2">
				<div class="ckbox_list">
					<table>
						<tr>
							<td><input id="checkbox_all" type="checkbox" onclick="checkAll()" /><span><font face="Verdana" size="1">Select All</Font></span></td>
						</tr>
					</table>
					<div id = "checkbox_options" >
						<table>
							<tr>
								<td><input name="cField" type="checkbox" value="adhocdate" /><span><font face="Verdana" size="1">Date</font></span></td>
								<td><input name="cField" type="checkbox" value="employeecode" /><span><font face="Verdana" size="1">Employee Code</Font></span></td>
								<td><input name="cField" type="checkbox" value="employeename"/><span><font face="Verdana" size="1">Employee Name</Font> </span></td>
								<td><input name="cField" type="checkbox" value="gangcode" /><span><font face="Verdana" size="1">Gang Code</Font></span></td>
							</tr>
						</table>
					</div>
				</div>
       </td>
  </tr>
	<tr>
	<td><font face="Verdana" size="1">From</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
  </tr>
	<tr>
	<td>
	<font face="Verdana" size="1">To</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
  </tr>
	<tr>
	<td>
    <font face="Verdana" size="1">Gang</font>
    </td><td>
	<select name="GANGCODE">
		<%cmdCallCombo "Gang","","",""%>
	</select></td>
  </tr>
	<tr>
	<td>
    <font face="Verdana" size="1">Allowance</font></td><td>
	<select name="Allowded" >
		<%cmdCallCombo "Allowded","","",""%>
	</select>
    </td>
  </tr>
	<tr>
	<td>
    <font face="Verdana" size="1">Employee</font></td><td>
	<select name="EMPLOYEECODE" >
		<%cmdCallCombo "EMPLOYEE","","",""%>
	</select>
    </td>
  </tr>
	<tr>
	<td colspan="2">
	<input type="submit" value="Enter" >
	</td>
	</tr>
	</Table>
<%Elseif Request("txtKind") ="Calculate" then%>
    <b>Calculate Payroll</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" >
<%Elseif Request("txtKind") ="Employee_Closing" then%>
    <b>Employee Close</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::CLOSE::" >
<%elseif Request("txtKind") ="HarvestingQuality" then%>
	<form action="../p_HarvestingQuality.asp" method="post" target="_blank" name="frmEntry">
    <table width="100%"><tr>
    <td>
    <font face="Verdana" size="1">Select Data Here : <br><br>
    </td>
    <td>
    <div class="ckbox_list">
    
	<input name="cField" type="checkbox" value="FDDATE" /><span><font face="Verdana" size="1">Tanggal</font></span>
	<input name="cField" type="checkbox" value="EMPCODE" /><span><font face="Verdana" size="1">NIK</Font></span>
	<input name="cField" type="checkbox" value="FIELDCODE"/><span><font face="Verdana" size="1">Blok</Font></span>
    
	</div>
    
    </td>
    </tr>
    <tr>
    <td>
	<font face="Verdana" size="1">Form</font>
    </td>
    <td>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="FDDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>
    <tr>
    <td>
    <font face="Verdana" size="1">Employee</font>
    </td>
    <td>
	<select name="EMPLOYEECODE" >
		<%cmdCallCombo "EMPLOYEE","","",""%>
	</select>
    </td>
    </tr>
	<tr><td colspan="2">
	<input type="submit" value="Enter" >
    </td></tr></table>	
	
<%elseif Request("txtKind") ="WorkshopList" then%>
	<form action="../p_WorkshopList.asp?txtKind=WorkshopList" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table>
	<tr>
	<br>
	<font size="1" face="Verdana"> 
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_From" maxlength="10">
	<a onMouseOut="window.status='';return true;" onMouseOver="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
	<font size="1" face="Verdana">To</font>
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_To" maxlength="10">
	<a onMouseOut="window.status='';return true;" onMouseOver="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');">
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
	<font size="1" face="Verdana">TC</font><input type="text" style="background:#FFFFCC;" name="LocationCode" maxlength="15" size="5">
	<input type="submit"  value="Enter">
			
		</font>
	</tr>
	</table>
    <input  type="hidden" value="Workshop" name="txtTable">
	<input  type="hidden" value="Data Master" name="txtModule">
	
<%elseif Request("txtKind") ="WorkshopRunningAccount" then%>
	<form action="../p_WorkshopActivity.asp?txtKind=WorkshopRunningAccount" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table width="100%">
	<tr><td>
	<font size="1" face="Verdana"> From</font></td><td>
	<input class="datepicker" type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_From" maxlength="10">
	</td></tr><tr><td>
	<font size="1" face="Verdana">To</font></td><td>
	<input class="datepicker" type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_To" maxlength="10">
	</td></tr><tr><td>
    <font size="1" face="Verdana">Activity</font></td><td>
	<select name="JOB">
		
	 <%cmdCallCombo "JOB","","",""%>
		
	</select>
    </td></tr><tr><td>
    <font size="1" face="Verdana">TT</font></td><td>
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
    </td></tr><tr><td>
	<font size="1" face="Verdana">TC</font></td><td><input type="text" style="background:#FFFFCC;" name="LocationCode" maxlength="15" size="5">
    </td></tr><tr><td colspan="2">
	<input type="submit"  value="Enter">
    </td>
			
		</font>
	</tr>
	</table>
    <input  type="hidden" value="Workshop" name="txtTable">
	<input  type="hidden" value="Data Master" name="txtModule">
	
<%Elseif Request("txtKind") ="Workshop_Closing" then%>
    <b>Workshop Close</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::CLOSE::" >
	
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
		<input type="submit"  value="Enter">
			</font>
	</font>
	
	<input  type="hidden" value="Vehicle" name="txtTable">
	<input  type="hidden" value="Data Master" name="txtModule">

<%Elseif Request("txtKind") ="VehicleRunningAccount" then%>
	<form action="../p_VehicleActivity.asp" method="post" target="_blank" name="frmEntry">
	<table width="100%">
    <tr><td>
	<font size="1" face="Verdana"> From</font></td><td>
	<input class="datepicker" type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="VEHDATE_From" maxlength="10">
	</td></tr><tr><td>
	<font size="1" face="Verdana">To</font></td><td>
	<input class="datepicker" type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="VEHDATE_To" maxlength="10">
	</td></tr><tr><td>
    <font size="1" face="Verdana">Vehicle</font></td><td>
	<select name="VEHICLECODE"><%cmdCallCombo "VEHICLE","","",""%>
		
	</select>
    </td></tr><tr><td>
    <font size="1" face="Verdana">Activity</font></td><td>
	<select name="JOBCODE"><%cmdCallCombo "JOB","","",""%>
		
	</select>
    </td></tr><tr><td>
    <font size="1" face="Verdana">TT</font></td><td>
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
    </td></tr><tr><td>
	<font size="1" face="Verdana">TC</font></td><td><input type="text" style="background:#FFFFCC;" name="LocationCode" maxlength="30" size="15">
    </td></tr>
		<tr>
	<td>
    <font face="Verdana" size="1">Category Transport</font>
    </td>
	<td>
		<select name="categorytransport">
			<option value=""> --All-- </option>
		<% 
			q = " SELECT fccode, fcname FROM vehicleactivitycategory ORDER BY LPAD(fccode, 2, 0)"
			set rsCombo = Server.CreateObject("ADODB.Recordset")
			'response.write q
			rsCombo.CursorType = 3
			rsCombo.CursorLocation = 3
			rsCombo.Open q, connstr
			if not rsCombo.eof then
				Do While Not rsCombo.eof
		%>
			<option value="<%=rsCombo("fccode")%>" <% If Request("categorytransport") = rsCombo("fccode") Then %> Selected <%End if%>><%=rsCombo("fccode")%> : <%=rsCombo("fcname")%></option> 
		<%    		rsCombo.movenext
				loop						
			End if
			rsCombo.Close
			set rsCombo = Nothing
			q=""
		%>
		</select>
    </td>
    </tr>
	
	<tr><td colspan="2">
    <input type="submit"  value="Enter">
     </td></tr></table>
			<!-- Akhir Tampilan Select form -->
			
		
	
	
<%Elseif Request("txtKind") ="VehicleOperation" then%>
	<form action="../p_VehicleOperation.asp" method="post" target="_blank" name="frmEntry">
	<table width="100%">
    <tr>
		<td><font size="1" face="Verdana"> From</font></td>
		<td><input class="datepicker" type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="VEHDATE_From" maxlength="10"></td>
	</tr>
	<tr>
		<td><font size="1" face="Verdana">To</font></td>
		<td><input class="datepicker" type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="VEHDATE_To" maxlength="10"></td>
	</tr>
	<tr>
		<td><font size="1" face="Verdana">Group</font></td>
		<td><select name="VEHICLEGROUPCODE"><%cmdCallComboSetup "VEHICLEGROUP","VEHICLEGROUPCODE","fcba",session("fcba")%></select></td>
	</tr>
	<tr>
		<td><font size="1" face="Verdana">Vehicle</font></td>
		<td><select name="VEHICLECODE"><%cmdCallCombo "VEHICLE","","",""%></select></td>
	</tr>
	<tr>
		<td><font size="1" face="Verdana">Vehicle</font></td>
		<td>
			<select name="vehiclestatus">
				<option value="" <% If Request("vehiclestatus") = "" Then %> Selected <%End if%>> --All-- </option>
				<%
				q = " SELECT fccode, fcname FROM parameter where fcba = '" &  Session("fcBA") & "' AND parcode = 'VO001'"
				Set rsComboA = Server.CreateObject("ADODB.Recordset")
				
				rsComboA.Open q, connstr
				if not rsComboA.eof then
					Do While Not rsComboA.eof
				%>
					<option value="<%=rsComboA("fccode")%>" <% If cstr("" & Request("vehiclestatus")) = cstr(rsComboA("fccode")) Then %> Selected <%End if%>><%=rsComboA("fcname")%></option> 
				<% 
						rsComboA.movenext
					loop
				End if
				rsComboA.Close : set rsComboA = Nothing : q = ""
				%>
			
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="2"><input type="submit"  value="Enter"></td>
	</tr>
			
		
	
<%Elseif Request("txtKind") ="Vehicle_Closing" then%>
    <b>Vehicle Close</b></a><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br>
	<input type="submit" value="::CLOSE::" >
	
<%elseif request("txtkind") ="ContractorList" then%>
	<form action="../PrintData.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<tr>
	<br>
	<font size="1" face="Verdana"> 
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_From" maxlength="10">
	<a onMouseOut="window.status='';return true;" onMouseOver="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
	<font size="1" face="Verdana">To</font>
	<input type="text" title="Type Date mm/dd/yyyy" value="<%=textDate(Date)%>" size="10" name="DATEACTIVITY_To" maxlength="10">
	<a onMouseOut="window.status='';return true;" onMouseOver="window.status='Date Picker';return true;" href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');">
	<img width="24" height="22" border="0" align="top" src="../Images/show-calendar.gif"></a>
    <font size="1" face="Verdana">Code</font><input type="text" style="background:#FFFFCC;" name="" maxlength="30" size="15">
	<input type="submit"  value="Enter">

	</font></font>
</tr>	
	<input  type="hidden" value="Contractor" name="txtTable">
	<input  type="hidden" value="Data Master" name="txtModule">

<%elseif Request("txtKind") ="ContractAgreement" then%>
	<form action="../p_AGREEMENTDETAIL.asp?" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <!--<table width="500px"><tr><td>
	<font size="1" face="Verdana">From</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	<font size="1" face="Verdana">To</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	 <font size="1" face="Verdana">Code</font></td><td><input type="text" style="background:#FFFFCC;" name="" maxlength="30" size="15">
     </td></tr><tr><td colspan="2">
	 <input type="submit"  value="Enter">
     </td></tr></table>-->

 <table class="showform" style="width:100%;">
 		<tr>
			<td><font face="Verdana" size="1">Agreement Code</font></td>
			<td>
				<input type="text" maxlength="30" name="AGREEMENTCODE" size="30" value="<%=AGREEMENTCODE%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contractor</font>
			</td>
			<td>
				<select name="CONTRACTORCODE" >
					<option value = ""<%If Len(Trim(CONTRACTORCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					Dim objectRs, objectSQL
					objectSQL	=	" SELECT FCCODE, FCNAME FROM CONTRACTOR WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If CONTRACTORCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Due Date From</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATE_FROM" size="10" value="<%=DATE_FROM%>"title="Type Date mm/dd/yyyy" >
		
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATE_TO" size="10" value="<%=DATE_TO%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Line No</font></td>
			<td>
				<input type="text" maxlength="10" name="LINENO" size="10" value="<%=LINENO%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Activity</font>
			</td>
			<td>
				<select name="JOBCODE" >
					<option value = ""<%If Len(Trim(JOBCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If JOBCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
    </table>
	
<%elseif Request("txtKind") ="AGREEMENTDETAIL" then%>
	
	<br>
	<img src="ICON/Contract.jpg" border="0" ><br>
	<b>Maintenance Contract Agreement</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
    <font face="Verdana" size="1">Activity</font>
	<select name="JOB" >
		<%cmdCallCombo "job","","",""%>
	</select>
    <font face="Verdana" size="1">TT</font>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<font face="Verdana" size="1">TC</font><input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
	<input type="submit" value="Enter" >
<%elseif Request("txtKind") ="WorkinProgress" then%>
	<form action="../p_WIP.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table width="100%">
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATE_FROM" size="10" value="<%=DATE_FROM%>"title="Type Date mm/dd/yyyy" >
		
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATE_TO" size="10" value="<%=DATE_TO%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Agreement Code</font></td>
			<td>
				<input type="text" maxlength="30" name="AGREEMENTCODE" size="30" value="<%=AGREEMENTCODE%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contractor</font>
			</td>
			<td>
				<select name="CONTRACTORCODE" >
					<option value = ""<%If Len(Trim(CONTRACTORCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM CONTRACTOR WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If CONTRACTORCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Line No</font></td>
			<td>
				<input type="text" maxlength="10" name="LINENO" size="10" value="<%=LINENO%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Activity</font>
			</td>
			<td>
				<select name="JOBCODE" >
					<option value = ""<%If Len(Trim(JOBCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If JOBCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
	</table>
	
<%elseif Request("txtKind") ="wipplan" then%>
	<form action="../p_WIPPLAN.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table width="100%">
		<tr>
			<td><font face="Verdana" size="1">From</font></td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATE_FROM" size="10" value="<%=DATE_FROM%>"title="Type Date mm/dd/yyyy" >
		
				<font face="Verdana" size="1">To</font>
				<input class="datepicker" type="text" maxlength="10" name="DATE_TO" size="10" value="<%=DATE_TO%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Agreement Code</font></td>
			<td>
				<input type="text" maxlength="30" name="AGREEMENTCODE" size="30" value="<%=AGREEMENTCODE%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contractor</font>
			</td>
			<td>
				<select name="CONTRACTORCODE" >
					<option value = ""<%If Len(Trim(CONTRACTORCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM CONTRACTOR WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If CONTRACTORCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Line No</font></td>
			<td>
				<input type="text" maxlength="10" name="LINENO" size="10" value="<%=LINENO%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Activity</font>
			</td>
			<td>
				<select name="JOBCODE" >
					<option value = ""<%If Len(Trim(JOBCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If JOBCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
	</table>

<%elseif Request("txtKind") ="wipfinal" then%>
	<form action="../p_wipfinal.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table width="100%">
		<tr>
			<td><font face="Verdana" size="1">Year</font></td>
			<td>
				<select name="txtyear">
				<%
					Yr = Year(Session("v_Contract"))
					mon = Month(Session("v_Contract"))
					MinYr = Yr - 2
					MaxYr = Yr + 2
					
					if len(v_txtyear) = 0 then v_txtyear = Yr
					if len(v_txtmonth) = 0 then v_txtmonth = mon
				%>
				<%For i = MinYr to MaxYr%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtyear) then%> Selected <%End If%>><%=i%></option>
				<%Next%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Month</font></td>
			<td>
				<select name="txtmonth">
				<%For i = 1 to 12%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtmonth) then%> Selected <%End If%>><%=textMonth(i)%></option>
				<%Next%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Agreement Code</font></td>
			<td>
				<input type="text" maxlength="30" name="AGREEMENTCODE" size="30" value="<%=AGREEMENTCODE%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Contractor</font>
			</td>
			<td>
				<select name="CONTRACTORCODE" >
					<option value = ""<%If Len(Trim(CONTRACTORCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM CONTRACTOR WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If CONTRACTORCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td><font face="Verdana" size="1">Line No</font></td>
			<td>
				<input type="text" maxlength="10" name="LINENO" size="10" value="<%=LINENO%>" title="Enter Contract Number" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Activity</font>
			</td>
			<td>
				<select name="JOBCODE" >
					<option value = ""<%If Len(Trim(JOBCODE)) = 0 then%> Selected <%End If%>>--ALL--</option>
					<%
					objectSQL	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'" & _
									" ORDER BY FCCODE"
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr

					If Not objectRS.EOF Then
						Do While Not objectRS.EOF %>
							<option value = "<%=objectRS("FCCODE")%>" <%If JOBCODE = objectRS("FCCODE") Then%> Selected <%End If%> ><%=objectRS("FCCODE")%> : <%=objectRS("FCNAME")%></option> 
						<%
							objectRS.movenext
						Loop
					End If
					objectRS.Close
					set objectRS = Nothing
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="submit" value="Preview" >
			</td>
		</tr>
	</table>
	
<%Elseif Request("txtKind") ="Contract_Closing" then%>
    <b>Contract Close</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::CLOSE::" >
	
<%elseif Request("txtKind") ="NurseryList" then%>
	<form action="../p_Nursery.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry"><br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Batch</font><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;">
 	<font face="Verdana" size="1">Crop</font><input type="text" size="15" maxlength="30" name="Crop"  style="background:#FFFFCC;">
 	<font face="Verdana" size="1">Source</font><input type="text" size="15" maxlength="30" name="Source"  style="background:#FFFFCC;">
	<input type="submit" value="Enter" >
	
<%elseif Request("txtKind") ="NurseryOrder" then%>
	
	<br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Nursery Order</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Batch</font><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;">
	<input type="submit" value="Enter" >
	
<%elseif Request("txtKind") ="NurseryCulling" then%>
	<form action="../p_NurseryCulling.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table class="showform">
		<tr>
			<td>
				<font face="Verdana" size="1">From</font></td><td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">To</font></td><td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Document Number</font>
			</td>
			<td>
				<input type="text" size="50" maxlength="50" name="txtNurseryNo"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Batch</font></td><td><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;"></td></tr><tr><td colspan="2">
				<input type="submit" value="Enter" >
			</td>
		</tr>
	</table>
	
<%elseif Request("txtKind") ="NurseryAdjustment" then%>
	<form action="../p_NurseryAdjustment.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table class="showform">
		<tr>
			<td>
				<font face="Verdana" size="1">From</font></td><td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">To</font></td><td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Document Number</font>
			</td>
			<td>
				<input type="text" size="50" maxlength="50" name="txtNurseryNo"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Batch</font></td><td><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;"></td></tr><tr><td colspan="2">
				<input type="submit" value="Enter" >
			</td>
		</tr>
	</table>
<%elseif Request("txtKind") ="NurseryIssue" then%>
	<form action="../p_NurseryIssue.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table class="showform">
		<tr>
			<td>
				<font face="Verdana" size="1">From</font>
			</td>
			<td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">To</font></td><td>
				<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Document Number</font>
			</td>
			<td>
				<input type="text" size="50" maxlength="50" name="txtNurseryNo"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Batch</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Job</font>
			</td>
			<td>
				<select name="txtJobCode" >
					<%cmdCallCombo "Job",Request("txtJobCode"),"",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td>
				<select name="LOCATIONTYPE">
					<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<font face="Verdana" size="1">Target Code</font>
			</td>
			<td>
				<input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
			</td>
		</tr>
		<tr>
			<td colspan="2">
				<input type="submit" value="Enter" >
			</td>
		</tr>
	</table>
<%elseif Request("txtKind") ="NurseryLedger" then%>
    <form action="../p_NurseryLedger.asp?txtKind=NurseryList" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
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
    <table class="showform">
    <tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear value="<%=txtyear%>"></td>
    </tr>
    <tr>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth value="<%=txtmonth%>"></td>
    </tr>
    <tr>
		 <td><font face="Verdana" size="1">Batch</font></td><td><input type="text" size="15" maxlength="30" name="txtnurserycode"  style="background:#FFFFCC;"></td>
     </tr>
    <tr>    
    <td colspan="2"><input type="submit" value="Enter"></td>
    </tr>
    </table>
	
<%Elseif Request("txtKind") ="Nursery_Closing" then%>
    <b>Nursery Close</b></a>
	<input type="submit" value="::CLOSE::" >
	
<%elseif Request("txtKind") ="ProjectList" then%>
	 <form action="../p_Project.asp?txtKind=ProjectList" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	 <br>
	<br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#FFFFCC;">
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
 	<font face="Verdana" size="1">Source</font><input type="text" size="15" maxlength="30" name="Source"  style="background:#FFFFCC;">
	<input type="submit" value="Enter" >
<%elseif Request("txtKind") ="Project_Production" then%>
	
	<br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Production</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#FFFFCC;">
	<input type="submit" value="Enter" >
<%elseif Request("txtKind") ="Project_Adjustment" then%>
	<br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Adjusment</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#FFFFCC;">
	<input type="submit" value="Enter" >
<%elseif Request("txtKind") ="Project_Issue" then%>
	<br>
	<img src="ICON/Nursery.jpg" border="0" ><br>
	<b>Maintenance Project Issue</b></a>
    <br>
     
	<font face="Verdana" size="1"> 
	<input type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.DATEACTIVITY_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="ProjectCode"  style="background:#FFFFCC;">
	<font face="Verdana" size="1">LC</font>
	<select name="LOCATIONTYPE">
		<%cmdCallComboSetup "LOCATIONTYPE","LOCATIONTYPECODE","",""%>
	</select>
	<font face="Verdana" size="1">TC</font><input type="text" size="15" maxlength="30" name="LocationCode"  style="background:#FFFFCC;">
	<input type="submit" value="Enter" >
	
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
	<font face="Verdana" size="1">Project Code</font><input type="text" size="15" maxlength="30" name="txtprojectcode"  style="background:#FFFFCC;">
    <input type="submit" value="Enter" >
	
<%Elseif Request("txtKind") ="Project_Close" then%>
    <b>Project Close</b></a>
	<input type="submit" value="::CLOSE::" >
<%Elseif Request("txtKind") ="Building_Closing" then%>
    <b>Building Close</b></a>
	<input type="submit" value="::CLOSE::" >
<%Elseif Request("txtKind") ="Building_Closing" then%>
    <b>Building Close</b></a>
	<input type="submit" value="::CLOSE::" >
<%Elseif Request("txtKind") ="General_Charges_Closing" then%>
    <b>General Charge Close</b></a>
	<input type="submit" value="::CLOSE::" >
<%Elseif Request("txtKind") ="Harvesting_Closing" then%>
    <b>Close Harvesting</b></a><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br>
	<input type="submit" value="::CLOSE::" >
<%Elseif Request("txtKind") ="Financial_Monthly_Closing" then%>
    <b>Closing This Month</b></a>
	<input type="submit" value="::CLOSE::" >
<%Elseif Request("txtKind") ="Previous_Month_UnClosing" then%>
    <b>un-Closing This Month</b></a>
	<input type="submit" value="::CLOSE::" >
<%Elseif Request("txtKind") ="Financial_Year_Closing" then%>
    <b>Close Financial Year</b></a>
	<input type="submit" value="::CLOSE::" >
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
	<table class="showform">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear value="<%=txtyear%>"></td>
    <tr>
    </tr>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth value="<%=txtmonth%>"></td>
    <tr>
    </tr>
		<td><font face="verdana" size="1">JOBGROUP</font></td><td><input type=text name=txtjobgroup></td>
    <tr>
    </tr>
		<td><font face="verdana" size="1">JOB</font></td><td><input type=text name=txtjob></td>
    <tr>
    </tr>
		<td colspan="2"><input type=submit value="Enter" ></td>
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
	<table class="showform">
	 <tr>
		<td width="100"><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear width="4" value="<%=txtYear%>"></td>
     
		<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtMonth width="3" value="<%=txtMonth%>"></td></tr>
     <tr>
		<td><font face="verdana" size="1">JOBGROUP</font></td><td><input type=text name=txtjobgroup width="6"></td>
		<td><font face="verdana" size="1">JOB</font></td><td><input type=text name=txtjob width="10"></td></tr>
     <tr>
		<td><font face="verdana" size="1">ST</font></td><td><input type=text name=txtST width="5"></td>
		<td><font face="verdana" size="1">SC </font></td><td><input type=text name=txtSC width="10"></td></tr>
    <tr>
		<td><font face="verdana" size="1">TT</font></td><td><input type=text name=txtTT width="5"></td>
		<td><font face="verdana" size="1">TC</font></td><td><input type=text name=txtTC width="10"></td></tr>
	<tr>
		<td><font face="verdana" size="1">Reference</font></td><td><input type=text name=txtreference width="4" value="<%=txtReference%>"></td>
		<td><font face="verdana" size="1">Reference 01</font></td><td><input type="text" name="txtreference01" width="4" value=""></td>
	</tr>
	<tr>
		<td><font face="verdana" size="1">Reference 02</font></td><td><input type="text" name="txtreference02" width="4" value=""></td>
		<td><font face="verdana" size="1">Reference 03</font></td><td><input type="text" name="txtreference03" width="4" value=""></td>
	</tr>
	<tr>
		<td><font face="verdana" size="1">Reference 04</font></td><td><input type="text" name="txtreference04" width="4" value=""></td>
		<td><font face="verdana" size="1">Reference 05</font></td><td><input type="text" name="txtreference05" width="4" value=""></td>
	</tr>
	<tr>
		<td><font face="verdana" size="1">Debit Credit</font></td>
		<td>
			<select name="txtDebitCredit" style="width:148px;">
				<option value="" selected></option>
				<option value="Db" >Debet</option>
				<option value="Cr" >Credit</option>
			</select>
		
		</td>
		
	</tr>
<tr>
		<td><br/><input type=submit value="Enter" ></td>
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
		<td><input type=submit value="Enter" ></td>
	</tr>
	</table>
	</form>
<%elseif Request("txtKind") ="Harvesting" then%>
<form action="../p_Harvesting.asp" method="post" target="_blank" name="frmEntry">
	<table class="showform">
	<tr>
	<td><font face="Verdana" size="1"> From</font></td><td>
    <input class="datepicker" type="text" maxlength="10" name="RECEPTIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td><font face="Verdana" size="1">To</font></td><td>
    <input class="datepicker" type="text" maxlength="10" name="RECEPTIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" ></td></tr><tr><td>
    <font face="Verdana" size="1">Chit No</font></td>
	<td><input type="text" size="15" maxlength="30" name="CHITNO"></td>
	</tr>
	<tr>
	<td><font face="Verdana" size="1">Field Code</font></td>
	<td><select name="FIELDCODE">
	<%cmdCallCombo "Field","FIELDCODE","",""%>
	</select></td></tr><tr><td>
    <font face="Verdana" size="1">Cropcode Code</font></td>					
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
	%></select>
    </td></tr><tr><td>
    <font face="Verdana" size="1">Product Code</font></td>
	<td><select name="ProductCODE">
		<%cmdCallComboSetup "PRODUCT","ProductCODE","",""%>
		</select>
	</td></tr><tr><td colspan="2">
    <input type="submit" value="Enter" ></td>
	</tr>
	</table>
</form>	
<%elseif Request("txtKind") ="HarvestingRound" then%>
     <form action="../p_HarvestingRound.asp" method="post" target="_blank" name="frmEntry">
	<table class="showform">
	<tr>
	<td><font face="verdana" size="1">Year</font></td><td><input type=text size="10" maxlength="4" name=HRNDYEAR>
    </td></tr><tr><td>
    <font face="verdana" size="1">Month</font></td><td><input type=text size="10" maxlength="2" name=HRNDMONTH>
    </td></tr><tr><td>
    <font face="Verdana" size="1">Field Code</font></td>
	<td><select name="FIELDCODE">
	<%cmdCallCombo "Field","FIELDCODE","",""%>
	</select></td></tr><tr><td>
    <font face="Verdana" size="1">Cropcode Code</font></td>					
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
	</td></tr><tr><td colspan="2">
    <input type="submit" value="Enter" >	</td>
	</tr>
	</table>
<%elseif Request("txtKind") ="Weather" then%>
	<form action="../p_Weather.asp" method="post" target="_blank" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="Verdana" size="1">From</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="WIDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	<font face="Verdana" size="1">To</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="WIDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
 	<font face="Verdana" size="1">Division</font></td><td><input type="text" size="15" maxlength="30" name="WIDIVISION">
    </td></tr><tr><td colspan="2">
	<input type="submit" value="Enter" >
    </td></tr></table>
	
<%elseif Request("txtKind") ="TBSPriceList" then%>
    
	<form action="../p_TBSPriceList.asp" method="post" target="_blank" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="Verdana" size="1">From</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="DATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	<font face="Verdana" size="1">To</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="DATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr>
	<tr><td>
	<font face="Verdana" size="1">Product Code</font></td>
	<td face="Verdana" size="1"  >
		<select name="PRODUCT_PRODUCTCODE">
		<%cmdCallComboSetup "PRODUCT","ProductCODE","FCBA",session("fcba")%>
		</select>
    </td>
	</tr>
	<tr><td colspan="2">
	<input type="submit" value="Enter" >
    </td></tr></table>
	
	
<%elseif Request("txtKind") ="PRODUCT" then%>
	
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
	<input type="submit" value="Enter" >	
<%elseif Request("txtKind") ="WEIGHBRIDGE" then%>
	
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Weightbridge</b></a>
    <br>
     
 	<font face="Verdana" size="1">Weighbridge Code</font><input type="text" size="15" maxlength="30" name="WEIGHBRIDGECODE">
	<input type="submit" value="Enter" >	

<%elseif Request("txtKind") ="CUSTOMER" then%>
	
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Customer</b></a>
    <br>
     
 	<font face="Verdana" size="1">Customer Code</font><input type="text" size="15" maxlength="30" name="CUSTOMERCODE">
	<input type="submit" value="Enter" >	
<%elseif Request("txtKind") ="GRADINGITEM" then%>
	
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Grading Item</b></a>
    <br>
     
 	<font face="Verdana" size="1">Item Code</font><input type="text" size="15" maxlength="30" name="ITEMCODE">
	<input type="submit" value="Enter" >	
<%elseif Request("txtKind") ="OPHarvRate" then%>
	
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
	<td><input type="submit" value="Enter" ></td>
	</tr>
	</table>

<%elseif Request("txtKind") ="WEIGHBRIDGETICKET" then%>
<form action="../p_WEIGHBRIDGETICKET.asp" method="post" target="_blank" name="frmEntry">
	<!--ADDED BY UKIE 14-NOV-2012-->
    <table class="showform"><tr><td>
	<font face="Verdana" size="1">From</font></td>
	<td style="width:100px;">
	<input class="datepicker" type="text" maxlength="10" name="Dateout_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td><td style="width:30px;">
	<font face="Verdana" size="1">To</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="Dateout_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr>
	<tr>
    <td>
	<font face="Verdana" size="1">Posting Date</font>
    </td>
    <td  colspan="3" >
	<input class="datepicker" type="text" maxlength="10" name="Posting_Date" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
    </td>
	</tr>
	<tr><td>
	<!--END ADDING-->
 	<font face="Verdana" size="1">Product Code</font></td>
	<td  colspan="3" >
		<select name="PRODUCT_PRODUCTCODE">
		<%cmdCallComboSetup "PRODUCT","ProductCODE","FCBA",session("fcba")%>
		</select>
        </td></tr><tr><td>
 	<font face="Verdana" size="1">Chit No</font></td>
	<td colspan="3" ><input type="text" size="15" maxlength="30" name="CHITNO">
    </td></tr>
	<tr><td >
 	<font face="Verdana" size="1">Customer Code</font></td>
	<td  colspan="3" >
		<select name="CUST_CUSTOMERCODE">
		<%cmdCallComboSetup "CUSTOMER","CUSTOMERCODE","FCBA",session("fcba")%>
		</select>
        </td></tr><tr><td colspan="2">
	<input type="submit" value="Enter" >
    </td></tr></table>	
</form>
	
<%elseif Request("txtKind") ="GRADINGRECORD" then%>
<form action="../p_GRADINGRECORD.asp" method="post" target="_blank" name="frmEntry">
<table class="showform"><tr><td>
	<font face="Verdana" size="1">From</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="DATETIME_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	<font face="Verdana" size="1">To</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="DATETIME_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
 	<font face="Verdana" size="1">Grading No</font></td><td><input type="text" size="15" maxlength="30" name="GRADINGNO"></td></tr><tr><td>
 	<font face="Verdana" size="1">Vehicle No</font></td><td><input type="text" size="15" maxlength="30" name="VEHICLENO"></td></tr><tr><td>
 	<font face="Verdana" size="1">Chitno</font></td><td><input type="text" size="15" maxlength="30" name="WBT_CHITNO">
    </td></tr><tr><td colspan="2">
	<input type="submit" value="Enter" >
    </td></tr></table>
</form>
<%elseif Request("txtKind") ="SHIFT" then%>
	
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Shift</b></a>
    <br>
     
 	<font face="Verdana" size="1">Shift Code</font><input type="text" size="15" maxlength="30" name="SHIFTCODE">
	<input type="submit" value="Enter" >	
<%elseif Request("txtKind") ="SHIFTREPORT" then%>
	<form action="../p_SHIFTREPORT.asp" method="post" target="_blank" name="frmEntry">
		<table class="showform">
		 <tr>
		   <td>
			<font face="Verdana" size="1">From</font>
		   </td>
		   <td>
			<input class="datepicker" type="text" maxlength="10" name="PRODUCTIONDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		   </td>
		 </tr>
		 <tr>
		   <td>
			<font face="Verdana" size="1">To</font>
		   </td>
		   <td>
			<input class="datepicker" type="text" maxlength="10" name="PRODUCTIONDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
		   </td>
		 </tr>
		 <tr>
		   <td>
			<font face="Verdana" size="1">Shift Code</font>
		   </td>
		   <td>
			<select name="SHIFT_SHIFTCODE"><%cmdCallComboSetup "SHIFT","SHIFT_SHIFTCODE","fcba",session("fcba")%></select>
		   </td>
		 </tr>
		 <tr>
		    <td colspan="2">
			<input type="submit" value="Enter" >
			</td>
		 </tr>
		</table>
	</form>
<%elseif Request("txtKind") ="PRODUCTSTORAGE" then%>
	
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Product Storage</b></a>
    <br>
     
 	<font face="Verdana" size="1">Storage Code</font><input type="text" size="15" maxlength="30" name="STORAGECODE">
 	<font face="Verdana" size="1">Product Code</font>
		<select name="PRODUCT_PRODUCTCODE">
		<%cmdCallComboSetup "PRODUCT","PRODUCT_PRODUCTCODE","FCBA",session("fcba")%>
		</select>
	<input type="submit" value="Enter" >	

<%elseif Request("txtKind") ="PRODUCTSTORAGEREADING" then%>
<form action="../p_PRODUCTSTORAGEREADING.asp" method="post" target="_blank" name="frmEntry">
<table class="showform" ><tr><td>
	<font face="Verdana" size="1">From</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="READINGDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
	<font face="Verdana" size="1">To</font></td><td>
	<input class="datepicker" type="text" maxlength="10" name="READINGDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td></tr><tr><td>
 	<font face="Verdana" size="1">Storage Code</font></td><td><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","",""%></select>	
    </td></tr><tr><td colspan="2">
	<input type="submit" value="Enter" >
    </td></tr></table>
</form>
<%elseif Request("txtKind") ="PRODUCTADJUSTMENT" then%>
	
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Product Adjusment</b></a>
    <br>
     
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="ADJUSTMENTDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.ADJUSTMENTDATE_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="ADJUSTMENTDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.ADJUSTMENTDATE_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Adjusment No</font><input type="text" size="15" maxlength="30" name="ADJUSTMENTNO">
 	<font face="Verdana" size="1">Storage Code</font><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","FCBA",session("fcba")%></select>
	<input type="submit" value="Enter" >	

<%elseif Request("txtKind") ="PRODUCTSTORAGETRANSFER" then%>
	
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Storage Transfer</b></a>
    <br>
     
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="TRANSFERDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.TRANSFERDATE_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="TRANSFERDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.TRANSFERDATE_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Transfer No</font><input type="text" size="15" maxlength="30" name="TRANSFERNO">
 	<font face="Verdana" size="1">Storage Code</font><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","",""%></select>
	<input type="submit" value="Enter" >	

<%elseif Request("txtKind") ="DELIVERYCONTRACT" then%>
<form action="../p_DELIVERYCONTRACT.asp" method="post" target="_blank" name="frmEntry">
	<TABLE class="showform">
	<TR><TD>
    <font face="Verdana" size="1">From</FONT></TD><TD><input class="datepicker" type="text" maxlength="10" name="STARTDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</TD></TR><TR><TD>
    <font face="Verdana" size="1">To</font></TD><TD><input class="datepicker" type="text" maxlength="10" name="STARTDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</TD></TR><TR><TD><font face="Verdana" size="1">Contract Code</font></TD><TD><input type="text" size="15" maxlength="30" name="CONTRACTCODE"></TD></TR><TR><TD><font face="Verdana" size="1">Customer Code</font></TD><TD><select name="CUST_CUSTOMERCODE"><%cmdCallComboSetup "CUSTOMER","CUSTOMERCODE","",""%></select>
    </TD></TR><TR><TD><font face="Verdana" size="1">Product Code</font></TD><TD><select name="PRODUCT_PRODUCTCODE"><%cmdCallComboSetup "PRODUCT","ProductCODE","FCBA",session("fcba")%></select>
    </TD></TR><TR><TD colspan="2"><input type="submit" value="Enter" ></TD>	
	</TR>
	</TABLE>
</form>	
<%elseif Request("txtKind") ="DELIVERYNOTE" then%>
<form action="../p_DELIVERYNOTE.asp" method="post" target="_blank" name="frmEntry">
	<TABLE class="showform">
	<TR>
	<TD><font face="Verdana" size="1">From</FONT></TD><TD><input class="datepicker" type="text" maxlength="10" name="DNOTE_DATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</TD></TR><TR><TD>
    <font face="Verdana" size="1">To</font></TD><TD><input class="datepicker" type="text" maxlength="10" name="DNOTE_DATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</TD></TR><TR><TD>
    <font face="Verdana" size="1">Delivery NOte Code</font></TD><TD><input type="text" size="15" maxlength="30" name="DELIVERYNOTECODE">
	</TD></TR><TR><TD>
    <font face="Verdana" size="1">Delivery Note Code</font></TD><TD><input type="text" size="15" maxlength="30" name="DELIVERYNOTECODE">
    </TD></TR><TR><TD>
    <font face="Verdana" size="1">Contract Code</font></TD>
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
			</select></TD></TR><TR><TD colspan="2">
            <input type="submit" value="Enter" ></TD>	
	</TR>
	</TABLE>
</form>
<%elseif Request("txtKind") ="DELIVERYRETURNNOTE" then%>
	
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Delivery Return Note</b></a>
    <br>
     
	<table>
	<tr>
	<td><font face="Verdana" size="1">From</font></td>
	<td><input type="text" maxlength="10" name="RETURNDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.RETURNDATE_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
	<td><font face="Verdana" size="1">To</font></td>
	<td><input type="text" maxlength="10" name="RETURNDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.RETURNDATE_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td>
 	<td><font face="Verdana" size="1">Delivery Note Code</font></td><td><input type="text" size="15" maxlength="30" name="DNOTE_DELIVERYNOTECODE"></td>
	</tr>
	<tr> 
 	<td><font face="Verdana" size="1">WB Chit No</font></td><td><input type="text" size="15" maxlength="30" name="WBT_CHITNO"></td>
 	<td><font face="Verdana" size="1">Storage Code</font></td><td><select name="PSTORAGE_STORAGECODE"><%cmdCallComboSetup "PRODUCTSTORAGE","PSTORAGE_STORAGECODE","FCBA",session("fcba")%></select></td>
	<td><input type="submit" value="Enter" ></td>	
	</tr>
	</table>
<%elseif Request("txtKind") ="DeliveryQuality" then%>
	
	<img src="IMAGES/add2.gif" border="0">
	<b>Maintenance Delivery Product Quality</b></a>
    <br>
     
	<font face="Verdana" size="1">From
	<input type="text" maxlength="10" name="QUALITYDATE_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.QUALITYDATE_From');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
	<font face="Verdana" size="1">To</font>
	<input type="text" maxlength="10" name="QUALITYDATE_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	<a href="javascript:show_calendar('frmEntry.QUALITYDATE_To');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
	<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
 	<font face="Verdana" size="1">Delivery Note Code</font><input type="text" size="15" maxlength="30" name="DELIVERYNOTE">
	<input type="submit" value="Enter" >		
<%Elseif Request("txtKind") ="Stock_Product" then%>
    <b>Stock Product Ledger</b><img src="ICON/StockLedger.jpg" border="0" ></a>
	
	<table>
	<tr>
		<td><font face="verdana" size="1">Year</font></td><td><input type=text name=txtyear></td>
		<td><font face="verdana" size="1">Month</font></td><td><input type=text name=txtmonth></td>
		<td><font face="verdana" size="1">Storage Code</font></td><td><input type=text name=txtSTORAGE_CODE></td>
		<td><font face="verdana" size="1">Product Code</font></td><td><input type=text name=txtPRODUCT_CODE></td>
		<td><input type=submit value="Enter" ></td>
	</tr>
	</table>
	
<%Elseif Request("txtKind") ="Close_Delivery" then%>
    <b>Close Delivery</b></a><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br>
	<input type="submit" value="::CLOSE::" >
<%Elseif Request("txtKind") ="Calculate_Mill_Product" then%>
    <b>Calculate Delivery</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" >
<%Elseif Request("txtKind") ="WeighBridge_Closing" then%>
    <b>WeighBridge Closing </b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" >
<%Elseif Request("txtKind") ="Production_Closing" then%>
    <b>Production Closing </b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" >
<%Elseif Request("txtKind") ="Delivery_Closing" then%>
    <b>Delivery Closing</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Calculate::" >
<%Elseif Request("txtKind") ="Clone" then%>
    <b>Clone Data Master to All Business Unit</b><br>
	<img src="ICON/Close.jpg" border="0" width="80" height="80"><br></a>
	<input type="submit" value="::Clone::" >
<%Elseif Request("txtKind") ="InvoiceLedger" then%>
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
	<table class="showform">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type="text" name="txtyear" value="<%=txtYear%>"></td>
    </tr>
    <tr>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type="text" name="txtmonth" value="<%=txtMonth%>"></td>
    </tr>
    <tr>
		<td><font face="verdana" size="1">SUPPLIER CODE</font></td><td><select name="txtsuppliercode" ><%cmdCallCombo "Supplier","","",""%></select></td>
	</tr>
    <tr>
    <td><font face="verdana" size="1">INVOICE CODE</font></td><td><input type="text" name="txtinvoicecode" /></td>
	</tr>
    <tr>
    <td colspan="2"><input type="submit" value="View"  /></td>
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
	<table class="showform">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td><input type="text" name="txtyear" value="<%=txtYear%>"></td>
    </tr>
    <tr>
		<td><font face="verdana" size="1">MONTH</font></td><td><input type="text" name="txtmonth" value="<%=txtMonth%>"></td>
    </tr>
    <tr>
		<td><font face="verdana" size="1">CONTRACTOR CODE</font></td><td><input type="text" name="txtcontractorcode" /></td>
    </tr>
    <tr>
		<td><font face="verdana" size="1">AGREEMENT CODE</font></td><td><input type="text" name="txtagreementcode" /></td>
	</tr>
    <tr>
    	<td><font face="verdana" size="1">AGREEMENT LINE NO</font></td><td><input type="text" name="txtagreementlineno" /></td>
	</tr>
    <tr>
    	<td colspan="2"><input type="submit" value="View"  /></td>
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
		<td><input type="submit" value="View"  /></td>
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
		<td><input type="submit" value="View" /></td>
	</tr>
	</table>
	</form>	
<%elseif Request("txtKind") ="Employee" then%>
	<form action="../PrintData.asp?txtTable=Employee" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <font face="Verdana" size="1">Select Data Here : <br><br>	
	<font face="Verdana" size="1">Code</font><input type="text" size="15" maxlength="30" name="FCCODE">
 	<font face="Verdana" size="1">Employee Name</font><input type="text" size="15" maxlength="30" name="FCNAME">
    <font face="Verdana" size="1">Gang</font>
	<select name="GANGCODE">
		<%cmdCallCombo "Gang","","",""%>
	</select>
	<input type="submit" value="Preview" >	
	</form>
<%elseif Request("txtKind") ="Bank" then%>

	<form action="../PrintData.asp?txtTable=Bank" method="post" target="_blank" style= 
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
	<input type=submit value="Enter" >
	</form>
	<%Elseif Request("txtKind") ="StockGroup" then%>
	
	<form action="../PrintData.asp?txtTable=StockGroup" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	 <table class="showform"><tr><td>
    <font face="verdana" size="1">Search : </font> </td><td> <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr><tr><td><font face="verdana" size="1">from Column</font></td><td>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		</select>
		</td></tr>
        <tr><td>
	<input type=submit value="Enter" >
    </td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="Supplier" then%>
	
	<form action="../PrintData.asp?txtTable=Supplier" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table class="showform"><tr><td>
    <font face="verdana" size="1">Search :</font></td><td> <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr><tr><td><font face="verdana" size="1">&nbsp;from Column</font></td><td>
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
</td></tr>
<tr><td colspan="2">
	<input type=submit value="Enter" >
    </td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="Store" then%>
	
	<form action="../PrintData.asp?txtTable=Store" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="verdana" size="1">Search : </font> </td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr><tr><td><font face="verdana" size="1">from Column</font></td><td>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
</td></tr><tr><td colspan="2">
	<input type=submit value="Enter" >
</td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="Gang" then%>
	
	<form action="../PrintData.asp?txtTable=Gang" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
     <table class="showform"><tr><td>
	<font face="verdana" size="1">Search : </font></td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr><tr><td>
    <font face="verdana" size="1">from Column</font></td><td>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
	</td></tr>
    <tr><td colspan="2">
	<input type=submit value="Enter" >
    </td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="Workshop" then%>
	
	<form action="../PrintData.asp?txtTable=Workshop" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="verdana" size="1">Search : </font></td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr><tr><td><font face="verdana" size="1">from Column</font></td><td>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
</td></tr>
<tr><td colspan="2">
	<input type=submit value="Enter" >
    </td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="Vehicle" then%>
	
	<form action="../PrintData.asp?txtTable=Vehicle" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="verdana" size="1">Search : </font></td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr><tr><td><font face="verdana" size="1">from Column</font></td><td>
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
</td></tr>
<tr><td colspan="2">
	<input type=submit value="Enter" >
    </td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="Contractor" then%>
	
	<form action="../PrintData.asp?txtTable=Contractor" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="verdana" size="1">Search :</font> </td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr><tr><td><font face="verdana" size="1">from Column</font></td><td>
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
</td></tr>
<tr><td colspan="2">
	<input type=submit value="Enter" >
     </td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="Building" then%>
	
	<form action="../PrintData.asp?txtTable=Building" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
    <table class="showform"><tr><td>
	<font face="verdana" size="1">Search :</font></td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5">
   </td></tr><tr><td> <font face="verdana" size="1">from Column</font></td><td>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="BUILDGROUPCODE_DESC">Building<br>Group Code</Option>
		<Option value="ITEM">Item</Option>
		<Option value="ACTIVATION">Active</Option>
		</select>
</td></tr>
<tr><td colspan="2">
	<input type=submit value="Enter" >
     </td></tr></table>
	</form>
	
	<%Elseif Request("txtKind") ="Nursery" then%>
	
	<form action="../p_Nursery.asp" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table class="showform">
   <!-- <tr>
    <td><font face="Verdana" size="1">From</font></td>
	<td>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_From" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	
	<font face="Verdana" size="1">To</font>
	<input class="datepicker" type="text" maxlength="10" name="DATEACTIVITY_To" size="10" value="<%=textDate(Date)%>"title="Type Date mm/dd/yyyy" >
	</td>
    </tr>-->
    <tr>
    <td>
 	<font face="Verdana" size="1">Batch</font></td>
	<td><input type="text" size="15" maxlength="30" name="NurseryCode"  style="background:#FFFFCC;">
    </td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Crop</font></td>
	<td><input type="text" size="15" maxlength="30" name="Crop"  style="background:#FFFFCC;">
    </td>
    </tr>
    <tr>
    <td>
 	<font face="Verdana" size="1">Source</font></td>
	<td><input type="text" size="15" maxlength="30" name="Source"  style="background:#FFFFCC;">
    </td>
    </tr>
    <tr>
    <td colspan="2" align="left">
	<input type="submit" value="Preview" >
    </td>
    </tr>
    </table>
	</form>
	
	
	<%Elseif Request("txtKind") ="Allowded" then%>
	
	<form action="../PrintData.asp?txtTable=Allowded" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table style="min-width:400px;"><tr><td>
    <font face="verdana" size="1">Search : </font></td><td> <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr>
    <tr><td><font face="verdana" size="1">from Column</font></td><td>
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="ALLDED_GROUP">Allowded<br>Group</Option>
		<Option value="ALLDED">Allowded</Option>
		<Option value="CROP">Crop</Option>
		<Option value="CONTROL_JOB">Control<br>Account</Option>
		</select>
 </td></tr>
         <tr><td colspan="2">
	<input type=submit value="Enter" >
     </td></tr></table>
	</form>
	
	<%Elseif Request("txtKind") ="GCMASTER" then%>
	
	<form action="../PrintData.asp?txtTable=GCMASTER" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table style="min-width:400px;"><tr><td>
    <font face="verdana" size="1">Search : </font></td><td> <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr>
    <tr><td><font face="verdana" size="1">from Column</font></td><td> 
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="ITEM">Item</Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="GCGROUP">Gc<br>Group</Option>
		<Option value="ACTIVATION">Activation</Option>
		</select>
 </td></tr>
         <tr><td colspan="2">
	<input type=submit value="Enter" >
    </td></tr></table>
	</form>
	
	<%Elseif Request("txtKind") ="Field" then%>
	
	<form action="../PrintData.asp?txtTable=Field" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table style="min-width:400px;"><tr><td>
    <font face="verdana" size="1">Search : </font></td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr>
    <tr><td><font face="verdana" size="1">from Column</font></td><td> 
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
 </td></tr>
         <tr><td colspan="2">
         <input type=submit value="Enter" >
         </td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="Station" then%>
	
	<form action="../PrintData.asp?txtTable=Station" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table style="min-width:400px;"><tr><td>
    <font face="verdana" size="1">Search :</font></td><td><input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr>
    <tr><td><font face="verdana" size="1">from Column</font></td><td> 
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="FCNAME">Description</Option>
		<Option value="JOBCODE">Control<br>Account</Option>
		</select>
 </td></tr>
         <tr><td colspan="2">
	<input type=submit value="Enter" >
    </td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="Machine" then%>
	
	<form action="../PrintData.asp?txtTable=Machine" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<table style="min-width:400px;"><tr><td>
    <font face="verdana" size="1">Search :</font></td><td> <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"></td></tr>
    <tr><td><font face="verdana" size="1">from Column</font></td><td> 
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="FCCODE">Code</Option>
		<Option value="DESCRIPTION">Description</Option>
		<Option value="FIXEDASSETCODE">Fixex<br>Asset</Option>
		<Option value="OWNERSHIP">Ownership</Option>
		<Option value="STATION_STATIONCODE">Station<br>Code</Option>
		<Option value="CONTROLJOB">Control<br>Account</Option>
		</select>		
</td></tr>
         <tr><td colspan="2">
	<input type=submit value="Enter" >
    </td></tr></table>
	</form>
	<%Elseif Request("txtKind") ="BusinessUnit" then%>
	
	<form action="../PrintData.asp?txtTable=BusinessUnit" method="post" target="_blank" style= 
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
	<input type=submit value="Enter" >
	</form>
	<%Elseif Request("txtKind") ="UserProfile" then%>
	
	<form action="../PrintData.asp?txtTable=UserProfile" method="post" target="_blank" style= 
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
	<input type=submit value="Enter" >
	</form>
	<%Elseif Request("txtKind") ="AccessRights" then%>
	
	<form action="../PrintData.asp?txtTable=AccessRights" method="post" target="_blank" style= 
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
	<input type=submit value="Enter" >
	</form>
	<%Elseif Request("txtKind") ="Trigger_Control" then%>
	
	<form action="../PrintData.asp?txtTable=Trigger_Control" method="post" target="_blank" style= 
	"screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<font face="verdana" size="1">Search : <input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" value="<%=txtSearch%>" width="5"><font face="verdana" size="1">&nbsp;from Column
		<select name="txtField" title="Case is not Sensitive">
		<Option value=""></Option>
		<Option value="ACTIVITY">Activity</Option>
		<Option value="SOURCE">Source Type</Option>
		<Option value="TARGET">Target Type</Option>
		</select>		
				&nbsp;&nbsp;
	<input type=submit value="Enter" >
	</form>
	<%Elseif Request("txtKind") ="Job" then%>
	
	<form action="../PrintData.asp?txtTable=Job" method="post" target="_blank" style= 
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
	<input type=submit value="Enter" >
	</form>		
	<%Elseif Request("txtKind") ="JobGroup" then%>
	
	<form action="../PrintData.asp?txtTable=JobGroup" method="post" target="_blank" style= 
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
	<input type=submit value="Enter" >
	</form>
	<%Else%>
		<input type="text" title="Case is not Sensitive" name="txtField" value="<%=txtField%>" width="5">	
	<%End if%>
    </div>
</div>   
</div>
</body>
</html>