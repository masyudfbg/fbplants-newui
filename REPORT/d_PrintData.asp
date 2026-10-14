<%'Option Explicit
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>SAC Plant</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
</head>
<body>
<table class="entry">
<tr>
 
 <table width="700">
  <tr> 
    <td rowspan="5" width="100"><img src="../IMAGES/latar.jpg" width="210" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099"><%=request("txtModule")%>/<%=request("txtKind")%></FONT></td></tr>
 </table>
<tr>
 <td>
    <%' -----------------ATAS ISI%>
		    <!--#include file="d_browsetable.asp"-->
	</b><br>
	<form action="show.asp" method="post">
	<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
	    <tr>
		<td align="center" valign="top" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">No.</td>
	    <%	
		'response.Write(q)
		response.Flush()
		Set rsCat = connStr.Execute(q)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1
		if txtKind="LocationType" then
		   nCountofField=nCountofField+4
		Elseif txtKind="BusinessUnit" then
		   nCountofField=nCountofField+4
		Elseif txtKind="Trigger_Control" then
		   nCountofField=nCountofField+4
		end if%>
			<%IF txtKind="UserProfile" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">User<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Passwords</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Full Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Department</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Accessrights</td>
			<%ElseIF txtKind="Bank" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Bank<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Bank<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Control<br>Job</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Account<br>No</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Address</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">City</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">State</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Remarks</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Created</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Active</td>
			<%ElseIF txtKind="Stock" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Stock<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Stock<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Item<br>Type</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Manufacture<br>Part.No</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Manufacture<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Supplier<br>Part.No</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Supplier<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">UOM</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Stock<br>Group<br>Code</td>
			<%ElseIF txtKind="StockGroup" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Stock<br>Group<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Stock<br>Group<br>Description</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Control<br>Job</td>
			<%ElseIF txtKind="Supplier" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Supplier<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Supplier<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Control<br>Job</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Create</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Active</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Address</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">City</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">State</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Remark</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Bank<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Bank<br>No.</td>
			<%ElseIF txtKind="Store" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Store<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Store<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Remarks</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Active</td>
			<%ElseIF txtKind="Gang" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Gang<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Gang<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Remarks</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Active</td>
			<%ElseIF txtKind="Employee" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Employee<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Employee<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Section<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Gang<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Of<br>Birth</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Religion</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Race</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Identity<br>Card</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">High<br>Education</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Year<br>Of<br>Graduation</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Of<br>Join</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Of<br>Designation</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Present<br>Grade</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Present<br>Grade<br>Detail</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Present<br>Rate</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Previous<br>Grade</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Previous<br>Grade<br>Detail</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Previous<br>Rate</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Marital<br>Status</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Spouse</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Dependent<br>Child</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Jamsostek<br>No</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Jamsostek<br>Category</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Position</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Terminate</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">NPWP</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Nationality</td>
			<%ElseIF txtKind="Workshop" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Workshop<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Workshop<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Remarks</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Active</td>
			<%ElseIF txtKind="Vehicle" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Vehicle<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Vehicle<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Vehicle<br>Group<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Created</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Active</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Registration<br>No</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Make</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Model</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Engine<br>No.</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Chasis<br>No.</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Year<br>Of<br>Made</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Year<br>Of<br>Purchase</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Ownership</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Vehicle<br>Status</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Control<br>Alocation</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Control<br>Job</td>
			<%ElseIF txtKind="Contractor" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Contractor<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Contractor<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Control<br>Account</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Created</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Active</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Address</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">City</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">State</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Remarks</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Bank<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Bank<br>Account<br>No.</td>
			
			<%ElseIF txtKind="Building" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Building<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Building<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Building<br>Group</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Item</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Activation</td>
			   
			<%ElseIF txtKind="NURSERY" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Nursery<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Description<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Planted</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Crop</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Clone</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Source</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Type</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Status</td>
			   
			<%ElseIF txtKind="Allowded" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Allowded<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Business<br>Unit</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Allowded<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Group</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Allowance<br>Or<br>Deduction</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Crop</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Control<br>Job</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Manual<br>Auto</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Type<br>[ 1 ]</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Type<br>[ 2 ]</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Type<br>[ 3 ]</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Type<br>[ 4 ]</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Min<br>Value</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Max<br>Value</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Rate<br>[ 1 ]</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Rate<br>[ 2 ]</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Rate<br>[ 3 ]</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Rate<br>[ 4 ]</td>
			
			<%ElseIF txtKind="GCMASTER" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">ITEM</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">FCCODE</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">FCNAME</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">GCGROUP</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">ACTIVATION</td>
			 			  
			<%ElseIF txtKind="Field" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Field<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Field<br>Description</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Division</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Hectarage<br>Planted</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Terrain<br>Type</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Soil<br>Type</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Ownership</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Active<br>or<br>No</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Crop<br>Type</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Planting<br>Date</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Planting<br>Material</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Total<br>Stand</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Status</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Previous<br>Semester<br>ABW</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Present<br>Semester<br>ABW</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Harvesting<br>Based<br>ABW</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Main<br>Road</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Collection<br>Road</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Main<br>Drain</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Collection<br>Drain</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Subsidiary<br>Drain</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Collection<br>Point</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Harvesting<br>Bridge</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Permanent<br>Bridge</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Wood<br>Bridge</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Buist</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Owlnest</td>
			<%ElseIF txtKind="Station" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Station<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Station<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Control<br>Account</td>
			<%ElseIF txtKind="Machine" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Machine<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Machine<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Fix<br>Asset<br>No.</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Ownership</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Date<br>Purchase</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Purchase<br>Cost</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Station<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Active<br>Or<br>No</td>
			<%ElseIF txtKind="BusinessUnit" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Busines<br>Unit<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Busines<br>Unit<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Company<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Company<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Regional</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Ownership</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Manager<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Telephone</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Fax</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Email</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">KTU</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Estate<br>Or<br>Mill</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Crop<br>Type</td>
			<%ElseIF txtKind="AccessRights" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">User<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Module<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Sub Module<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Authorized</td>
			<%ElseIF txtKind="Trigger_Control" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Job<Br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Source<br>Type</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Target<br>Type</td>
			<%ElseIF txtKind="Job" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Activity<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Activity<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Job<br>Group<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Job<br>for<br>Crop<br>Type</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Job<br>for<br>Field<br>Status</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Job<br>Category</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">UOM</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Account<br>Code</td>
			<%ElseIF txtKind="JobGroup" then%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Job<br>Group<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Job<br>Group<br>Name</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Parent<br>Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Parent<br>Description</td>
            <%Else%>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Code</td>
			   <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">Description</td>
			   <%for i = 2 to nCountofField-4%>
				     <td align="center" valign="top" style="<%=cmdstyleTop%>"><font face="Verdana" size="1"><%=rsCat(i).Name%></td>
			   <%Next%>
			<%End if%>   
		</tr>
		<%While Not rsCat.EOF%>
			<tr>
			<td align="left" valign="top" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
			<%for i = 0 to nCountofField-4
			    if len(trim(rscat(i).value))>0 then%>
				   <td valign="top" style="<%=cmdstyle%>" align="left"><font face="Verdana" size="1"><%=rsCat(i)%></td>
				<%else%>
				   <td valign="top" style="<%=cmdstyle%>" align="left"><font face="Verdana" size="1">-</td>
				<%end if%>
			<%next%>
			</tr>
		  <%
		  rsCat.MoveNext
		  no=no+1
		Wend
		rsCat.close()
		rsCat=null%>
		<tr><td colspan="<%=nCountofField-4%>">&nbsp;</td></tr>
		<tr><td colspan="<%=nCountofField-4%>" align="left" valign="top"><font face="Verdana" size="1">Total Record : <%=NO - 1%> and Printed on <%=Formatdatetime(Now,1)%></td></tr>
		<input  type="hidden" value="<%=txtKind%>" name="txtKind">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		</form>
	</table>
  <%'------------------ BAWAH ISI%>
  </td>
</tr>
<!--#include file="../Bottom_Report.asp"-->