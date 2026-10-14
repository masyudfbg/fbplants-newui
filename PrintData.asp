<%'Option Explicit
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
		<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
		<title>ePlaS Reporting - Data Master <%=request("txtTable")%></title>
		<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
        <link rel="stylesheet" href="/css/jquery-ui.css" />
		<script language="JavaScript" src='/js/x.js'></script>
        <script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
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

	</head>
	
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
          <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
		
             <tr>
                 <th colspan="50" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;" >
					<h4>&nbsp;&nbsp;PT. <%=Session("COMPANY")%></h4>
					<h5>&nbsp;&nbsp;<%=Session("BA")%></h5>
				</div>
		  <div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>Data Master <%=request("txtTable")%></h2>
				</div>
                <div style="clear:both;"></div>
				
			</div>
              </th>
			</tr>
			
			<!-----------------ATAS ISI----------------->
			<!--#include file="browsetable.asp"-->
			<% if txtTable="Field" then %>
			<tr>
				<th colspan="50" height="50" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:left;">
				<style>
				#head_filter .block_f .name_f {
					font-size:12px;
					position:absolute;
					margin-top:-20px;
					
				}
				#head_filter .block_f {
				height:30px; float:left;  margin-left:10px;
				}
				.mrgn-left10 { margin-left:10px;}
				
				</style>
					<form action="PrintData.asp?txtTable=Field" method="post">
						<div id="head_filter" style="position:fixed; top:85px; width:1500px;">
							<div class="block_f">
								<span class="name_f">Include Date</span>
								<%if  LEN(TRIM(txtValidForm)) > 0 then%>
									<input class="datepicker" type="text" name="txtValidForm" value="<%=txtValidForm%>" style="height:20px; width:80px;">
								<%else%>
									<input class="datepicker" type="text" name="txtValidForm" value="<%=textDate(date)%>" style="height:20px; width:80px;">
								<%end if%>
							</div>
							<div class="block_f mrgn-left10" >
						<% 	Dim rsDIVISION
							qfilter = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0001' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
							Set rsDIVISION = Server.CreateObject("ADODB.Recordset")
							rsDIVISION.CursorType = 3
							rsDIVISION.CursorLocation = 3
							rsDIVISION.Open qfilter, connstr
							if not rsDIVISION.eof then%>
								<span class="name_f">Division</span>
								<select style="min-width:100px" name="DIVISION" title="DIVISION Character a s/d z or 1 s/d 9">         
								<option value=""></option> 
							<% Do While Not rsDIVISION.eof%>
								<option value="<%=rsDIVISION("FCCODE")%>" <% If Request("DIVISION") = rsDIVISION("FCCODE") Then %> Selected <%End if%>><%=rsDIVISION("FCNAME")%></option> 
							<%rsDIVISION.movenext
							loop
							End if
							rsDIVISION.Close
							set rsDIVISION = Nothing
							qfilter=""
							%>
								</select>
							</div>
							<div class="block_f mrgn-left10" >
							<% 	Dim rsStatus
							qstat = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0006' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
							Set rsStatus = Server.CreateObject("ADODB.Recordset")
							rsStatus.CursorType = 3
							rsStatus.CursorLocation = 3
							rsStatus.Open qstat, connstr
							if not rsStatus.eof then%>
								<span class="name_f">Status</span>
								<select style="min-width:70px" name="STATUS" title="STATUS Character a s/d z or 1 s/d 9">         
								<option value=""></option> 
							<%			Do While Not rsStatus.eof%>
								<option value="<%=rsStatus("FCCODE")%>" <% If Request("STATUS") = rsStatus("FCCODE") Then %> Selected <%End if%>><%=rsStatus("FCNAME")%></option> 
							<%rsStatus.movenext
							loop
							End if
							rsStatus.Close
							set rsStatus = Nothing
							qstat=""
							%>
								</select>
							</div>
							<div class="block_f mrgn-left10" >
							
							<% 	Dim rsCropB
							qcrop = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
							Set rsCropB = Server.CreateObject("ADODB.Recordset")
							rsCropB.CursorType = 3
							rsCropB.CursorLocation = 3
							rsCropB.Open qcrop, connstr
							if not rsCropB.eof then%>
							<span class="name_f">Crop</span>
							<select style="min-width:150px" name="CROPTYPE" title="CROPTYPE Character a s/d z or 1 s/d 9">         
							<option value=""></option> 
							<%			Do While Not rsCropB.eof%>
								<option value="<%=rsCropB("FCCODE")%>" <% If Request("CROPTYPE") = rsCropB("FCCODE") Then %> Selected <%End if%>><%=rsCropB("FCNAME")%></option> 
							<%rsCropB.movenext
							loop
							End if
							rsCropB.Close
							set rsCropB = Nothing
							qcrop=""
							%>
							</select>
							</div>
							<div class="block_f mrgn-left10" >
							<% 	Dim rsSoil
							qsoil = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0003' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
							Set rsSoil = Server.CreateObject("ADODB.Recordset")
							rsSoil.CursorType = 3
							rsSoil.CursorLocation = 3
							rsSoil.Open qsoil, connstr
							if not rsSoil.eof then%>
							<span class="name_f">Soil</span>
							<select style="min-width:150px" name="SOILTYPE" title="SOILTYPE Character a s/d z or 1 s/d 9">         
							<option value=""></option> 
							<%			Do While Not rsSoil.eof%>
								<option value="<%=rsSoil("FCCODE")%>" <% If Request("SOILTYPE") = rsSoil("FCCODE") Then %> Selected <%End if%>><%=rsSoil("FCNAME")%></option> 
							<%rsSoil.movenext
							loop
							End if
							rsSoil.Close
							set rsSoil = Nothing
							qsoil=""
							%>
							</select>
							</div>
							<div class="block_f mrgn-left10" >
							<% 	Dim rsTerrain
							qterrain = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0002' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
							Set rsTerrain = Server.CreateObject("ADODB.Recordset")
							rsTerrain.CursorType = 3
							rsTerrain.CursorLocation = 3
							rsTerrain.Open qterrain, connstr
							if not rsTerrain.eof then%>
							<span class="name_f">Terrain</span>
							<select style="min-width:210px" name="TERRAINTYPE" title="TERRAINTYPE Character a s/d z or 1 s/d 9">         
							<option value=""></option> 
							<%			Do While Not rsTerrain.eof%>
								<option value="<%=rsTerrain("FCCODE")%>" <% If Request("TERRAINTYPE") = rsTerrain("FCCODE") Then %> Selected <%End if%>><%=rsTerrain("FCNAME")%></option> 
							<%rsTerrain.movenext
							loop
							End if
							rsTerrain.Close
							set rsTerrain = Nothing
							qterrain=""
							%>
							</select>
							</div>
							<div class="block_f mrgn-left10">
							<% 	Dim rsOwnership
							qown = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0004' and fcba='" & Session("fcBA") & "' order by P.FCNAME desc"
							Set rsOwnership = Server.CreateObject("ADODB.Recordset")
							rsOwnership.CursorType = 3
							rsOwnership.CursorLocation = 3
							rsOwnership.Open qown, connstr
							if not rsOwnership.eof then%>
							<span class="name_f">Ownership</span>
							<select style="min-width:100px" name="OWNERSHIP" title="OWNERSHIP Character a s/d z or 1 s/d 9">         
							<option value=""></option> 
							<%			Do While Not rsOwnership.eof%>
							
								<option value="<%=rsOwnership("FCCODE")%>" <% If Request("OWNERSHIP") = rsOwnership("FCCODE") Then %> Selected <%End if%>><%=rsOwnership("FCNAME")%></option> 
								
							<%rsOwnership.movenext
							loop
							End if
							rsOwnership.Close
							set rsOwnership = Nothing
							qown=""
							%>
							</select>
							</div>
							<div class="block_f mrgn-left10">
							<span class="name_f">Harvesting Section</span>
							<select style="min-width:150px" name="HARVESTINGSECTION">
							<option value = "" 
							<%If Len(Trim(txtHARVESTINGSECTION)) = 0 then%> Selected <%End If%>>OTHERS</option>
							<%
							strHarvestingSection =	" SELECT '' FCCODE, 'ALL' FCNAME FROM DUAL " & _
													" UNION " & _
													" SELECT FCCODE, FCNAME FROM HARVESTINGGROUP" & _
													" WHERE FCBA = '" & Session("FCBA") & "'" & _
													" ORDER BY FCCODE"
							
							Set rsHarvestingSection = Server.CreateObject("ADODB.Recordset")
							rsHarvestingSection.Open strHarvestingSection, connstr, 3, 3 
							
							if not rsHarvestingSection.eof then
							Do While Not rsHarvestingSection.eof
							%>

							<option value="<%=rsHarvestingSection("FCCODE")%>" <% If txtHARVESTINGSECTION = rsHarvestingSection("FCCODE") Then %> Selected <%End if%>><%=rsHarvestingSection("FCNAME")%></option> 
							<%
								rsHarvestingSection.movenext
							loop
							end if
							strHarvestingSection = ""
							rsHarvestingSection.Close
							set rsHarvestingSection = Nothing
							%>
							</select>
							</div>
							<div class="block_f mrgn-left10">
								<span class="name_f">Activation</span>
								<select style="min-width:100px" name="ACTIVATION" title="ACTIVATION Character a s/d z or 1 s/d 9">
								<%If Len(Trim(Request("ACTIVATION"))) = 0 then%> 
									<Option value="" selected></Option>
									<Option value="Y" >Yes</Option>
									<Option value="N">No</Option>
								<%elseIf Request("ACTIVATION") = "Y" then%> 
									<Option value=""></Option>
									<Option value="Y" selected>Yes</Option>
									<Option value="N">No</Option>
								<%elseIf Request("ACTIVATION") = "N" then%>
									<Option value=""></Option>
									<Option value="Y">Yes</Option>
									<Option value="N" selected>No</Option>
									<%end if%>
								</select>
							</div>
							<input class="mrgn-left10" type="submit" value="View">
						</div>
				</th>
				
			</tr>
			<%else%>
			<form action="show.asp" method="post">
			<% end if%>
			    	 
			<tr>
					
<%	
	' response.Write(q)
	response.Flush()
	Set rsCat = connStr.Execute(q)
	if Not rsCat.eof then
		nCountofField	=	rsCat.fields.count - 1
	Else
		nCountofField	=	rsCat.fields.count - 1
	end if
	
	NO=1
	if txtTable = "LocationType" then
	   nCountofField	=	nCountofField 
	Elseif txtTable = "BusinessUnit" then
	   nCountofField	=	nCountofField 
	Elseif txtTable = "Trigger_Control" then
	   nCountofField	=	nCountofField 
	Elseif txtTable = "MaritalStatus" then
	   nCountofField	=	nCountofField 
	'Special for Administrator, Added by NAF, 2-May-2013   wiji surat rekomendasi ke BPN, atas nama Ibu Anne
	Elseif txtModule = "Administrator" then
	   nCountofField	=	nCountofField + 6
	end if
%>
<%IF txtTable="UserProfile" then%>
   <th>User Name</th>
   <th>Passwords</th>
   <th>Full Name</th>
   <th>Department</th>
   <th>Accessrights</th>
<%ElseIF txtTable="Nursery" then%>
   <th>Nursery Code</th>
   <th>Description Name</th>
   <th>Date Planted</th>
   <th>Crop</th>
   <th>Clone</th>
   <th>Source</th>
   <th>Type</th>
   <th>Status</th>
<%ElseIF txtTable="UnitOfMeasure" then%>
   <th>Unit Of Measurement Code</th>
   <th>Unit Of Measurement Desc</th>
<%ElseIF txtTable="Bank" then%>
   <th>Bank Code</th>
   <th>Bank Name</th>
   <th>Control Job</th>
   <th>Account No</th>
   <th>Address</th>
   <th>City</th>
   <th>State</th>
   <th>Remarks</th>
   <th>Date Created</th>
   <th>Active</th>
   <th>Trigger System</th>
   <th>Bank Type</th>
<%ElseIF txtTable="BankStock" then%>
   <th>Stock Code</th>
   <th>Stock Name</th>
   <th>Item Type</th>
   <th>Manufacture Part.No</th>
   <th>Manufacture Code</th>
   <th>Supplier Part.No</th>
   <th>Supplier Code</th>
   <th>UOM</th>
   <th>Stock Group Code</th>
   <th>Stock Sub.Group Code</th>
   
   <th>Is Less Zero</th>
<%ElseIF txtTable="Stock" then%>
   <th>Stock Code</th>
   <th>Stock Name</th>
   <th>Item Type</th>
   <th>Manufacture Part.No</th>
   <th>Manufacture Code</th>
   <th>Supplier Part.No</th>
   <th>Supplier Code</th>
   <th>UOM</th>
   <th>Stock Group Code</th>
   <th>Is Less Zero</th>
<%ElseIF txtTable="StockGroup" then%>
   <th>Stock Group Code</th>
   <th>Stock Group Description</th>
   <th>Control Job</th>
<%ElseIF txtTable="Supplier" then%>
   <th>Supplier Code</th>
   <th>Supplier Name</th>
   <th>Control Job</th>
   <th>Date Create</th>
   <th>Active</th>
   <th>Address</th>
   <th>City</th>
   <th>State</th>
   <th>Remark</th>
   <th>Bank Name</th>
   <th>Bank No.</th>
   <th>Bank Ownership</th>
   <th>Phone</th>
   <th>Fax</th>
   <th>TYPE</th>
<%ElseIF txtTable="Store" then%>
   <th>Store Code</th>
   <th>Store Name</th>
   <th>Remarks</th>
   <th>Active</th>
<%ElseIF txtTable="Gang" then%>
   <th>Gang Code</th>
   <th>Gang Name</th>
   <th>Remarks</th>
   <th>Active</th>
   <th>Gang Type</th>
<%ElseIF txtTable="Transporter" then%>
   <th>Code</th>
   <th>Description</th>
   <th>Control Job</th>
   <th>Date Created</th>
   <th>Activation</th>
   <th>Address</th>
   <th>City</th>
   <th>State</th>
   <th>Remarks</th>
   <th>Bank Name</th>
   <th>Bank Account</th>
<%ElseIF txtTable="MultiLevel_Harvesting_Rate" then%>
   <th>Level No</th>
   <th>Crop Code</th>
   <th>Division</th>
   <th>Planting Year</th>
   <th>Terrain Type</th>
   <th>Min Bunch</th>
   <th>Max Bunch</th>
   <th>Rate</th>
   <th>Rate Holiday</th>
 <%ElseIF txtTable="Employee" then%>
	<th>Employee Code</font></th>
	<th>Employee Name</font></th>
	<th>Section Name</font></th>
	<th>Gang Code</font></th>
	<th>Date Of Birth</font></th>
	<th>Religion</font></th>
	<th>Race</font></th>
	<th>Identity Card</font></th>
	<th>High Education</font></th>
	<th>Year of Graduation</font></th>
	<th>Date of Join</font></th>
	<th>Date of Designation</font></th>
	<th>Present Grade</font></th>
	<th>Present Grade Detail</font></th>
	<th>Present Rate</font></th>
	<th>Previous Grade</font></th>
	<th>Previous Grade Detail</font></th>
	<th>Previous Rate</font></th>
	<th>Maritial Status</font></th>
	<th>Spouse</font></th>
	<th>Dependent Child</font></th>
	<th>Jamsostek No</font></th>
	<th>Jamsostek Category</font></th>
	<th>Position</font></th>
	<th>Date Terminate</font></th>
	<th>Nationality</font></th>
	<th>NPWP</font></th>
	<th>Nature Type</font></th>
	<th>Included Tax Calculation</font></th>
	<th>Tax Status</font></th>
	<th>Payment Type</font></th>
	<th>Bank Name</font></th>
	<th>Account No</font></th>
	<th>Ownership of Account</font></th>
	<th>Gender</font></th>
	<th>Tax Paid by Company</font></th>
	<th>Address</font></th>
	<th>BPJS Kesehatan Category</font></th>
	<th>BPJS Kesehatan No</font></th>
	<th>Golongan</font></th>
	<th>NIK Perusahaan</font></th>
<%ElseIf Request("txttable") ="RangeTaxGroup" then%>
	<tr>	
	<th>FC CODE</th>	
	<th>FC NAME</th>	
	<th>RANGE FROM</th>	
	<th>RANGE UNTIL</th>	
	<th>TARIFF PERCENTAGE</th>	
	<th>REMARKS</th>
	</tr>
<%ElseIf Request("txttable") ="MaritalStatus" then%>
	<tr>	
	<th>CODE</th>	
	<th>NAME</th>	
	</tr>
<%ElseIf Request("txttable") ="UntaxableIncome" then%>
	<tr>
	<th>MARITAL STATUS CODE</th>	
	<th>FCNAME</th>	
	<th>UTI EMPLOYEE</th>	
	<th>UTI SPOUSE</th>	
	<th>UTI CHILD</th>	
	<th>REMARKS</th>
	</tr>
<%ElseIf Request("txttable") ="EmployeeAllowdedDetail" then%>
	<tr>
	<th>EMPLOYEE CODE</th>
	<th>EMPLOYEE NAME</th>
	<th>ALLOWDED CODE</th>
	<th>ALLOWDED NAME</th>	
	<th>RATE 1</th>	
	<th>RATE 2</th>	
	<th>RATE 3</th>	
	<th>RATE 4</th>
	<th>RATE 5</th>
	<th>REMARKS</th>
	</tr>
<%ElseIF txtTable="Workshop" then%>
   <th>Workshop Code</th>
   <th>Workshop Name</th>
   <th>Remarks</th>
   <th>Active</th>
<%ElseIF txtTable="VehicleGroup" then%>
   <th>Vehicle Group Code</th>
   <th>Vehicle Group Name</th>
   <th>Units</th>
   <th>Workshop Factor</th>
   <th>Cost Per Vehicle</th>
   <th>Noof Vehicle</th>
   <th>Vehicle Type Code</th>
<%ElseIF txtTable="Vehicle" then%>
   <th>Vehicle Code</th>
   <th>Vehicle Name</th>
   <th>Vehicle Group Code</th>
   <th>Vehicle Sub Group Code</th>
   <th>Date Created</th>
   <th>Active</th>
   <th>Registration No</th>
   <th>Make</th>
   <th>Model</th>
   <th>Engine No.</th>
   <th>Chasis No.</th>
   <th>Year Of Made</th>
   <th>Year Of Purchase</th>
   <th>Ownership</th>
   <th>Vehicle Status</th>
   <th>Control Alocation</th>
   <th>Control Job</th>
   <th>Date Terminate</th>
<%ElseIF txtTable="Contractor" then%>
   <th>Contractor Code</th>
   <th>Contractor Name</th>
   <th>Control Account</th>
   <th>Date Created</th>
   <th>Active</th>
   <th>Address</th>
   <th>City</th>
   <th>State</th>
   <th>Remarks</th>
   <th>Bank Name</th>
   <th>Bank Account No.</th>
<%ElseIF txtTable="Building" then%>
   <th>Building Code</th>
   <th>Building Name</th>
   <th>Building Group</th>
   <th>Item</th>
   <th>Activation</th>
<%ElseIF txtTable="Allowded" then%>
   <th>Allowded Code</th>
   <th>Business Unit</th>
   <th>Allowded Name</th>
   <th>Group</th>
   <th>Allowance Or Deduction</th>
   <th>Crop</th>
   <th>Control Job</th>
   <th>Manual Auto</th>
   <th>Type [ 1 ]</th>
   <th>Type [ 2 ]</th>
   <th>Type [ 3 ]</th>
   <th>Type [ 4 ]</th>
   <th>Min Value</th>
   <th>Max Value</th>
   <th>Rate [ 1 ]</th>
   <th>Rate [ 2 ]</th>
   <th>Rate [ 3 ]</th>
   <th>Rate [ 4 ]</th>
   <th>Remarks</th>
<%ElseIF txtTable="Field" then%>
   <th>Field Code</th>
   <th>Field Description</th>
   <th>Division</th>
   <th>Hectarage Planted</th>
   <th>Terrain Type</th>
   <th>Soil Type</th>
   <th>Ownership</th>
   <th>Koperasi</th>
   <th>Active or No</th>
   <th>Crop Type</th>
   <th>Planting Date</th>
   <th>Planting Material</th>
   <th>Total Stand</th>
   <th>Status</th>
   <th>Previous Semester ABW</th>
   <th>Present Semester ABW</th>
   <th>Harvesting Based ABW</th>
   <th>Main Road</th>
   <th>Collection Road</th>
   <th>Main Drain</th>
   <th>Collection Drain</th>
   <th>Subsidiary Drain</th>
   <th>Collection Point</th>
   <th>Harvesting Bridge</th>
   <th>Permanent Bridge</th>
   <th>Wood Bridge</th>
   <th>Buist</th>
   <th>Owlnest</th>
   <th>Isactive Harvested</th>
   <th>Valid Form</th>
<%ElseIF txtTable="Station" then%>
   <th>Station Code</th>
   <th>Station Name</th>
   <th>Control Account</th>
<%ElseIF txtTable="OpHarvRate" then%>
   <th>Crop Code</th>
   <th>ABW Census</th>
   <th>Terrain Type</th>
   <th>ABW Min</th>
   <th>ABW Max</th>
   <th>Basis Quantity</th>
   <th>Basis Rate</th>
   <th>Extra Basis Rate</th>
   <th>Basis Ha</th>
<%ElseIF txtTable="Machine" then%>
   <th>Machine Code</th>
   <th>Machine Name</th>
   <th>Fix Asset No.</th>
   <th>Ownership</th>
   <th>Date Purchase</th>
   <th>Purchase Cost</th>
   <th>Station Code</th>
   <th>Active Or No</th>
<%ElseIF txtTable="BusinessUnit" then%>
   <th>Busines Unit Code</th>
   <th>Busines Unit Name</th>
   <th>Company Code</th>
   <th>Company Name</th>
   <th>Regional</th>
   <th>Ownership</th>
   <th>Manager Name</th>
   <th>Telephone</th>
<%ElseIF txtTable="AccessRights" then%>
   <th>User Name</th>
   <th>Module Code</th>
   <th>Sub Module Code</th>
   <th>Authorized</th>
<%ElseIF txtTable="Trigger_Control" then%>
   <th>Job Code</th>
   <th>Source Type</th>
   <th>Target Type</th>
<%ElseIF txtTable="Job" then%>
   <th>Activity Code</th>
   <th>Activity Name</th>
   <th>Job.Group Code</th>
   <th>Job.Crop Type</th>
   <th>Job.Field Status</th>
   <th>Job Category</th>
   <th>Job.Balancesheet</th>
   <th>Job.ProfitLoss</th>
   <th>Job.CashFlow</th>
   <th>Unit-UOM</th>  
   <th>Output-UOM</th>
   <th>Account Code</th>
<%ElseIF txtTable="Jobmaster" then%>
   <th>Activity Code</th>
   <th>Activity Name</th>
   <th>Job.Group Code</th>
   <th>Job.Crop Type</th>
   <th>Job.Field Status</th>
   <th>Job Category</th>
   <th>Job.Balancesheet</th>
   <th>Job.ProfitLoss</th>
   <th>Job.CashFlow</th>
   <th>Job.Budget</th>   
   <th>Unit-UOM</th>  
   <th>Output-UOM</th>
   <th>Account Code</th>
<%ElseIF txtTable="JobGroup" then%>
   <th>Job Group Code</th>
   <th>Job Group Name</th>
   <th>Parent Code</th>
   <th>Parent Description</th>
<% ' Special for Administrator, Added by NAF, 2 May 2013 
%>   
<%ElseIF txtTable="Roles" then%>
   <th>Roles</th>
   <th>Roles Name</th>
<%ElseIF txtTable="RolesMenu" then%>
   <th>Roles</th>
   <th>Menu Code</th>
   <th>Menu Name</th>
<%ElseIF txtTable="Users" then%>
   <th>User Code</th>
   <th>Password</th>
   <th>Full Name</th>
   <th>Department</th>
   <th>Accessrights</th>
   <th>Active</th>
   <th>Valid From</th>
   <th>Valid Until</th>
<%ElseIF txtTable="UsersRoles" then%>
   <th>User Code</th>
   <th>Roles</th>
   <th>Active</th>
   <th>Valid From</th>
   <th>Valid Until</th>
<%ElseIF txtTable="UsersCompany" then%>
   <th>User Code</th>
   <th>Company</th>
   <th>Default</th>
<%ElseIF txtTable="UsersBusinessUnit" then%>
   <th>User Code</th>
   <th>Business Unit</th>
   <th>Default</th>
<%ElseIF txtTable="UsersChanged" then%>
   <th>User Code</th>
   <th>Default Business Unit</th>
   <th>Default Company</th>
<%ElseIF txtTable="Koperasi" then%>
	<th>Kode</th>
	<th>Deskripsi</th>
	<th>Tanggal Berdiri</th>
	<th>No Badan Hukum</th>
	<th>Tanggal Badan Hukum</th>
	<th>Alamat</th>
	<th>Kota</th>
	<th>Negara</th>
	<th>Telp</th>
	<th>Fax</th>
	<th>Bentuk Koperasi</th>
	<th>Jenis Koperasi</th>
	<th>Kelompok Koperasi</th>
	<th>Sektor Usaha</th>
	<th>NIK</th>
	<th>Sertifikasi</th>
	<th>Status</th>
					<% 
						' End of Administrator, Added by NAF, 2 May 2013
					%>
					<%Else%>
					   <th>Code</th>
					   <th>Description</th>
					   <%for i = 2 to nCountofField%>
							 <th><%=rsCat(i).Name%></th>
					   <%Next%>
					<%End if%>   
				</tr>
                </thead>
				<%
				'---- Nama kolom query untuk melengkapi judul kolom yang kurang (js/fbp-table-auto.js)
				Dim fbtNames : fbtNames = ""
				for i = 0 to nCountofField
					fbtNames = fbtNames & """" & Replace(rsCat(i).Name & "", """", "") & ""","
				next
				%>
				<script type="text/javascript">window.FBT_FIELD_NAMES = [<%=fbtNames%>];</script>
                <tbody>
				<%If (rsCat.eof and rsCat.bof) Then%>
					<tr>
						<td Colspan="<%=nCountofField-2%>">Data Is Empty</td>
					</tr>
				<%End If%>
				<%While Not rsCat.EOF%>
					<tr>
					<%for i = 0 to nCountofField
						if len(trim(rscat(i).value)) > 0 then%>
						   <td><%=rsCat(i)%></td>
						<%else%>
						   <td>-</td>
						<%end if%>
					<%next%>
					</tr>
				  <%
				  rsCat.MoveNext
				  no = no + 1
				Wend
				rsCat.close()
				rsCat = null%>
				<input  type="hidden" value="<%=txtTable%>" name="txtTable">
				<input  type="hidden" value="<%=txtModule%>" name="txtModule">
				</form>
                </tbody>
			</table>
			<!------------------ BAWAH ISI ------------------>
			<!--#include file="Bottom_Report.asp"-->
		</div> 
	</body>
</html>