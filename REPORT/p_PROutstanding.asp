<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtDateFrom, v_txtDateTo, v_txtPRCode, v_txtItemCode

v_txtGroupPlantation  = Request("txtGroupPlantation")
v_txtPlantation  	  = Request("txtPlantation")
v_txtGroupRegion  	  = Request("txtGroupRegion")
v_txtRegion  		  = Request("txtRegion")
v_txtcompany  		  = Request("txtCompany")
v_txtfcba 			  = Request("txtfcba")

v_txtDateFrom = trim(Request("txtDateFrom"))
v_txtDateTo = trim(Request("txtDateTo"))
v_txtPRCode = trim(Request("txtPRCode"))
v_txtItemCode = trim(Request("txtItemCode"))
ScriptName = Request.ServerVariables("SCRIPT_NAME")

if "" & v_txtDateFrom = "" then v_txtDateFrom = textDate(date)
if "" & v_txtDateTo = "" then v_txtDateTo = textDate(date)
nSpanTotal = 0
add_newspan = 0

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" WHERE company is not null "

orderCons = " ORDER BY GroupPlantation"
if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
   whereCons = wherecons & " AND GroupPlantation = '" & v_txtGroupPlantation & "' "
	add_newspan = add_newspan +1
end if
if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
	if v_txtPlantation <> "--ALL--" then
	   whereCons = whereCons & " AND  Plantation = '" & v_txtPlantation & "' "
	end if   
	selectCons = selectCons & ",Plantation"
	orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
	add_newspan = add_newspan +1
else
	nSpanTotal = nSpanTotal - 1 
end if
if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
	if v_txtGroupRegion <> "--ALL--" then
	   whereCons = whereCons & " AND  GroupRegion = '" & v_txtGroupRegion & "' "
	end if
	selectCons = selectCons & " ,GroupRegion"
	orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
	add_newspan = add_newspan +1
else
	nSpanTotal = nSpanTotal - 1 
end if
if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
	if v_txtRegion <> "--ALL--" then
	   whereCons = whereCons & " AND  Region = '" & v_txtRegion & "' "
	end if
	selectCons = selectCons & " ,Region"
	orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
	add_newspan = add_newspan +1
else
	nSpanTotal = nSpanTotal - 1 
end if
if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
	if v_txtCompany  <> "--ALL--" then
	   whereCons = whereCons & " AND  company = '" & v_txtCompany & "' "
	end if
	selectCons = selectCons & " ,company"
	orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
	q = " SELECT fccode, fccompanycode, fcname, fccompanyname, fcmanager, fcktu, fccrop, fctype FROM BusinessUnit WHERE fccompanycode = '" & v_txtCompany & "'"	
	set rsCat = connStr.Execute(q)
	if not rsCat.eof then
	   CompanyName=rsCat("FCCOMPANYNAME")				   
	end if
	add_newspan = add_newspan +1
else
	nSpanTotal = nSpanTotal - 1 	
end if
if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
	if v_txtfcba  <> "--ALL--" then
	   whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
	end if
	selectCons = selectCons & ",fcba"
	orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
	set rsCat = connStr.Execute(q)
	if not rsCat.eof then
	   fcbaname = rsCat("fcName")
	   CompanyName=rsCat("FCCOMPANYNAME")
	   MgrName = rsCat("FCMANAGER")
	   KTUName = rsCat("FCKTU")
	   CropName=rsCat("FCCROP")
	   FcbaType=rsCat("fcType")
	end if
	add_newspan = add_newspan +1
else
	nSpanTotal = nSpanTotal - 1 
end if
%>
<html>

<head>
<title>ePlaS Reporting - PR Outstanding</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
		xAddEventListener(window, 'load',
		  function() {
			new xTableHeaderFixed('xthf-site', window);
		  }, false
		);
	</script>
	<link rel="stylesheet" href="/css/jquery-ui.css" />
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
    <script>
		$(function() {
		$( ".datepicker" ).datepicker();
		});
	</script>
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
		.printhide{display:none;}
		}
	</style>
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
		<form name="frmEntry" method="post" action="<%=ScriptName%>">
		<div class="printhide">
		<!--#include file="FormSelectNewSum.html"-->
		<table id="form-table" class="centered-margin">
			<tr>
				<form name="frmEntry" method="post" action="<%=ScriptName%>">
				<td><font face="verdana" size="1">Dari</font></td>
				<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
				</td>
				<td><font face="verdana" size="1">Sampai</font></td>
				<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
				</td>
				<td><font face="verdana" size="1">NO PR</font></td><td><input type="text" name="txtPRCode" value="<%=v_txtPRCode%>" /></td>
				<td><font face="verdana" size="1">KODE BARANG</font></td><td><input type="text" name="txtItemCode" value="<%=v_txtItemCode%>" /></td>
				<td><input type="submit" value="View" /></td>
				</form>
			</tr>
		</table>
		</div>
		</form>
	</div>
	<table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
			<tr>
				<th colspan="<%=add_newspan+10%>" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4>
				<%Else%>
					<h4>&nbsp;</h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					<h5><%=FcbaName%></h5>
				<%Else%>
					<h5>&nbsp;</h5>
				<% End if%>
					</div> 
	 
					<div id="subheader" class="centered">
						<h2>LAPORAN PR OUTSTANDING</h2>		
						<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
							<h5>TANGGAL: <%=v_txtdateto%></h5>
						<% else %>
							<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
						<% end if %>
					</div>
				</div>
			</th>
		</tr>
     <tr>

		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
		<th rowspan="2" style="color:#fff;">NO PR</th>		
		<th rowspan="2" style="color:#fff;">LINE ID</th>
		<th rowspan="2" style="color:#fff;">TANGGAL</th>
		<th rowspan="2" style="color:#fff;">KODE BARANG</th>
		<th rowspan="2" style="color:#fff;">NAMA BARANG</th>
		<th rowspan="2" style="color:#fff;">TANGGAL PENGIRIMAN</th>
		<th rowspan="2" style="color:#fff;">DESCRIPTION</th>
		<th rowspan="2" style="color:#fff;">CLASSIFICATION</th>
		<th rowspan="2" style="color:#fff;">STATUS</th>
		<th colspan="4" style="color:#fff;">QUANTITY</th>
	</tr>
	<tr>		
		<th style="color:#fff;">DI MINTA</th>
		<th style="color:#fff;">DI SETUJUI</th>
		<th style="color:#fff;">DI ORDER</th>	
		<th style="color:#fff;">SISA</th>
	</tr>
	</thead>
	<tbody>
	<%
	Dim rsPROut, strPROut , noline
	
	strPROut	=	" SELECT data.fcba," & _
					"        data.prcode," & _
					"        data.lineid," & _
					"        data.prdate," & _
					"        data.itemcode," & _
					"        stock.fcname," & _
					"        data.wanteddeliverydate," & _
					"        data.DESCRIPTION," & _
					"        data.fcname data_fcname," & _
					"        data.rowstateid," & _
					"        data.requestedqty," & _
					"        data.approvedqty," & _
					"        data.orderedqty," & _
					"        data.outstandingOrderedQty" & _
					"   FROM (SELECT a.fcba," & _
					"                a.prcode," & _
					"                a.lineid," & _
					"                b.prdate," & _
					"                a.itemcode," & _
					"                a.wanteddeliverydate," & _
					"                a.DESCRIPTION," & _
					"                a.requestedqty," & _
					"                c.fcname," & _
					"                a.rowstateid," & _
					"                a.approvedqty," & _
					"                a.orderedqty," & _
					"                (a.approvedqty - a.orderedqty) outstandingOrderedQty" & _
					"           FROM purchaserequestdetails a, purchaserequest b, DOCUMENT_CLASSIFICATION c " & _
					"          WHERE a.prcode = b.prcode AND c.FCCODE=b.document_classification and a.fcba = b.fcba and b.rowstateid='Approved'" & _
					" 			GROUP BY a.fcba, a.prcode, a.lineid, a.itemcode, b.prdate, a.wanteddeliverydate, a.DESCRIPTION, a.requestedqty, c.fcname, a.rowstateid, a.approvedqty, a.orderedqty) data," & _
					"        stock" & _
					"  WHERE     data.outstandingOrderedQty > 0" & _
					"        AND data.itemcode = stock.fccode" & _
					"        AND data.fcba = stock.fcba"
if  len(trim(v_txtDateFrom)) > 0 then
	strPROut = strPROut & " AND data.prdate >=	TO_DATE('" & v_txtDateFrom & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDateTo)) > 0 then
	strPROut = strPROut &  " AND data.prdate <= TO_DATE('" & v_txtDateTo & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtPRCode)) > 0 then
	strPROut = strPROut &  " AND data.prcode = '" & v_txtPRCode & "'"
end if	
if  len(trim(v_txtItemCode)) > 0 then
	strPROut = strPROut &  " AND data.itemcode = '" & v_txtItemCode & "'"
end if

	strMain = " SELECT " & strFieldCons & "," & _
	           "        prcode" & _
	           "        , lineid" & _
	           "        , prdate" & _
	           "        , itemcode" & _
	           "        , fcname" & _
	           "        , wanteddeliverydate" & _
			   "        , DESCRIPTION" & _
			   "        , data_fcname" & _
			   "        , rowstateid" & _
	           "        , requestedqty" & _
	           "        , approvedqty" & _
	           "        , orderedqty" & _
	           "        , outstandingOrderedQty" & _
	           " FROM (" & strPROut & " )"
			   
	strPROut = " SELECT GroupPlantation " & selectCons & "" & _
	           "        , prcode" & _
	           "        , lineid" & _
	           "        , prdate" & _
	           "        , itemcode" & _
	           "        , fcname" & _
	           "        , wanteddeliverydate" & _
			   "        , DESCRIPTION" & _
			   "        , data_fcname" & _
			   "        , rowstateid" & _
	           "        , requestedqty" & _
	           "        , approvedqty" & _
	           "        , orderedqty" & _
	           "        , outstandingOrderedQty" & _
	           " FROM (" & strMain & " )" & whereCons & "" & _
			   orderCons & ", prdate, prcode, lineid"
	
    'response.write strPROut	
	Set rsPROut = Server.CreateObject("ADODB.Recordset")
	rsPROut.Open strPROut, connstr, adOpenStatic
    
	if rsPROut.bof and rsPROut.eof then
	%>
	<tr>
			<td align="left" colspan="<%=add_newspan+10%>"><strong>DATA TIDAK DITEMUKAN</strong></td>
	</tr>	
	<%
	elseif Not rsPROut.eof then
		Dim nrequestedqty, napprovedqty, norderedqty, noutstandingOrderedQty
		
		nrequestedqty			=	0
		napprovedqty 			=	0
		norderedqty 			=	0
		noutstandingOrderedQty	=	0
		while Not rsPROut.eof  %>
		           
		  <tr>
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsPROut("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPROut("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPROut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsPROut("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPROut("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPROut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsPROut("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsPROut("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPROut("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPROut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsPROut("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsPROut("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsPROut("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPROut("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPROut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsPROut("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsPROut("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsPROut("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsPROut("Company")
					   tempfcba = "--ALL--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsPROut("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsPROut("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsPROut("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsPROut("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsPROut("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsPROut("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsPROut("fcba")
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
						    <%=rsPROut("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
			<td align="left"><%=rsPROut("prcode")%></td>
			<td align="right"><%=rsPROut("lineid")%></td>
			<td align="right"><%=rsPROut("prdate")%></td>
			<td align="left"><%=rsPROut("itemcode")%></td>
			<td align="left"><%=rsPROut("fcname")%></td>
			<td align="right"><%=rsPROut("wanteddeliverydate")%></td>
			<td align="left"><%=rsPROut("DESCRIPTION")%></td>
			<td align="left"><%=rsPROut("data_fcname")%></td>
			<td align="left"><%=rsPROut("rowstateid")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPROut("requestedqty")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPROut("approvedqty")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPROut("orderedqty")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPROut("outstandingOrderedQty")),0)%></td>
		  </tr>	
		      
<%
			nrequestedqty = nrequestedqty + cdbl(rsPROut("requestedqty"))
			napprovedqty = napprovedqty + cdbl(rsPROut("approvedqty"))
			norderedqty = norderedqty + cdbl(rsPROut("orderedqty"))
			noutstandingOrderedQty = noutstandingOrderedQty + cdbl(rsPROut("outstandingOrderedQty"))
			
			rsPROut.MoveNext
		wend
%>
		  <tr bgcolor="#F4F4A0">
			<td align="center" colspan="<%=add_newspan+9%>"><strong>TOTAL</td>
			<td align="right"><strong><%=FormatNumber(nrequestedqty,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(napprovedqty,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(norderedqty,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(noutstandingOrderedQty,0)%></strong></td>
		  </tr>	
<%
	end if
rsPROut.Close
set rsPROut = Nothing
%>
	</tbody>
</table>
<!--#include file="bottom.Asp"-->