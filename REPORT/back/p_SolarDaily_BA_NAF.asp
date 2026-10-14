<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim v_txtDate
	Dim rsAttDaily, strAttDaily, strGroup, strOrder, strSolarDaily, strMain
	Dim NO, cClass, nSpanTotal
	Dim txtGroupVH,v_txtGroupVH,txtVH,v_txtVH
	Dim sprevFcba, SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation,SprevGroupVH,SprevVH
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempGroupVH,tempVH
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	
	v_txtGroupPlantation   = Request("txtGroupPlantation")
	v_txtPlantation  	   = Request("txtPlantation")
	v_txtGroupRegion  	   = Request("txtGroupRegion")
	v_txtRegion  		   = Request("txtRegion")
	v_txtcompany  		   = Request("txtCompany")
	v_txtFcba 		       = trim(request("txtFcba"))
	v_txtDate			   = trim(request("txtDate"))
	v_txtGroupVH           = Request("txtGroupVH")
	v_txtVH				   = Request("txtVH")	
	if v_txtDate = "" then
	   v_txtDate = textdate(Now())
    end if	
	
	
	'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  
	nSpanTotal = 6
	
	    Dim selectFields,strFieldCons,selectCons,whereCons
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, BUSINESSUNIT_API.get_OrderData(fcba) norut, fcba, " 

						   
			wherecons = " where company is not null "
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' " 
			end if
			if v_txtGroupVH  <> "--ALL--" and v_txtGroupVH <> "" then
			   whereCons = wherecons & " And DESCRIPTION = '" & v_txtGroupVH & "' "
			end if
			
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   CompanyName=rsCat("FCCOMPANYNAME")			   
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 	
			end if
			
			if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",norut,fcba"
			    
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
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
	        if v_txtVH	<> "--SUMMARY--" and v_txtVH <> "" then
			   if v_txtVH	<> "--ALL--"  then
			      whereCons = wherecons & " And LOCATIONCODE = '" & v_txtVH & "' "
			   end if
			  selectCons = selectCons & " ,LOCATIONCODE,FCNAME "
			else
			    nSpanTotal = nSpanTotal - 2 
			end if 
	
	
	strMain = 	" SELECT " & strFieldCons & " " & _
				" vehiclegroup_api.get_OrderReport(vehicle_api.get_vehiclegroupcode(LOCATIONCODE,fcba),fcba) VHrut,LOCATIONTYPE," & _
				" decode(LOCATIONTYPE,'VH',DESCRIPTION,'LAIN-LAIN') DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME," & _
				" SUM (QUANTITY_0) QUANTITY_0," & _
				" SUM (QUANTITY_1) QUANTITY_1," & _
				" SUM (QUANTITY_2) QUANTITY_2" & _
				" FROM (SELECT FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME," & _
				" QUANTITY QUANTITY_0," & _
				" 0 QUANTITY_1," & _
				" 0 QUANTITY_2" & _
				" FROM VIEW_SOLARUSED_SUMMARY" & _
				" WHERE SIVDATE = TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
				" UNION" & _
				" SELECT FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME," & _
				" 0 QUANTITY_0," & _
				" SUM (QUANTITY) QUANTITY_1," & _
				" 0 QUANTITY_2" & _
				" FROM VIEW_SOLARUSED_SUMMARY" & _
				" WHERE SIVDATE >=" & _
				" TO_DATE (" & _
				" TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')," & _
				" 'MM')" & _
				" || '/01/'" & _
				" || TO_CHAR (" & _
				" TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')," & _
				" 'YYYY')," & _
				" 'MM/DD/YYYY')" & _
				" AND SIVDATE <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
				" GROUP BY FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME" & _
				" UNION" & _
				" SELECT FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME," & _
				" 0 QUANTITY_0," & _
				" 0 QUANTITY_1," & _
				" SUM (QUANTITY) QUANTITY_2" & _
				" FROM VIEW_SOLARUSED_SUMMARY" & _
				" WHERE SIVDATE <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
				" GROUP BY FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME)" & _         
				" GROUP BY FCBA," & _
				" LOCATIONTYPE," & _
				" DESCRIPTION," & _
				" LOCATIONCODE," & _
				" FCNAME"
	
	
	strSolarDaily = "select DESCRIPTION,GroupPlantation " & selectCons & " ,SUM (QUANTITY_0) QUANTITY_0,SUM (QUANTITY_1) QUANTITY_1,SUM (QUANTITY_2) QUANTITY_2 from ( " & strMain & " )" & whereCons & "  group by  VHrut,DESCRIPTION,GroupPlantation " & selectCons & "  Order by VHrut,DESCRIPTION,GroupPlantation " & selectCons  
	'response.write strSolarDaily
	Set rsAttDaily = Server.CreateObject("ADODB.Recordset")
	rsAttDaily.Open strSolarDaily, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Daily Solar Report</title>
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
<form name="frmEntry" method="post" action="p_SolarDaily_BA_NAF.asp">
	<div id="wrapper">
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	</div>
<!--#include file="FormSelectNewSum.html"-->
<table id="form-table" class="centered-margin">
	<tr>
		<th>TANGGAL</th>	
		<th>GROUP</th>
		<th>KODE PEMAKAI</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<td align="center">
			<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
				<input type="text" maxlength="10" name="txtDate" size="10" value="<%=v_txtDate%>" title="Type Date mm/dd/yyyy" >
			<% else %>
				<input type="text" maxlength="10" name="txtDate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
			<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtDate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center">
			<select name="txtGroupVH">	
				<option value="--ALL--">--ALL--</option>
			<%
				Dim rsVH
				Set rsVH = Server.CreateObject("ADODB.Recordset")
				rsVH.CursorLocation = adUseClient
				strsql =  " SELECT 'LAIN-LAIN' as Description from DUAL UNION select Description  from ( select distinct ORDERREPORT,Description from vehiclegroup where user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null order by ORDERREPORT) " 

				rsVH.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
				if not rsVH.eof then
					DO while not rsVH.eof
					if v_txtGroupVH = rsVH("Description") then
			%>			<option value="<%=rsVH("Description")%>" Selected><%=rsVH("Description")%></option>
					<%else%>
						<option value="<%=rsVH("Description")%>"><%=rsVH("Description")%></option>
					<%end if%>
			<%			rsVH.movenext
					loop
				End if
				rsVH.Close
				set rsVH = nothing
			%>
			</select>		
		</td>
		<td align="center">
			<select name="txtVH">	
				<option value="--SUMMARY--">--SUMMARY--</option>
			<%
				Dim rsVH1
				Set rsVH1 = Server.CreateObject("ADODB.Recordset")
				rsVH1.CursorLocation = adUseClient
				strsql =  " SELECT '--ALL--' as FCCODE, '' as FCNAME from DUAL UNION  SELECT DISTINCT FCCODE, FCNAME FROM Vehicle WHERE user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null ORDER BY FCNAME ASC"

				rsVH1.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
				if not rsVH1.eof then
					DO while not rsVH1.eof
					if v_txtVH = rsVH1("fccode") then
			%>			<option value="<%=rsVH1("fccode")%>" Selected><%=rsVH1("fccode")%> - <%=rsVH1("FCNAME")%></option>
					<%else%>
						<option value="<%=rsVH1("fccode")%>"><%=rsVH1("fccode")%> - <%=rsVH1("FCNAME")%></option>
					<%end if%>
			<%			rsVH1.movenext
					loop
				End if
				rsVH1.Close
				set rsVH1 = nothing
			%>
			</select>		
		</td>
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
</table>

<div id="header-wrapper">
		<div id="header">
		 <% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
		 <% End if%>
		 <% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
		 <% End if%>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN PEMAKAIAN SOLAR</h3>
		<h5>TANGGAL: <%=v_txtDate%></h5>
	</div>
</div>


<table border="0" width="100%" id="product-table">
   
	<tr>
		<th style="background-color:#00A300;color:#fff;">GROUP PEMAKAI</th>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th style="background-color:#00A300;color:#fff;">PLANTATION</th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;">GROUP REGION</th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;">REGION</th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;">COMPANY</th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;">SITE</th>
		<% end if %>
		<% if v_txtVH <>  "--SUMMARY--" and v_txtVH <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;">KODE PEMAKAI</th>
	           <th style="background-color:#00A300;color:#fff;">NAMA PEMAKAI</th>
	    <% end if %>
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th style="background-color:#00A300;color:#fff;">HARI INI <%=v_txtDate%></th>
		<% else %>
		<th style="background-color:#00A300;color:#fff;">HARI INI</th>
		<% end if%>
		
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th style="background-color:#00A300;color:#fff;">SAMPAI HARI INI <%=v_txtDate%></th>
		<% else %>
		<th style="background-color:#00A300;color:#fff;">SAMPAI HARI INI</th>
		<% end if%>
		<th style="background-color:#00A300;color:#fff;">SAMPAI DENGAN BULAN INI</th>
	</tr>
	

<%
Dim tempjob_forfieldstatus
Dim nTotal_0, nTotal_1, nTotal_2

	if not rsAttDaily.EOF then
		nTotal_0 = 0
		nTotal_1 = 0
		nTotal_2 = 0
		NO = 1
		
		sprevFcba=""
		
		while not rsAttDaily.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
			 
					<% tempGroupVH = rsAttDaily("description") 
					  
					%>
					<% If SprevGroupVH <>  rsAttDaily("description")  then %>		  
						  <td align="left">
							  <a href="p_SolarDaily_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtDate=<%=v_txtDate%>">
							  <%=rsAttDaily("description")%>
							  </a>
						  </td>
						  <% SprevGroupVH= rsAttDaily("description") %>	
					 <% else	%>
						  <td align="left"></td>					
					 <% end if  %>
					 
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     tempGroupVH = rsAttDaily("description") 
						   tempPlantation = rsAttDaily("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempVH ="--SUMMARY--"
					%>
					<% if SprevPlantation <> rsAttDaily("Plantation")   then %>
						<td align="center">
							<a href="p_SolarDaily_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtDate=<%=v_txtDate%>">
							<%=rsAttDaily("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsAttDaily("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  tempGroupVH = rsAttDaily("description") 
					    
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsAttDaily("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsAttDaily("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempVH ="--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsAttDaily("GroupRegion") then %>
						<td align="center">
							<a href="p_SolarDaily_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtDate=<%=v_txtDate%>">
							<%=rsAttDaily("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsAttDaily("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% tempGroupVH = rsAttDaily("description") 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsAttDaily("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsAttDaily("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsAttDaily("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempVH = "--SUMMARY--"
					   
					%>
					
					<% if SprevRegion <> rsAttDaily("Region") then %>
						<td align="center">
							<a href="p_SolarDaily_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtDate=<%=v_txtDate%>">
							<%=rsAttDaily("region")%>
							</a>
						</td>
						<% SprevRegion = rsAttDaily("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   tempGroupVH = rsAttDaily("description") 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsAttDaily("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsAttDaily("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsAttDaily("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsAttDaily("Region") 
					   else
					       tempRegion = ""
					   end if	
					   tempCompany =rsAttDaily("Company")
					   tempfcba = "--ALL--"	
					   tempVH ="--SUMMARY--"
					 %>
					
					<% if SprevCompany <> rsAttDaily("Company") then %>
						<td align="center">
							<a href="p_SolarDaily_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtDate=<%=v_txtDate%>">
							<%=rsAttDaily("company")%>
							</a>
						</td>
						<% SprevCompany = rsAttDaily("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
					 <% 
					   tempGroupVH = rsAttDaily("description") 
					   					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsAttDaily("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsAttDaily("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsAttDaily("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtcompany <>  "--SUMMARY--" And v_txtcompany <>  "" then
					      tempCompany = rsAttDaily("Company")
					   else
					      tempCompany = ""
                       end if				   
					   tempfcba = rsAttDaily("fcba") 
					   tempVH ="--ALL--"					   
					 %>
					<% if sprevFcba <> rsAttDaily("fcba") then %>
						<td align="center">
						    <a href="p_SolarDaily_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtDate=<%=v_txtDate%>">
							<%=rsAttDaily("fcba")%>
						</td>
						<% sprevFcba = rsAttDaily("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtVH <>  "--SUMMARY--" and v_txtVH <>  "" then %>
				    <% 
					   tempGroupVH = rsAttDaily("description") 
					   					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsAttDaily("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsAttDaily("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsAttDaily("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtcompany <>  "--SUMMARY--" And v_txtcompany <>  "" then
					      tempCompany = rsAttDaily("Company")
					   else
					      tempCompany = ""
                       end if	
                       if v_txtfcba <>  "--SUMMARY--"  And  v_txtfcba <>  "" Then 					   
					      tempfcba = rsAttDaily("fcba") 
					   else
					       tempfcba = ""
                       end if					   
                       if v_txtVH <>  "--SUMMARY--" and v_txtVH <>  "" then
					      tempVH = rsAttDaily("locationcode")
					   else
					      tempVH =""
                       end if					   
					 %>
					 <% If SprevVH <>  rsAttDaily("locationcode")  then %>		  
						  <td align="left">
							  <a href="p_SolarDaily_BA_NAF.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtDate=<%=v_txtDate%>">
							  <%=rsAttDaily("locationcode") %>
							  </a>
						  </td>
						  <% SprevVH= rsAttDaily("locationcode") %>	
						  <td align="left"><%=rsAttDaily("fcname")%></td>
					 <% else	%>
						  <td align="left"></td>	
						  <td align="left"><%=rsAttDaily("fcname")%></td>		
					 <% end if  %>	
				<% end if  %>	
				<td align="right"><%=FormatNumber(rsAttDaily("QUANTITY_0"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QUANTITY_1"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QUANTITY_2"),2)%></td>
			</tr>
<%
			NO = NO + 1
			nTotal_0 = nTotal_0 + Cdbl(rsAttDaily("QUANTITY_0"))
			nTotal_1 = nTotal_1 + Cdbl(rsAttDaily("QUANTITY_1"))
			nTotal_2 = nTotal_2 + Cdbl(rsAttDaily("QUANTITY_2"))

			
			rsAttDaily.MoveNext
		Wend
	end if
	rsAttDaily.Close
	set rsAttDaily = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="<%= 2 + nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(nTotal_0,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_1,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_2,2)%></td>				
			</tr>
</table>
</form>


