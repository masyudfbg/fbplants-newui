<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth,v_txtDate
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
	v_txtDate			   = trim(request("txtDate"))
	v_txtGroupVH           = Request("txtGroupVH")
	v_txtVH				   = Request("txtVH")	
	if v_txtDate = "" then
	   v_txtDate = textdate(Now())
    end if	
	
	        v_txtyear 		= Request("txtyear")
			v_txtmonth 		= lpad(Request("txtmonth"),"0",2)
	'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  
	nSpanTotal = 6
	
	    Dim selectFields,strFieldCons,selectCons,whereCons,orderCons 
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, BUSINESSUNIT_API.get_OrderData(fcba) norut, fcba, " 

						   
			wherecons = " where company is not null "
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' " 
			end if
			if v_txtGroupVH  <> "--ALL--" and v_txtGroupVH <> "" then
			   whereCons = wherecons & " And DESCRIPTION = '" & v_txtGroupVH & "' "
			end if
			
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' " 
			end if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
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
			    orderCons  = orderCons & ",norut "
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
			  orderCons  = orderCons & ",LOCATIONCODE,FCNAME "
			else
			    nSpanTotal = nSpanTotal - 2 
			end if 
	
	
	strMain = 	" SELECT " & strFieldCons & " " & _
				"    vehiclegroup_api.get_OrderReport(vehicle_api.get_vehiclegroupcode(LOCATIONCODE,fcba),fcba) VHrut,LOCATIONTYPE, " & _
				"	 decode(LOCATIONTYPE,'VH',DESCRIPTION,'LAIN-LAIN') DESCRIPTION,LOCATIONCODE,FCNAME,QTY_01,QTY_02,QTY_03,QTY_04, " & _
				"	 QTY_05,QTY_06,QTY_07,QTY_08,QTY_09,QTY_10,QTY_11,QTY_12,QTY_SBI " & _           
				"	 from ( SELECT FCBA,LOCATIONTYPE,DESCRIPTION,LOCATIONCODE,FCNAME, " & _
				"				 Decode(to_char(sivdate,'MM'),1,sum(nvl(QUANTITY,0)),0)  QTY_01, " & _
				"				 Decode(to_char(sivdate,'MM'),2,sum(nvl(QUANTITY,0)),0)  QTY_02, " & _
				"				 Decode(to_char(sivdate,'MM'),3,sum(nvl(QUANTITY,0)),0)  QTY_03, " & _
				"				 Decode(to_char(sivdate,'MM'),4,sum(nvl(QUANTITY,0)),0)  QTY_04, " & _
				"				 Decode(to_char(sivdate,'MM'),5,sum(nvl(QUANTITY,0)),0)  QTY_05, " & _
				"				 Decode(to_char(sivdate,'MM'),6,sum(nvl(QUANTITY,0)),0)  QTY_06, " & _
				"				 Decode(to_char(sivdate,'MM'),7,sum(nvl(QUANTITY,0)),0)  QTY_07, " & _
				"				 Decode(to_char(sivdate,'MM'),8,sum(nvl(QUANTITY,0)),0)  QTY_08, " & _
				"				 Decode(to_char(sivdate,'MM'),9,sum(nvl(QUANTITY,0)),0)  QTY_09, " & _
				"				 Decode(to_char(sivdate,'MM'),10,sum(nvl(QUANTITY,0)),0)  QTY_10, " & _
				"				 Decode(to_char(sivdate,'MM'),11,sum(nvl(QUANTITY,0)),0)  QTY_11, " & _
				"				 Decode(to_char(sivdate,'MM'),12,sum(nvl(QUANTITY,0)),0)  QTY_12, " & _
				"				 sum(nvl(QUANTITY,0)) QTY_SBI  " & _ 
				"			  FROM VIEW_SOLARUSED_SUMMARY " & _
				"			  where to_char(sivdate,'YYYYMM') <= '" & v_txtyear & v_txtmonth &"'" & _
				"			  and  to_char(sivdate,'YYYY')= '" & v_txtyear &"'" & _
				"			  group by FCBA,LOCATIONTYPE,DESCRIPTION, " & _
				"					   LOCATIONCODE,FCNAME,to_char(sivdate,'MM') " & _
				"			 ) "
	
	
	strSolarDaily = " select DESCRIPTION,GroupPlantation " & selectCons & " ,SUM(QTY_01) QTY_01,SUM(QTY_02) QTY_02,SUM(QTY_03) QTY_03,SUM(QTY_04) QTY_04,SUM(QTY_05) QTY_05,SUM(QTY_06) QTY_06,SUM(QTY_07) QTY_07,SUM(QTY_08) QTY_08,SUM(QTY_09) QTY_09,SUM(QTY_10) QTY_10,SUM(QTY_11) QTY_11,SUM(QTY_12) QTY_12,SUM(QTY_SBI) QTY_SBI " & _
	                " from ( " & strMain & " )" & whereCons & "  group by  VHrut,DESCRIPTION,GroupPlantation " & selectCons & "  Order by VHrut,DESCRIPTION,GroupPlantation " & orderCons  
	'response.write strSolarDaily
	Set rsAttDaily = Server.CreateObject("ADODB.Recordset")
	rsAttDaily.Open strSolarDaily, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - LAPORAN BULANAN PEMAKAIAN SOLAR</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->
</head>

<body>
<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	</div>
<form name="frmEntry" method="post" action="p_SolarMonthly.asp">
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
		<th>&nbsp;</th>
		<th>BULAN</th>
		<th>TAHUN</th>	
		<th>GROUP</th>
		<th>KODE PEMAKAI</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<td align="center">
		<% 
					MM=Request("txtMonth")
					yyyy=Request("txtYear")

					mon = Month(Session("v_Costbook"))
					If len(MM)=0 then MM=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = 2015
					MaxYr = Yr  + 2
					If len(yyyy)= 0 then yyyy = Yr
				%>
		
		   <td>
				<font size="2" color="#0000FF" ></font>
					<select name="txtMonth">
						<%
							For i = 1 to 12
						%>
							<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
						<%
							Next
						%>
					</select>
		  </td>	
		   <td>
				<font size="2"  color="#0000FF" ></font>
				
				<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
				</select>
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
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
        <tr>
         <th colspan="18" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">

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
		<h2>LAPORAN BULANAN PEMAKAIAN SOLAR</h3>
		<h5>PERIODE : <%=v_txtmonth%> - <%=v_txtyear%> , KEBUN : <%=v_txtFcba%></h5>
	</div>
</div>

</th>
        </tr>
    
	<tr>
	
		<th style="background-color:#00A300;color:#fff;">GROUP PEMAKAI</th>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<% end if %>
		<% if v_txtVH <>  "--SUMMARY--" and v_txtVH <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;">KODE PEMAKAI</th>
	           <th style="background-color:#00A300;color:#fff;">NAMA PEMAKAI</th>
	    <% end if %>
		
		<th style="background-color:#00A300;color:#fff;" >JAN</th>
		<th style="background-color:#00A300;color:#fff;" >FEB</th>
		<th style="background-color:#00A300;color:#fff;" >MAR</th>
		<th style="background-color:#00A300;color:#fff;" >APR</th>
		<th style="background-color:#00A300;color:#fff;" >MAY</th>
		<th style="background-color:#00A300;color:#fff;" >JUN</th>
		<th style="background-color:#00A300;color:#fff;" >JUL</th>
		<th style="background-color:#00A300;color:#fff;" >AUG</th>
		<th style="background-color:#00A300;color:#fff;" >SEP</th>
		<th style="background-color:#00A300;color:#fff;" >OCT</th>
		<th style="background-color:#00A300;color:#fff;" >NOV</th>
		<th style="background-color:#00A300;color:#fff;" >DES</th>
		<th style="background-color:#00A300;color:#fff;" >SBI</th>
		
	</tr>
	</thead>
<tbody>

<%
Dim tempjob_forfieldstatus
Dim nTotal_1, nTotal_2, nTotal_3,nTotal_4, nTotal_5, nTotal_6,nTotal_7, nTotal_8, nTotal_9,nTotal_10, nTotal_11, nTotal_12,nTotal_SBI

	if not rsAttDaily.EOF then
		nTotal_1 = 0
		nTotal_2 = 0
		nTotal_3 = 0
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
							  <a href="p_SolarMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>">
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
							<a href="p_SolarMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>">
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
							<a href="p_SolarMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>">
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
							<a href="p_SolarMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>">
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
							<a href="p_SolarMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>">
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
						    <a href="p_SolarMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>">
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
							  <a href="p_SolarMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtGroupVH=<%=tempGroupVH%>&txtVH=<%=tempVH%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>">
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
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("QTY_SBI"),2)%></td>
			</tr>
<%
			NO = NO + 1
			nTotal_1 = nTotal_1 + Cdbl(rsAttDaily("QTY_01"))
			nTotal_2 = nTotal_2 + Cdbl(rsAttDaily("QTY_02"))
			nTotal_3 = nTotal_3 + Cdbl(rsAttDaily("QTY_03"))
			nTotal_4 = nTotal_4 + Cdbl(rsAttDaily("QTY_04"))
			nTotal_5 = nTotal_5 + Cdbl(rsAttDaily("QTY_05"))
			nTotal_6 = nTotal_6 + Cdbl(rsAttDaily("QTY_06"))
			nTotal_7 = nTotal_7 + Cdbl(rsAttDaily("QTY_07"))
			nTotal_8 = nTotal_8 + Cdbl(rsAttDaily("QTY_08"))
			nTotal_9 = nTotal_9 + Cdbl(rsAttDaily("QTY_09"))
			nTotal_10 = nTotal_10 + Cdbl(rsAttDaily("QTY_10"))
			nTotal_11 = nTotal_11 + Cdbl(rsAttDaily("QTY_11"))
			nTotal_12 = nTotal_12 + Cdbl(rsAttDaily("QTY_12"))
			nTotal_SBI = nTotal_SBI + Cdbl(rsAttDaily("QTY_SBI"))
			rsAttDaily.MoveNext
		Wend
	end if
	rsAttDaily.Close
	set rsAttDaily = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%= 2 + nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(nTotal_1,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_2,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_3,2)%></td>	
                <td align="right"><%=FormatNumber(nTotal_4,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_5,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_6,2)%></td>
                <td align="right"><%=FormatNumber(nTotal_7,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_8,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_9,2)%></td>
                <td align="right"><%=FormatNumber(nTotal_10,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_11,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_12,2)%></td>
                <td align="right"><%=FormatNumber(nTotal_SBI,2)%></td>				
			</tr>
        </tbody> 
</table>
</form>


</body>
</html>