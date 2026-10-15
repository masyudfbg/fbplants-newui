<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsDailyVehicleSolar, strDailyVehicleSolar, strGroup, strOrder
	Dim cClass
	Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth,txtdatefrom, txtdateto
	Dim v_txtDateFrom, v_txtDateTo, v_txtVehicleGroup, v_txtVehicleCode, txtVehicleCode,v_TxtTypeFCBA 
	Dim nSpanTotal, strMain
	Dim sprevVehicleGroup, sprevVehicleCode
	Dim tempVehicleGroup, tempVehicleCode
	Dim sprevFcba, SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	Dim tempData_0, tempData_1, tempData_2, tempData_3,tempData_4, tempData_5, tempData_6, tempData_7,tempData_8, tempData_9, tempData_10, tempData_11
    Dim rscomp , strcomp
    Dim tempRasio_01,tempRasio_02,tempRasio_03,tempRasio_04,tempRasio_05,tempRasio_06,tempRasio_07,tempRasio_08,tempRasio_09,tempRasio_10,tempRasio_11,tempRasio_12,tempRasio_SBI
	
	v_txtGroupPlantation   = Request("txtGroupPlantation")
	v_txtPlantation  	   = Request("txtPlantation")
	v_txtGroupRegion  	   = Request("txtGroupRegion")
	v_txtRegion  		   = Request("txtRegion")
	v_txtcompany  		   = Request("txtCompany")
    v_txtfcba			   = Request("txtfcba")
	v_txtDateFrom		   = Request("txtdatefrom")
	v_txtDateTo			   = Request("txtdateto")
	v_txtVehicleGroup	   = Request("txtVehicleGroup")
	v_txtVehicleCode	   = Request("txtVehicleCode")
	v_TxtTypeFCBA          = Request("TxtTypeFCBA")
	if Request("TxtTypeFCBA") <> "" then
	   v_TxtTypeFCBA          = Request("TxtTypeFCBA")
    end if 	
	nSpanTotal	= 8
	       v_txtyear 		= Request("txtyear")
		   v_txtmonth 		= lpad(Request("txtmonth"),"0",2)
	
	 Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',FCBA) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',FCBA) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',FCBA) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',FCBA) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',FCBA) company, BUSINESSUNIT_API.get_OrderData(FCBA) norut, FCBA, " 

						   
			wherecons = " where company is not null "
			
			if v_txtVehicleGroup <> "--SUMMARY--" AND v_txtVehicleGroup <> "" then
			      if v_txtVehicleGroup <> "--ALL--" then
					 whereCons = whereCons & " And  VEHICLEGROUPCODE = '" & v_txtVehicleGroup & "'"
				  end if
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
	        if  v_txtVehicleCode <> "--SUMMARY--" And v_txtVehicleCode <> "" then
			     if v_txtVehicleCode <> "--ALL--" then
					whereCons = whereCons & " AND VEHICLECODE = '" & v_txtVehicleCode & "'"
				 end if
				 selectCons = selectCons & " ,VEHICLECODE,VEHICLENAME "
				 orderCons  = orderCons & ",VEHICLECODE,VEHICLENAME "
			else
			    nSpanTotal = nSpanTotal - 2 
			end if
	
	              
										
	StrMain = " SELECT   " & strFieldCons & "  vehiclegroup_api.get_OrderReport(VEHICLEGROUPCODE,fcba) VHrut,VEHICLEGROUPCODE,DESCRIPTION,VEHICLECODE,VEHICLENAME, " & _ 
	          "   UNIT_01,UNIT_02,UNIT_03,UNIT_04,UNIT_05,UNIT_06,UNIT_07,UNIT_08,UNIT_09,UNIT_10,UNIT_11,UNIT_12, " & _ 
			  "   (UNIT_01+UNIT_02+UNIT_03+UNIT_04+UNIT_05+UNIT_06+UNIT_07+UNIT_08+UNIT_09+UNIT_10+UNIT_11+UNIT_12) UNIT_SBI, " & _ 
			  "   SOLAR_01,SOLAR_02,SOLAR_03,SOLAR_04,SOLAR_05,SOLAR_06,SOLAR_07,SOLAR_08,SOLAR_09,SOLAR_10,SOLAR_11,SOLAR_12," & _ 
			  "   (SOLAR_01+SOLAR_02+SOLAR_03+SOLAR_04+SOLAR_05+SOLAR_06+SOLAR_07+SOLAR_08+SOLAR_09+SOLAR_10+SOLAR_11+SOLAR_12) SOLAR_SBI" & _ 
	          " From (  Select A.FCBA, C.VEHICLEGROUPCODE VEHICLEGROUPCODE, C.DESCRIPTION,A.VEHICLECODE VEHICLECODE,B.FCNAME VEHICLENAME,  " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),1,sum(NVL(A.UNIT,0)),0) UNIT_01, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),2,sum(NVL(A.UNIT,0)),0) UNIT_02, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),3,sum(NVL(A.UNIT,0)),0) UNIT_03, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),4,sum(NVL(A.UNIT,0)),0) UNIT_04, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),5,sum(NVL(A.UNIT,0)),0) UNIT_05, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),6,sum(NVL(A.UNIT,0)),0) UNIT_06, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),7,sum(NVL(A.UNIT,0)),0) UNIT_07, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),8,sum(NVL(A.UNIT,0)),0) UNIT_08, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),9,sum(NVL(A.UNIT,0)),0) UNIT_09, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),10,sum(NVL(A.UNIT,0)),0) UNIT_10, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),11,sum(NVL(A.UNIT,0)),0) UNIT_11, " & _   
			  "			  decode(to_char(A.ACTIONDATE,'MM'),12,sum(NVL(A.UNIT,0)),0) UNIT_12, " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),1,sum(NVL(A.SOLAR,0)),0) SOLAR_01, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),2,sum(NVL(A.SOLAR,0)),0) SOLAR_02 ,  " & _ 
			  "			  decode(to_char(A.ACTIONDATE,'MM'),3,sum(NVL(A.SOLAR,0)),0) SOLAR_03, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),4,sum(NVL(A.SOLAR,0)),0) SOLAR_04, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),5,sum(NVL(A.SOLAR,0)),0) SOLAR_05, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),6,sum(NVL(A.SOLAR,0)),0) SOLAR_06, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),7,sum(NVL(A.SOLAR,0)),0) SOLAR_07, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),8,sum(NVL(A.SOLAR,0)),0) SOLAR_08, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),9,sum(NVL(A.SOLAR,0)),0) SOLAR_09, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),10,sum(NVL(A.SOLAR,0)),0) SOLAR_10, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),11,sum(NVL(A.SOLAR,0)),0) SOLAR_11, " & _  
			  "			  decode(to_char(A.ACTIONDATE,'MM'),12,sum(NVL(A.SOLAR,0)),0) SOLAR_12 " & _ 
			  " 		   FROM VIEW_VEHICLEPRODUCTIVITY A, VEHICLE B, VEHICLEGROUP C  " & _ 
			  "		  WHERE TO_CHAR(A.ACTIONDATE,'YYYYMM') <= '" & v_txtyear & v_txtmonth &"'" & _
			  "			AND TO_CHAR(A.ACTIONDATE,'YYYY') = '" & v_txtyear &"'" & _ 
			  "			AND A.VEHICLECODE = B.FCCODE  " & _ 
			  "			AND B.VEHICLEGROUPCODE = C.VEHICLEGROUPCODE  " & _ 
			  "			AND A.FCBA = B.FCBA  " & _ 
			  "			AND B.FCBA = C.FCBA  " & _ 
			  " 	  GROUP BY A.FCBA, C.VEHICLEGROUPCODE,  C.DESCRIPTION,A.VEHICLECODE,B.FCNAME,to_char(A.ACTIONDATE,'MM')  " & _ 
			  "    ) WHERE USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = '" & v_TxtTypeFCBA  &"' " 							
										
	strDailyVehicleSolar =  " , SUM(UNIT_01) UNIT_01, SUM(SOLAR_01) SOLAR_01,DECODE(SUM(UNIT_01),0,0, SUM(SOLAR_01)/SUM(UNIT_01)) RASIO_SU_01 ,DECODE(SUM(SOLAR_01),0,0, SUM(UNIT_01)/SUM(SOLAR_01)) RASIO_US_01  " & _
	                        " , SUM(UNIT_02) UNIT_02, SUM(SOLAR_02) SOLAR_02,DECODE(SUM(UNIT_02),0,0, SUM(SOLAR_02)/SUM(UNIT_02)) RASIO_SU_02 ,DECODE(SUM(SOLAR_02),0,0, SUM(UNIT_02)/SUM(SOLAR_02)) RASIO_US_02  " & _
							" , SUM(UNIT_03) UNIT_03, SUM(SOLAR_03) SOLAR_03,DECODE(SUM(UNIT_03),0,0, SUM(SOLAR_03)/SUM(UNIT_03)) RASIO_SU_03 ,DECODE(SUM(SOLAR_03),0,0, SUM(UNIT_03)/SUM(SOLAR_03)) RASIO_US_03   " & _
							" , SUM(UNIT_04) UNIT_04, SUM(SOLAR_04) SOLAR_04,DECODE(SUM(UNIT_04),0,0, SUM(SOLAR_04)/SUM(UNIT_04)) RASIO_SU_04 ,DECODE(SUM(SOLAR_04),0,0, SUM(UNIT_04)/SUM(SOLAR_04)) RASIO_US_04   " & _	
							" , SUM(UNIT_05) UNIT_05, SUM(SOLAR_05) SOLAR_05,DECODE(SUM(UNIT_05),0,0, SUM(SOLAR_05)/SUM(UNIT_05)) RASIO_SU_05 ,DECODE(SUM(SOLAR_05),0,0, SUM(UNIT_05)/SUM(SOLAR_05)) RASIO_US_05   " & _
							" , SUM(UNIT_06) UNIT_06, SUM(SOLAR_06) SOLAR_06,DECODE(SUM(UNIT_06),0,0, SUM(SOLAR_06)/SUM(UNIT_06)) RASIO_SU_06 ,DECODE(SUM(SOLAR_06),0,0, SUM(UNIT_06)/SUM(SOLAR_06)) RASIO_US_06  " & _ 
							" , SUM(UNIT_07) UNIT_07, SUM(SOLAR_07) SOLAR_07,DECODE(SUM(UNIT_07),0,0, SUM(SOLAR_07)/SUM(UNIT_07)) RASIO_SU_07 ,DECODE(SUM(SOLAR_07),0,0, SUM(UNIT_07)/SUM(SOLAR_07)) RASIO_US_07   " & _
							" , SUM(UNIT_08) UNIT_08, SUM(SOLAR_08) SOLAR_08,DECODE(SUM(UNIT_08),0,0, SUM(SOLAR_08)/SUM(UNIT_08)) RASIO_SU_08 ,DECODE(SUM(SOLAR_08),0,0, SUM(UNIT_08)/SUM(SOLAR_08)) RASIO_US_08   " & _	
							" , SUM(UNIT_09) UNIT_09, SUM(SOLAR_09) SOLAR_09,DECODE(SUM(UNIT_09),0,0, SUM(SOLAR_09)/SUM(UNIT_09)) RASIO_SU_09 ,DECODE(SUM(SOLAR_09),0,0, SUM(UNIT_09)/SUM(SOLAR_09)) RASIO_US_09   " & _
							" , SUM(UNIT_10) UNIT_10, SUM(SOLAR_10) SOLAR_10,DECODE(SUM(UNIT_10),0,0, SUM(SOLAR_10)/SUM(UNIT_10)) RASIO_SU_10 ,DECODE(SUM(SOLAR_10),0,0, SUM(UNIT_10)/SUM(SOLAR_10)) RASIO_US_10   " & _
							" , SUM(UNIT_11) UNIT_11, SUM(SOLAR_11) SOLAR_11,DECODE(SUM(UNIT_11),0,0, SUM(SOLAR_11)/SUM(UNIT_11)) RASIO_SU_11 ,DECODE(SUM(SOLAR_11),0,0, SUM(UNIT_11)/SUM(SOLAR_11)) RASIO_US_11  " & _
							" , SUM(UNIT_12) UNIT_12, SUM(SOLAR_12) SOLAR_12,DECODE(SUM(UNIT_12),0,0, SUM(SOLAR_12)/SUM(UNIT_12)) RASIO_SU_12 ,DECODE(SUM(SOLAR_12),0,0, SUM(UNIT_12)/SUM(SOLAR_12)) RASIO_US_12   " & _
							" , SUM(UNIT_SBI) UNIT_SBI, SUM(SOLAR_SBI) SOLAR_SBI,DECODE(SUM(UNIT_SBI),0,0, SUM(SOLAR_SBI)/SUM(UNIT_SBI)) RASIO_SU_SBI ,DECODE(SUM(SOLAR_SBI),0,0,SUM(UNIT_SBI)/SUM(SOLAR_SBI)) RASIO_US_SBI " 
	
	strDailyVehicleSolar = " SELECT VHrut,VEHICLEGROUPCODE,DESCRIPTION,GroupPlantation " & selectCons & strDailyVehicleSolar  & " from ( " & strMain & " )" & whereCons & "  group by  VHrut,VEHICLEGROUPCODE,DESCRIPTION,GroupPlantation " & selectCons & "  Order by VHrut,VEHICLEGROUPCODE,DESCRIPTION,GroupPlantation " & orderCons  


	
	
	Set rsDailyVehicleSolar = Server.CreateObject("ADODB.Recordset")
	'response.write strDailyVehicleSolar
	rsDailyVehicleSolar.Open strDailyVehicleSolar, connstr, adOpenStatic 
	
%>

<html>

<head>
<title>ePlaS Reporting - Laporan Bulanan Konsumsi Solar Kendaraan</title>
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
<form name="frmEntry" method="post" action="p_MonthlyVehicleSolar.asp">
	<div id="wrapper"> 
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>

	
	</div>
	<input type="hidden" name="TxtTypeFCBA" value="<%=v_TxtTypeFCBA%>" >
<!--#include file="FormSelectNewSumME.html"-->	
<table id="form-table" class="centered-margin">
	<tr>
		<th>&nbsp;</th>
		<th>BULAN</th>
		<th>TAHUN</th>	
		<th>VEHICLE GROUP</th>
		<th>VEHICLE CODE</th>
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
		
		
		

		<td>
				<select name="txtVehicleGroup">	
				<option value="">--ALL--</option>
		<%
			Dim rsVehicleGroup
			Set rsVehicleGroup = Server.CreateObject("ADODB.Recordset")
			rsVehicleGroup.CursorLocation = adUseClient
			strsql =  " SELECT DISTINCT VehicleGroupCode, Description FROM VehicleGroup WHERE user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null ORDER BY Description ASC"
			rsVehicleGroup.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsVehicleGroup.eof then
				DO while not rsVehicleGroup.eof
				if v_txtVehicleGroup = rsVehicleGroup("VehicleGroupCode") then
		%>			<option value="<%=rsVehicleGroup("VehicleGroupCode")%>" Selected><%=rsVehicleGroup("VehicleGroupCode")%> - <%=rsVehicleGroup("Description")%></option>
				<%else%>
					<option value="<%=rsVehicleGroup("VehicleGroupCode")%>"><%=rsVehicleGroup("VehicleGroupCode")%> - <%=rsVehicleGroup("Description")%></option>
				<%end if%>
		<%			rsVehicleGroup.movenext
				loop
			End if
			rsVehicleGroup.Close
			set rsVehicleGroup = nothing
		%>
		</select>		
		</td>
		  <td>
				<select name="txtVehicleCode">	
					 <option value="--SUMMARY--">--SUMMARY--</option>
						<%
							Dim rsVehicleCode
							Set rsVehicleCode = Server.CreateObject("ADODB.Recordset")
							rsVehicleCode.CursorLocation = adUseClient
							strsql =  " SELECT '--ALL--' as FCCODE, '' as FCNAME from DUAL UNION  SELECT DISTINCT FCCODE, FCNAME FROM Vehicle WHERE user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null ORDER BY FCNAME ASC"
							rsVehicleCode.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsVehicleCode.eof then
								DO while not rsVehicleCode.eof
								if v_txtVehicleCode = rsVehicleCode("FCCODE") then
						%>			<option value="<%=rsVehicleCode("FCCODE")%>" Selected><%=rsVehicleCode("FCCODE")%> - <%=rsVehicleCode("FCNAME")%></option>
								<%else%>
									<option value="<%=rsVehicleCode("FCCODE")%>"><%=rsVehicleCode("FCCODE")%> - <%=rsVehicleCode("FCNAME")%></option>
								<%end if%>
						<%			rsVehicleCode.movenext
								loop
							End if
							rsVehicleCode.Close
							set rsVehicleCode = nothing
						%>
		     </select>		
		</td>
		
		
		
		
		
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
</table>

<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
        <tr>
            <th colspan="45" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
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
                        <h2>LAPORAN BULANAN KONSUMSI SOLAR KENDARAAN</h3>
                        <h5>PERIODE : <%=v_txtmonth%> - <%=v_txtyear%> , KEBUN/PABRIK : <%=v_txtFcba%></h5>
                    </div>
                </div>
			</th>
        </tr>
    
	<tr>
	
	
	<% if trim(request("txtVehicleGroup")) <> "--SUMMARY--" then %>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">GROUP KENDARAAN</th>
	<% else
		nSpanTotal = nSpanTotal - 1 		
	end if %>		
    
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<% end if %>
	
	<% if v_txtVehicleCode <> "--SUMMARY--" and v_txtVehicleCode <> "" then %>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">KODE KENDARAAN</th>
		<th  rowspan="2"style="background-color:#00A300;color:#fff;">NAMA KENDARAAN</th>
	<% end if %>
	
	    <th colspan="3" style="background-color:#00A300;color:#fff;" >JAN</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >FEB</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >MAR</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >APR</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >MAY</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >JUN</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >JUL</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >AUG</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >SEP</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >OCT</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >NOV</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >DES</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;" >SBI</th>	
</tr>
<tr>	
    <% DIM K
		   FOR K = 1 TO 13 %>
	<th style="background-color:#00A300;color:#fff;">KM-HM-JAM</th>
	<th style="background-color:#00A300;color:#fff;">SOLAR</th>
	<th style="background-color:#00A300;color:#fff;">RASIO</th>
	<% NEXT %>	
</tr>
</thead>
<tbody>
<%
Dim nTotal_1,nTotal_2,nTotal_3,nTotal_4,nTotal_5,nTotal_6,nTotal_7,nTotal_8,nTotal_9,nTotal_10,nTotal_11,nTotal_12,nTotal_sbi
	if not rsDailyVehicleSolar.EOF then

		nTotal_1 = 0
		nTotal_2 = 0
		nTotal_3 = 0
		NO = 1
		
		while not rsDailyVehicleSolar.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
			        <%     tempVehicleGroup = rsDailyVehicleSolar("VEHICLEGROUPCODE")
						   if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
								tempPlantation = rsDailyVehicleSolar("Plantation")
							else
								tempPlantation = "--SUMMARY--"
							end if
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempVehicleCode ="--SUMMARY--"
					%>
				    <% if sprevVehicleGroup <> rsDailyVehicleSolar("VEHICLEGROUPCODE") then %>  
						<td align="left">
							<a href="p_MonthlyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>&TxtTypeFCBA=<%=v_TxtTypeFCBA%>">
							<%=rsDailyVehicleSolar("VEHICLEGROUPCODE")%> - <%=rsDailyVehicleSolar("DESCRIPTION")%>
							</a>
						</td>	
						<% sprevVehicleGroup = rsDailyVehicleSolar("VEHICLEGROUPCODE") %>	
					<% else	%>
						<td align="left"></td>					
					<% end if %> 
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     tempVehicleGroup = rsDailyVehicleSolar("VEHICLEGROUPCODE")
						   tempPlantation = rsDailyVehicleSolar("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempVehicleCode ="--SUMMARY--"
					%>
					<% if SprevPlantation <> rsDailyVehicleSolar("Plantation")   then %>
						<td align="center">
							<a href="p_MonthlyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>&TxtTypeFCBA=<%=v_TxtTypeFCBA%>">
							<%=rsDailyVehicleSolar("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsDailyVehicleSolar("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  tempVehicleGroup = rsDailyVehicleSolar("VEHICLEGROUPCODE")
					    
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsDailyVehicleSolar("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsDailyVehicleSolar("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempVehicleCode = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsDailyVehicleSolar("GroupRegion") then %>
						<td align="center">
							<a href="p_MonthlyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>&TxtTypeFCBA=<%=v_TxtTypeFCBA%>">
							<%=rsDailyVehicleSolar("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsDailyVehicleSolar("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% tempVehicleGroup = rsDailyVehicleSolar("VEHICLEGROUPCODE")
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsDailyVehicleSolar("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsDailyVehicleSolar("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsDailyVehicleSolar("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempVehicleCode= "--SUMMARY--"
					   
					%>
					
					<% if SprevRegion <> rsDailyVehicleSolar("Region") then %>
						<td align="center">
							<a href="p_MonthlyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>&TxtTypeFCBA=<%=v_TxtTypeFCBA%>">
							<%=rsDailyVehicleSolar("region")%>
							</a>
						</td>
						<% SprevRegion = rsDailyVehicleSolar("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   tempVehicleGroup = rsDailyVehicleSolar("VEHICLEGROUPCODE")
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsDailyVehicleSolar("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsDailyVehicleSolar("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsDailyVehicleSolar("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsDailyVehicleSolar("Region") 
					   else
					       tempRegion = ""
					   end if	
					   tempCompany =rsDailyVehicleSolar("Company")
					   tempfcba = "--ALL--"	
					   tempVehicleCode  ="--SUMMARY--"
					 %>
					
					<% if SprevCompany <> rsDailyVehicleSolar("Company") then %>
						<td align="center">
							<a href="p_MonthlyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>&TxtTypeFCBA=<%=v_TxtTypeFCBA%>">
							<%=rsDailyVehicleSolar("company")%>
							</a>
						</td>
						<% SprevCompany = rsDailyVehicleSolar("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
					 <% 
					   tempVehicleGroup = rsDailyVehicleSolar("VEHICLEGROUPCODE") 
					   					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsDailyVehicleSolar("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsDailyVehicleSolar("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsDailyVehicleSolar("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtcompany <>  "--SUMMARY--" And v_txtcompany <>  "" then
					      tempCompany = rsDailyVehicleSolar("Company")
					   else
					      tempCompany = ""
                       end if				   
					   tempfcba = rsDailyVehicleSolar("fcba") 
					   tempVehicleCode ="--ALL--"					   
					 %>
					<% if sprevFcba <> rsDailyVehicleSolar("fcba") then %>
						<td align="center">
						    <a href="p_MonthlyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>&TxtTypeFCBA=<%=v_TxtTypeFCBA%>">
							<%=rsDailyVehicleSolar("fcba")%>
						</td>
						<% sprevFcba = rsDailyVehicleSolar("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtVehicleCode <> "--SUMMARY--" and v_txtVehicleCode <> "" then %>	
                   <%
					   tempVehicleGroup = rsDailyVehicleSolar("VEHICLEGROUPCODE") 
					   					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsDailyVehicleSolar("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsDailyVehicleSolar("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsDailyVehicleSolar("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtcompany <>  "--SUMMARY--" And v_txtcompany <>  "" then
					      tempCompany = rsDailyVehicleSolar("Company")
					   else
					      tempCompany = ""
                       end if
                       if v_txtfcba <> "--SUMMARY--" and v_txtfcba <> ""	then				   
					     tempfcba = rsDailyVehicleSolar("fcba") 
					   else
					      tempfcba= ""
                       end if	 
					   tempVehicleCode =rsDailyVehicleSolar("VEHICLECODE")					   
					   %>
				   
					<% if sprevVehicleCode <> rsDailyVehicleSolar("VEHICLECODE") then %>
					
						<td align="left">
							<a href="p_MonthlyVehicleSolar.asp?txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
							<%=rsDailyVehicleSolar("VEHICLECODE")%>
							</a>
						</td>
						<td align="left">
							<%=rsDailyVehicleSolar("VEHICLENAME")%>
						</td>
						<% sprevVehicleCode = rsDailyVehicleSolar("VEHICLECODE") %>	
					<% else	%>
						<td align="left"></td>					
						<td align="left"></td>					
					<% end if %>	
				<% end if %>
				
				
				
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_01"),2)%></td>
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_01 = cdbl(rsDailyVehicleSolar("RASIO_SU_01"))
				   else 
				       tempRasio_01 = cdbl(rsDailyVehicleSolar("RASIO_US_01"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_01,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_02"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_02 = cdbl(rsDailyVehicleSolar("RASIO_SU_02"))
				   else 
				       tempRasio_02 = cdbl(rsDailyVehicleSolar("RASIO_US_02"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_02,2)%></%></td>
			    
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_03"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_03 = cdbl(rsDailyVehicleSolar("RASIO_SU_03"))
				   else 
				       tempRasio_03 = cdbl(rsDailyVehicleSolar("RASIO_US_03"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_03,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_04"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_04 = cdbl(rsDailyVehicleSolar("RASIO_SU_04"))
				   else 
				       tempRasio_04 = cdbl(rsDailyVehicleSolar("RASIO_US_04"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_04,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_05"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_05 = cdbl(rsDailyVehicleSolar("RASIO_SU_05"))
				   else 
				       tempRasio_05 = cdbl(rsDailyVehicleSolar("RASIO_US_05"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_05,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_06"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_06 = cdbl(rsDailyVehicleSolar("RASIO_SU_06"))
				   else 
				       tempRasio_06 = cdbl(rsDailyVehicleSolar("RASIO_US_06"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_06,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_07"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_07 = cdbl(rsDailyVehicleSolar("RASIO_SU_07"))
				   else 
				       tempRasio_07 = cdbl(rsDailyVehicleSolar("RASIO_US_07"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_07,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_08"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_08 = cdbl(rsDailyVehicleSolar("RASIO_SU_08"))
				   else 
				       tempRasio_08 = cdbl(rsDailyVehicleSolar("RASIO_US_08"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_08,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_09"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_09 = cdbl(rsDailyVehicleSolar("RASIO_SU_09"))
				   else 
				       tempRasio_09 = cdbl(rsDailyVehicleSolar("RASIO_US_09"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_09,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_10"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_10 = cdbl(rsDailyVehicleSolar("RASIO_SU_10"))
				   else 
				       tempRasio_10 = cdbl(rsDailyVehicleSolar("RASIO_US_10"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_10,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_11"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_11 = cdbl(rsDailyVehicleSolar("RASIO_SU_11"))
				   else 
				       tempRasio_11 = cdbl(rsDailyVehicleSolar("RASIO_US_11"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_12,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_12"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_12 = cdbl(rsDailyVehicleSolar("RASIO_SU_12"))
				   else 
				       tempRasio_12 = cdbl(rsDailyVehicleSolar("RASIO_US_12"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_12,2)%></%></td>
				
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("UNIT_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsDailyVehicleSolar("SOLAR_SBI"),2)%></td>	
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempRasio_SBI = cdbl(rsDailyVehicleSolar("RASIO_SU_SBI"))
				   else 
				       tempRasio_SBI = cdbl(rsDailyVehicleSolar("RASIO_US_SBI"))		   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempRasio_SBI,2)%></%></td>
			</tr>
			
<%
			NO = NO + 1
			nTotal_1 = nTotal_1 + cdbl(rsDailyVehicleSolar("SOLAR_01"))
			nTotal_2 = nTotal_2 + cdbl(rsDailyVehicleSolar("SOLAR_02"))
			nTotal_3 = nTotal_3 + cdbl(rsDailyVehicleSolar("SOLAR_03"))
			nTotal_4 = nTotal_4 + cdbl(rsDailyVehicleSolar("SOLAR_04"))
			nTotal_5 = nTotal_5 + cdbl(rsDailyVehicleSolar("SOLAR_05"))
			nTotal_6 = nTotal_6 + cdbl(rsDailyVehicleSolar("SOLAR_06"))
			nTotal_7 = nTotal_7 + cdbl(rsDailyVehicleSolar("SOLAR_07"))
			nTotal_8 = nTotal_8 + cdbl(rsDailyVehicleSolar("SOLAR_08"))
			nTotal_9 = nTotal_9 + cdbl(rsDailyVehicleSolar("SOLAR_09"))
			nTotal_10 = nTotal_10 + cdbl(rsDailyVehicleSolar("SOLAR_10"))
			nTotal_11 = nTotal_11 + cdbl(rsDailyVehicleSolar("SOLAR_11"))
			nTotal_12 = nTotal_12 + cdbl(rsDailyVehicleSolar("SOLAR_12")) 
			nTotal_sbi = nTotal_sbi + cdbl(rsDailyVehicleSolar("SOLAR_SBI"))
			
			rsDailyVehicleSolar.MoveNext
		Wend
	end if
	rsDailyVehicleSolar.Close
	set rsDailyVehicleSolar = Nothing
%>
           <tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%= nSpanTotal %>">TOTAL</td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_1,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_2,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_3,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_4,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_5,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_6,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_7,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_8,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_9,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_10,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_11,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_12,2)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_sbi,2)%></td>
				<td align="right"></td>
			</tr>
		</tbody> 	
</table>
</form>



