<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsDailyVehicleSolar, strDailyVehicleSolar, strGroup, strOrder
	Dim NO, cClass
	Dim txtdatefrom, txtdateto
	Dim v_txtDateFrom, v_txtDateTo, v_txtVehicleGroup, v_txtVehicleCode, txtVehicleCode
	Dim nSpanTotal, strMain
	Dim sprevVehicleGroup, sprevVehicleCode
	Dim tempVehicleGroup, tempVehicleCode
	Dim sprevFcba, SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	Dim tempData_0, tempData_1, tempData_2, tempData_3,tempData_4, tempData_5, tempData_6, tempData_7,tempData_8, tempData_9, tempData_10, tempData_11
    Dim rscomp , strcomp

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
	 			
	nSpanTotal	= 8
	
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
	
	               ' strMain = 	"  SELECT " & strFieldCons & " " & _
								' "  vehiclegroup_api.get_OrderReport(C.VEHICLEGROUPCODE,a.fcba) VHrut, C.VEHICLEGROUPCODE VEHICLEGROUPCODE," & _
								' "  C.DESCRIPTION DESCRIPTION," & _
								' "  A.VEHICLECODE VEHICLECODE," & _
								' "  B.FCNAME VEHICLENAME," & _
								' "  A.UNIT UNIT," & _
								' "  A.HOURS HOURS," & _
								' "  A.SOLAR SOLAR," & _
								' "  C.ORDERREPORT ORDERREPORT," & _
								' "  ROUND (DECODE (A.UNIT, 0, 0, A.SOLAR / A.UNIT), 2) RASIO" & _
							' "  FROM VIEW_VEHICLEPRODUCTIVITY A, VEHICLE B, VEHICLEGROUP C" & _
								' "  WHERE " & _
										' "  A.ACTIONDATE BETWEEN TO_DATE ('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
										' "  AND TO_DATE ('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
										' "  AND A.VEHICLECODE = B.FCCODE" & _
										' "  AND B.VEHICLEGROUPCODE = C.VEHICLEGROUPCODE" & _
										' "  AND A.FCBA = B.FCBA" & _                  
										' "  AND B.FCBA = C.FCBA " & _ 
	                                    ' "  GROUP BY A.FCBA,C.VEHICLEGROUPCODE,C.DESCRIPTION,A.VEHICLECODE,B.FCNAME,A.UNIT,A.HOURS ,A.SOLAR ,C.ORDERREPORT"
										
	StrMain = " Select " & strFieldCons & " vehiclegroup_api.get_OrderReport(VEHICLEGROUPCODE,fcba) VHrut,VEHICLEGROUPCODE,DESCRIPTION,VEHICLECODE,VEHICLENAME, " & _ 
			  " sum(unit_hi) unit_hi,sum(hours_hi) hours_hi,sum(solar_hi) solar_hi " & _ 
			  "  ,sum(unit_shi) unit_shi,sum(hours_shi) hours_shi,sum(solar_shi) solar_shi " & _ 
			  "  ,sum(unit_sbi) unit_sbi,sum(hours_sbi) hours_sbi,sum(solar_sbi) solar_sbi " & _ 
			  " from (  " & _                                                        
			  " SELECT  a.fcba, " & _ 
				" 	   C.VEHICLEGROUPCODE VEHICLEGROUPCODE, " & _ 
				" 	   C.DESCRIPTION DESCRIPTION, " & _ 
				" 	   A.VEHICLECODE VEHICLECODE, " & _ 
				" 	   B.FCNAME VEHICLENAME, " & _ 
				" 	   sum(A.UNIT) UNIT_HI, " & _ 
				" 	   sum(A.HOURS) HOURS_HI, " & _ 
				" 	   sum(A.SOLAR) SOLAR_HI, " & _ 
				" 	   0 UNIT_SHI, " & _ 
				" 	   0 HOURS_SHI, " & _ 
				" 	   0 SOLAR_SHI, " & _ 
				" 	   0 UNIT_SBI, " & _ 
				" 	   0 HOURS_SBI, " & _ 
				" 	   0 SOLAR_SBI " & _ 
				"   FROM VIEW_VEHICLEPRODUCTIVITY A, VEHICLE B, VEHICLEGROUP C " & _ 
				"  WHERE A.ACTIONDATE = TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY') " & _ 
				"    AND A.VEHICLECODE = B.FCCODE " & _ 
				"    AND B.VEHICLEGROUPCODE = C.VEHICLEGROUPCODE " & _ 
				"    AND A.FCBA = B.FCBA " & _ 
				"    AND B.FCBA = C.FCBA " & _ 
				"  GROUP BY A.FCBA, " & _ 
				" 		  C.VEHICLEGROUPCODE, " & _ 
				" 		  C.DESCRIPTION," & _ 
				" 		  A.VEHICLECODE, " & _ 
				" 		  B.FCNAME " & _ 
				"   UNION " & _ 
				" 	   SELECT  a.fcba, " & _ 
				" 	   C.VEHICLEGROUPCODE VEHICLEGROUPCODE, " & _ 
				" 	   C.DESCRIPTION DESCRIPTION, " & _ 
				" 	   A.VEHICLECODE VEHICLECODE, " & _ 
				" 	   B.FCNAME VEHICLENAME, " & _ 
				" 	   0 UNIT_HI, " & _ 
				" 	   0 HOURS_HI, " & _ 
				" 	   0 SOLAR_HI, " & _ 
				" 	   sum(A.UNIT) UNIT_SHI, " & _ 
				" 	   sum(A.HOURS) HOURS_SHI, " & _ 
				" 	   sum(A.SOLAR) SOLAR_SHI, " & _ 
				" 	   0 UNIT_SBI, " & _ 
				" 	   0 HOURS_SBI, " & _ 
				" 	   0 SOLAR_SBI " & _ 
				"   FROM VIEW_VEHICLEPRODUCTIVITY A, VEHICLE B, VEHICLEGROUP C " & _ 
				"  WHERE A.ACTIONDATE <= TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY') " & _ 
				"    AND to_char(A.ACTIONDATE,'MMYYYY') = to_char(TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY'),'MMYYYY') " & _ 
				"    AND A.VEHICLECODE = B.FCCODE " & _ 
				"    AND B.VEHICLEGROUPCODE = C.VEHICLEGROUPCODE " & _ 
				"    AND A.FCBA = B.FCBA " & _ 
				"    AND B.FCBA = C.FCBA " & _ 
				"  GROUP BY A.FCBA, " & _ 
				" 		  C.VEHICLEGROUPCODE, " & _ 
				" 		  C.DESCRIPTION, " & _ 
				" 		  A.VEHICLECODE, " & _ 
				" 		  B.FCNAME " & _ 
				"   UNION " & _ 
				" 	   SELECT  a.fcba, " & _ 
				" 	   C.VEHICLEGROUPCODE VEHICLEGROUPCODE, " & _ 
				" 	   C.DESCRIPTION DESCRIPTION, " & _ 
				" 	   A.VEHICLECODE VEHICLECODE, " & _ 
				" 	   B.FCNAME VEHICLENAME, " & _ 
				" 	   0 UNIT_HI, " & _ 
				" 	   0 HOURS_HI, " & _ 
				" 	   0 SOLAR_HI, " & _ 
				" 	   0 UNIT_SHI, " & _ 
				" 	   0 HOURS_SHI, " & _ 
				" 	   0 SOLAR_SHI, " & _ 
				" 	   sum(A.UNIT) UNIT_SBI, " & _ 
				" 	   sum(A.HOURS) HOURS_SBI, " & _ 
				" 	   sum(A.SOLAR) SOLAR_SBI " & _ 
				"   FROM VIEW_VEHICLEPRODUCTIVITY A, VEHICLE B, VEHICLEGROUP C " & _ 
				"  WHERE A.ACTIONDATE <= TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY') " & _ 
				"    AND to_char(A.ACTIONDATE,'YYYY') = to_char(TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY'),'YYYY') " & _ 
				"    AND A.VEHICLECODE = B.FCCODE " & _ 
				"    AND B.VEHICLEGROUPCODE = C.VEHICLEGROUPCODE " & _ 
				"    AND A.FCBA = B.FCBA " & _ 
				"    AND B.FCBA = C.FCBA " & _ 
				"  GROUP BY A.FCBA, " & _ 
				" 		  C.VEHICLEGROUPCODE, " & _ 
				" 		  C.DESCRIPTION, " & _ 
				" 		  A.VEHICLECODE, " & _ 
				" 		  B.FCNAME       " & _ 
				" 		  )            " & _  
				" 	   group by FCBA,VEHICLEGROUPCODE,DESCRIPTION,VEHICLECODE,VEHICLENAME 	" 								
										
	strDailyVehicleSolar =  " , ROUND(SUM(NVL(UNIT_HI,0)),2) UNIT_HI  " & _ 
							" , ROUND(SUM(NVL(HOURS_HI,0)),2) HOURS_HI  " & _
							" , ROUND(SUM(NVL(SOLAR_HI,0)),2) SOLAR_HI  " & _	
							" , ROUND(DECODE(SUM(nvl(UNIT_HI,0)),0,0, SUM(nvl(SOLAR_HI,0))/SUM(nvl(UNIT_HI,0))),2) RASIO_HI  " & _
	                        " , ROUND(SUM(NVL(UNIT_SHI,0)),2) UNIT_SHI  " & _ 
							" , ROUND(SUM(NVL(HOURS_SHI,0)),2) HOURS_SHI  " & _
							" , ROUND(SUM(NVL(SOLAR_SHI,0)),2) SOLAR_SHI  " & _	
							" , ROUND(DECODE(SUM(nvl(UNIT_SHI,0)),0,0, SUM(nvl(SOLAR_SHI,0))/SUM(nvl(UNIT_SHI,0))),2) RASIO_SHI  " & _
							" , ROUND(SUM(NVL(UNIT_SBI,0)),2) UNIT_SBI  " & _ 
							" , ROUND(SUM(NVL(HOURS_SBI,0)),2) HOURS_SBI  " & _
							" , ROUND(SUM(NVL(SOLAR_SBI,0)),2) SOLAR_SBI  " & _	
							" , ROUND(DECODE(SUM(nvl(UNIT_SBI,0)),0,0, SUM(nvl(SOLAR_SBI,0))/SUM(nvl(UNIT_SBI,0))),2) RASIO_SBI  " 
	
	strDailyVehicleSolar = " SELECT VHrut,VEHICLEGROUPCODE,DESCRIPTION,GroupPlantation " & selectCons & strDailyVehicleSolar  & " from ( " & strMain & " )" & whereCons & "  group by  VHrut,VEHICLEGROUPCODE,DESCRIPTION,GroupPlantation " & selectCons & "  Order by VHrut,VEHICLEGROUPCODE,DESCRIPTION,GroupPlantation " & orderCons  


	
	' if len(trim(v_txtVehicleCode)) > 0 then 
		' if v_txtVehicleCode <> "--SUMMARY--" then
			' strDailyVehicleSolar = strDailyVehicleSolar & " , VEHICLECODE, VEHICLENAME "
		' end if
	' else
	    ' strDailyVehicleSolar = strDailyVehicleSolar & " , VEHICLECODE, VEHICLENAME "
	' end if	
	
	
	
	
	
	' strDailyVehicleSolar =  strDailyVehicleSolar & " , ROUND(SUM(NVL(UNIT,0)),2) UNIT  " & _ 
												   ' " , ROUND(SUM(NVL(HOURS,0)),2) HOURS  " & _
												   ' " , ROUND(SUM(NVL(SOLAR,0)),2) SOLAR  " & _	
												   ' " , ROUND(DECODE(SUM(nvl(UNIT,0)),0,0, SUM(nvl(SOLAR,0))/SUM(nvl(UNIT,0))),2) RASIO  " 
												   
	
    ' strDailyVehicleSolar = strDailyVehicleSolar & " FROM ( " & strMain & " ) "
	
	' ' Where Condition
	
	
	' if len(trim(v_txtVehicleGroup)) > 0 then 
		' if v_txtVehicleGroup <> "--SUMMARY--" then
			' strDailyVehicleSolar = strDailyVehicleSolar & " WHERE VEHICLEGROUPCODE = '" & v_txtVehicleGroup & "'"
		' end if
	' else
		' strDailyVehicleSolar = strDailyVehicleSolar & " WHERE VEHICLEGROUPCODE like '%' "
	' end if
	
	' if len(trim(v_txtVehicleCode)) > 0 then 
		' if v_txtVehicleCode <> "--SUMMARY--" then
			' strDailyVehicleSolar = strDailyVehicleSolar & " AND VEHICLECODE = '" & v_txtVehicleCode & "'"
		' end if
	' end if
	
	
	
	' ' Group By 
	' strDailyVehicleSolar = strDailyVehicleSolar & " GROUP BY ORDERREPORT, VEHICLEGROUPCODE, DESCRIPTION  "
	
	
	' if len(trim(v_txtVehicleCode)) > 0 then 
		' if v_txtVehicleCode <> "--SUMMARY--" then
			' strDailyVehicleSolar = strDailyVehicleSolar & " , VEHICLECODE, VEHICLENAME "
		' end if
	' else
	    ' strDailyVehicleSolar = strDailyVehicleSolar & " , VEHICLECODE, VEHICLENAME "
	' end if	
	
	' if len(trim(v_txtfcba)) > 0 then 
		' if v_txtfcba <> "--SUMMARY--" then
			' strDailyVehicleSolar = strDailyVehicleSolar & " , FCBA "
		' end if
	' else
	    ' strDailyVehicleSolar = strDailyVehicleSolar & " , FCBA "
	' end if	
	
	
	' ' Order By
	' strDailyVehicleSolar = strDailyVehicleSolar & " ORDER BY ORDERREPORT, VEHICLEGROUPCODE "
	
	
	' if len(trim(v_txtVehicleCode)) > 0 then 
		' if v_txtVehicleCode <> "--SUMMARY--" then
			' strDailyVehicleSolar = strDailyVehicleSolar & " , VEHICLECODE "
		' end if
	' else
	    ' strDailyVehicleSolar = strDailyVehicleSolar & " , VEHICLECODE "
	' end if	

	' if len(trim(v_txtfcba)) > 0 then 
		' if v_txtfcba <> "--SUMMARY--" then
			' strDailyVehicleSolar = strDailyVehicleSolar & " , FCBA "
		' end if
	' else
	    ' strDailyVehicleSolar = strDailyVehicleSolar & " , FCBA "
	' end if		

	
	Set rsDailyVehicleSolar = Server.CreateObject("ADODB.Recordset")
	'response.write strDailyVehicleSolar
	rsDailyVehicleSolar.Open strDailyVehicleSolar, connstr, adOpenStatic 
	
%>

<html>

<head>
<title>ePlaS Reporting - Laporan Harian Konsumsi Solar Kendaraan</title>
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
<form name="frmEntry" method="post" action="p_DailyVehicleSolar.asp">
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
		<th>VEHICLE GROUP</th>
		<th>VEHICLE CODE</th>
		<th>&nbsp;</th>
	</tr>
	<tr>

		
		
		<td align="center">
		<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<!--<a class="calendar" href="javascript:show_calendar('frmEntry.txtDateTo');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>-->
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
            <th colspan="15" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
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
                        <h2>LAPORAN HARIAN KONSUMSI SOLAR KENDARAAN</h3>
                        <h5>TANGGAL: <%=v_txtDateFrom%> s/d <%=v_txtDateTo%></h5>
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
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">PLANTATION</th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">GROUP REGION</th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">REGION</th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">COMPANY</th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">SITE</th>
		<% end if %>
	
	<% if v_txtVehicleCode <> "--SUMMARY--" and v_txtVehicleCode <> "" then %>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">KODE KENDARAAN</th>
		<th  rowspan="2"style="background-color:#00A300;color:#fff;">NAMA KENDARAAN</th>
	<% end if %>
    <th colspan="3" style="background-color:#00A300;color:#fff;">HARI INI (HI) </th> 	
	<th colspan="3" style="background-color:#00A300;color:#fff;">SAMPAI HARI INI (SHI)</th> 	
	<th colspan="3" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI(SBI)</th> 	
</tr>
<tr>	
	<th style="background-color:#00A300;color:#fff;">KM-HM-JAM</th>
	<th style="background-color:#00A300;color:#fff;">SOLAR</th>
	<th style="background-color:#00A300;color:#fff;">RASIO</th>
	<th style="background-color:#00A300;color:#fff;">KM-HM-JAM</th>
	<th style="background-color:#00A300;color:#fff;">SOLAR</th>
	<th style="background-color:#00A300;color:#fff;">RASIO</th>
	<th style="background-color:#00A300;color:#fff;">KM-HM-JAM</th>
	<th style="background-color:#00A300;color:#fff;">SOLAR</th>
	<th style="background-color:#00A300;color:#fff;">RASIO</th>
</tr>
</thead>
<tbody>
<%
Dim nTotal_0
Dim nTotal_1
Dim nTotal_2, nTotal_3,nTotal_6,nTotal_10
	if not rsDailyVehicleSolar.EOF then
		nTotal_0 = 0
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
							<a href="p_DailyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
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
							<a href="p_DailyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
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
							<a href="p_DailyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
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
							<a href="p_DailyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
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
							<a href="p_DailyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
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
						    <a href="p_DailyVehicleSolar.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
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
							<a href="p_DailyVehicleSolar.asp?txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
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
				
				
				
				<% 
				  ' <<< HI >>>
				  if rsDailyVehicleSolar("UNIT_HI") <> "" then 
				       tempData_0 = cdbl(rsDailyVehicleSolar("UNIT_HI"))
				   else 
					   tempData_0 = 0 
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempData_0,0)%></%></td>

				
				
				<% if rsDailyVehicleSolar("SOLAR_HI") <> "" then 
				       tempData_2 = cdbl(rsDailyVehicleSolar("SOLAR_HI"))
				   else 
				       tempData_2 = 0 
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempData_2,0)%></%></td>
				
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempData_3 = cdbl(rsDailyVehicleSolar("RASIO_HI"))
				   else 
				       if cdbl(tempData_2) > 0 then 
				          tempData_3 = cdbl(tempData_0) / cdbl(tempData_2)
					   else
					      tempData_3 = 0
                       end if 					   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempData_3,2)%></%></td>
			
			
			<% 
				  ' <<< SHI >>>
				  if rsDailyVehicleSolar("UNIT_SHI") <> "" then 
				       tempData_4 = cdbl(rsDailyVehicleSolar("UNIT_SHI"))
				   else 
					   tempData_4 = 0 
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempData_4,0)%></%></td>

				
				
				<% if rsDailyVehicleSolar("SOLAR_SHI") <> "" then 
				       tempData_6 = cdbl(rsDailyVehicleSolar("SOLAR_SHI"))
				   else 
				       tempData_6 = 0 
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempData_6,0)%></%></td>
				
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempData_7 = cdbl(rsDailyVehicleSolar("RASIO_SHI"))
				   else 
				       if cdbl(tempData_6) > 0 then 
				          tempData_7 = cdbl(tempData_4) / cdbl(tempData_6)
					   else
					      tempData_7 = 0
                       end if 					   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempData_7,2)%></%></td>
			
			<% 
				  ' <<< SBI >>>
				  if rsDailyVehicleSolar("UNIT_SBI") <> "" then 
				       tempData_8 = cdbl(rsDailyVehicleSolar("UNIT_SBI"))
				   else 
					   tempData_8 = 0 
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempData_8,0)%></%></td>

				
				
				<% if rsDailyVehicleSolar("SOLAR_SBI") <> "" then 
				       tempData_10 = cdbl(rsDailyVehicleSolar("SOLAR_SBI"))
				   else 
				       tempData_10 = 0 
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempData_10,0)%></%></td>
				
				<% if tempVehicleGroup <> "MB" AND  tempVehicleGroup <> "AB" AND  tempVehicleGroup <> "BS" AND  tempVehicleGroup <> "DT"  AND  tempVehicleGroup <> "TK" then 
				       tempData_11 = cdbl(rsDailyVehicleSolar("RASIO_SBI"))
				   else 
				       if cdbl(tempData_10) > 0 then 
				          tempData_11 = cdbl(tempData_8) / cdbl(tempData_10)
					   else
					      tempData_11 = 0
                       end if 					   
				   end if
				 %>
				<td align="right"><%=FormatNumber(tempData_11,2)%></%></td>
			</tr>
			
<%
			NO = NO + 1
			nTotal_0 = nTotal_0 + tempData_0
			
			nTotal_2 = nTotal_2 + tempData_2
			nTotal_6 = nTotal_6 + tempData_6
			nTotal_10 = nTotal_10 + tempData_10
			nTotal_3 = nTotal_3 + tempData_3
			
			rsDailyVehicleSolar.MoveNext
		Wend
	end if
	rsDailyVehicleSolar.Close
	set rsDailyVehicleSolar = Nothing
%>
           <tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%= nSpanTotal %>">TOTAL</td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_2,0)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_6,0)%></td>
				<td align="right"></td>
				<td align="right"></td>
				<td align="right"><%=FormatNumber(nTotal_10,0)%></td>
				<td align="right"></td>
			</tr>
		</tbody> 	
</table>
</form>



