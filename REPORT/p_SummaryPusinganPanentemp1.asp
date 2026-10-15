<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtDateFrom, v_txtDateFrom2, v_txtDateTo2, v_txtDateTo,v_txtSupplier, CropName, FcbaType
Dim ScriptName, cClass, wherecons, selectCons, orderCons, q, CompanyName, fcbaname, MgrName, KTUName
Dim nBRUTO, nSORTASI, nNETTO, selectifkebun, selectifpabrik, v_txtfcba2

 if Trim(Request("txtDateFrom")) <> "" then
    v_txtDateFrom = Trim(Request("txtDateFrom"))
 else
   v_txtDateFrom = textDate(date)
 end if
if Trim(Request("txtDateTo")) <> "" then
    v_txtDateTo = Trim(Request("txtDateTo"))
 else
    v_txtDateTo = textDate(date)
 end if
 
 if Trim(Request("txtDateFrom2")) <> "" then
    v_txtDateFrom2 = Trim(Request("txtDateFrom2"))
 else
   v_txtDateFrom2 = textDate(date)
 end if
if Trim(Request("txtDateTo2")) <> "" then
    v_txtDateTo2 = Trim(Request("txtDateTo2"))
 else
    v_txtDateTo2 = textDate(date)
 end if
 

v_txtGroupPlantation 			  = Request("txtGroupPlantation")
v_txtPlantation 			      = Request("txtPlantation")
v_txtGroupRegion 			      = Request("txtGroupRegion")
v_txtRegion   	 			      = Request("txtRegion")
v_txtCompany 			          = Request("txtCompany")
v_txtFcba 			              = Request("txtFcba")
v_txtDateFrom 			          = Request("txtDateFrom")


 
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
Dim rsTerimaTBS, strTerimaTBS, nSpanTotal
strTerimaTBS ="select Regioncode,CompanyCode,fcba,fieldcode,LongInterval,Afdeling,Harvestingsection,hectarageplanted " & _
              " from ( select businessunit_api.get_Groupplantation(fcba) Groupplantation, businessunit_api.get_plantation(fcba) " & _
			  "         PlantationCode, businessunit_api.get_Groupregion(fcba)     GroupRegionCode, businessunit_api.get_region(fcba) " & _
			  "         Regioncode, REGION_API.get_OrderData(businessunit_api.get_region(fcba)) orderregion, businessunit_api.get_CompanyCode(fcba) CompanyCode,company_api.get_OrderData(businessunit_api.get_CompanyCode(fcba)) ordercompany ,h.fcba , BUSINESSUNIT_API.get_OrderData(fcba) orderfcba, h.fieldcode, h.intervalnumber " & _
			  "         LongInterval, field_api.getValid_Afdeling(h.fieldcode,h.fcba,ondate) Afdeling," & _
              "			field_api.GetValid_Harvestingsection(h.fieldcode,h.fcba,ondate) Harvestingsection, " & _
			  "         field_api.GetValid_hectarageplanted(h.fieldcode,h.fcba,ondate)  hectarageplanted,  " & _ 
			  "         field_api.GetValid_Activation(h.fieldcode,h.fcba,ondate)  Activation,  " & _ 
			  "         field_api.GetValid_status(h.fieldcode,h.fcba,ondate)  status  " & _ 
			  "         from harvesting_interval_tab h where h.intervalnumber  > 10 and ondate = to_date('" & v_txtDateFrom & "', 'mm/dd/yyyy')  " & _
			  "      ) where  LongInterval = HARVESTING_INTERVAL_API.GetMaxPusinganPerDay(fcba,to_date('" & v_txtDateFrom & "', 'mm/dd/yyyy') ) " & _
			  "        AND status='TM'  and Activation='Y'" 

wherecons = " and companycode is not null "
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
			end  if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  PlantationCode = '" & v_txtPlantation & "' "
				end if   
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
		        
			     
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegionCode = '" & v_txtGroupRegion & "' "
				end if
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Regioncode = '" & v_txtRegion & "' "
				end if
				
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  companycode = '" & v_txtCompany & "' "
				end if
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
			else
			    nSpanTotal = nSpanTotal - 1 
			end if 
					
'response.write strTerimaTBS & whereCons
Set rsTerimaTBS = Server.CreateObject("ADODB.Recordset")
rsTerimaTBS.Open strTerimaTBS & whereCons & " order by orderregion,ordercompany ,orderfcba , Afdeling" , connstr, adOpenStatic


%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Pusingan Terpanjang</title>
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
</head>

<body>
	<div id="wrapper">
	<!--#include file="menutop.asp"-->
	<!--#include file="../print_excel.asp"-->
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<div id="header-wrapper" style=" ">
	<!--#include virtual="/REPORT/FormSelectNewSumE.html"-->
    <table id="form-table" class="centered-margin">

	<tr>		
		<td><font face="verdana" size="1">Date</font></td>
		<td align="center">
		<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
	    </td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</table>
	</div>
	</form>

<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT.<%=CompanyName%></h4>
		<h5><%=fcbaname%></h5>
	    </div>
		<div id="subheader" class="centered"  style="width:50%; margin:auto;">
		<h2>SUMMARY PUSINGAN PANEN TERPANJANG</h3>
		<br/>
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdatefrom%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%></h5>
		<% end if %>
		</div>
	</div>
 	</th>
   </tr><br>
	<tr>
	  <th rowspan="2" style="color:#FFFFFF" width='5%'>NO</th>
	  <th rowspan="2" style="color:#FFFFFF" width='10%'>Group Kebun</th>
	  <th rowspan="2" style="color:#FFFFFF" width='10%'>PT</th>
	  <th rowspan="2" style="color:#FFFFFF" width='10%'>Kebun</th>
	  <th rowspan="2" style="color:#FFFFFF" width='15%'>Pusingan Terpanjang<br> Hari Ini</th>
	  <th colspan="4" style="color:#FFFFFF" width='45%'>Lokasi</th>
    </tr>
	<tr>
	    <th style="color:#FFFFFF" width='10%'>Blok</th>
		<th style="color:#FFFFFF" width='10%'>Afdeling</th>
		<th style="color:#FFFFFF" width='10%'>Seksi Panen</th>
		<th style="color:#FFFFFF" width='10%'>Luas</th>
	</tr>
    </thead>
	<tbody>
<%
Dim tempfcba,tempCompany,tempregion,tempTotHafcba,tempTotHaCompany,templongintervalregion
Dim WEfcba,WECompany, WEregion
DIM avgfcba,avgCompany

tempfcba    = ""
tempCompany = ""
tempregion  = ""
WEfcba      = 0
WECompany   = 0 
WEregion    = 0
avgfcba     = 0
tempTotHafcba = 0
tempTotHaCompany = 0
templongintervalregion =0

if rsTerimaTBS.bof and rsTerimaTBS.eof then
%>
	<tr><td colspan="9">NO DATA AVAILABLE</td></tr>
<%else
		NO = 1
		Do while not rsTerimaTBS.eof
			if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			
			If tempCompany = "" then
			   tempTotHaCompany = cdbl(rsTerimaTBS("HECTARAGEPLANTED"))
			end if
			
			If tempfcba <> rsTerimaTBS("FCBA") and tempfcba <> "" then
			   	avgfcba= cdbl(WEfcba)/cdbl(tempTotHafcba)
				
			%>
			    <tr class="<%=cClass%>">
					<td align="center"  colspan="4"> Sub Total <%=tempfcba %></td>
					<td align="right"><%=FormatNumber(avgfcba,2)%></td>
					<td align="center"  colspan="3">  </td>
					<td align="right"><%=FormatNumber(tempTotHafcba,2)%></td>
				</tr>
         	<%		
			    
				
				If tempCompany = rsTerimaTBS("COMPANYCODE")  then
				   tempTotHaCompany = cdbl(tempTotHafcba) + cdbl(rsTerimaTBS("HECTARAGEPLANTED"))
				   WECompany = cdbl(WECompany) + cdbl(rsTerimaTBS("LONGINTERVAL")) * cdbl(rsTerimaTBS("HECTARAGEPLANTED")) 		
				end if
				
				WEfcba = 0
				tempTotHafcba = cdbl(rsTerimaTBS("HECTARAGEPLANTED"))
				
				WEfcba  =  cdbl(WEfcba) + cdbl(rsTerimaTBS("LONGINTERVAL")) * cdbl(rsTerimaTBS("HECTARAGEPLANTED")) 
				
		    else
                tempTotHafcba = cdbl(tempTotHafcba) + cdbl(rsTerimaTBS("HECTARAGEPLANTED"))
				WEfcba  =  cdbl(WEfcba) + cdbl(rsTerimaTBS("LONGINTERVAL")) * cdbl(rsTerimaTBS("HECTARAGEPLANTED")) 			
			End if
			
			If tempCompany <> rsTerimaTBS("COMPANYCODE") and tempCompany <> "" then
			   				
			End if
			
			
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=NO%></td>
		<td align="center"><%=rsTerimaTBS("REGIONCODE")%></td>
		<td align="center"><%=rsTerimaTBS("COMPANYCODE")%></td>
		<td align="center"><%=rsTerimaTBS("FCBA")%></td>
		<td align="right"><%=FormatNumber(rsTerimaTBS("LONGINTERVAL"),2)%></td>
		<td align="center"><%=rsTerimaTBS("FIELDCODE")%></td>
		<td align="center"><%=rsTerimaTBS("AFDELING")%></td>
		<td align="center"><%=rsTerimaTBS("HARVESTINGSECTION")%></td>
		<td align="right"><%=FbpNum2(rsTerimaTBS("HECTARAGEPLANTED"))%></td>
	</tr>
<%
            NO = NO + 1
			'GrandTOTAL(1) =  GrandTOTAL(1) + cdbl(rsTerimaTBS("bruto"))
			'GrandTOTAL(2) =  GrandTOTAL(2) + cdbl(rsTerimaTBS("Sortase"))
			'GrandTOTAL(3) =  GrandTOTAL(3) + cdbl(rsTerimaTBS("bruto")) - cdbl(rsTerimaTBS("sortase"))
			response.write  tempTotHaCompany
			tempfcba    = rsTerimaTBS("FCBA")
			tempCompany = rsTerimaTBS("COMPANYCODE")
			tempregion  = rsTerimaTBS("COMPANYCODE")
			
			rsTerimaTBS.MoveNext
					
		loop
end if 

rsTerimaTBS.Close
set rsTerimaTBS = Nothing

  If cdbl(tempTotHafcba) > 0 then

	avgfcba= cdbl(WEfcba)/cdbl(tempTotHafcba)
	avgCompany= cdbl(WEcompany)/cdbl(tempTotHaCompany)

%>
       <tr class="<%=cClass%>">
		   <td align="center"  colspan="4"> Sub Total <%=tempfcba %></td>
    	   <td align="right"><%=FormatNumber(avgfcba,2)%></td>
    	   <td align="center"  colspan="3">  </td>
		   <td align="right"><%=FormatNumber(tempTotHafcba,2)%></td>
		</tr>   
       <tr class="<%=cClass%>">
			<td align="center"  colspan="4"> Sub Total <%=tempCompany %></td>
			<td align="right"><%=FormatNumber(avgCompany,2)%></td>
			<td align="center"  colspan="3">  </td>
			<td align="right"><%=FormatNumber(tempTotHaCompany,2)%></td>
		</tr>
<% End If %>		
       </tbody>
</table>
</body>
   </html>
