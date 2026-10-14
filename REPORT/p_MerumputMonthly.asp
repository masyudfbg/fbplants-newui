<%@Language=VBScript%>
<% Option Explicit %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsMerumputTM, strMerumputTM, strGroup, strOrder, strQuery
	Dim  MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth,v_txtdatefrom, v_txtdateto, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass, v_txtStatusTmTbm, v_txtjobcode, v_txtGroup
	Dim nSpanTotal
	Dim txtdatefrom, txtdateto, txtDivision, txtBlok, txtPlantingDate, txtStatusTmTbm, txtjobcode, txtGroup
	Dim sprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode, sprevStatus, sprevJobCode, sprevGroup
    Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
	Dim tempDivision, tempPlantingDate,  tempBlok,  tempStatus,  tempJobcode,  tempGroup 
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	Dim ExtendedQuery, supportExtendedQuery, strExtendedQuery,strEPS2 
    'Dim rscomp , strcomp,txtcompany ,txcompanyname,txtComp,v_txtComp
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
		
	 nSpanTotal = 10
	
	v_txtGroupPlantation   = Request("txtGroupPlantation")
	v_txtPlantation  	   = Request("txtPlantation")
	v_txtGroupRegion  	   = Request("txtGroupRegion")
	v_txtRegion  		   = Request("txtRegion")
	v_txtcompany  		   = Request("txtCompany")
	v_txtfcba 			   = Request("txtfcba")
	v_txtDivision		   = Request("txtDivision")
	v_txtBlok			   = Request("txtBlok")
	v_txtPlantingDate	   = Request("txtPlantingDate")
	v_txtdatefrom		   = Request("txtdatefrom")
	v_txtdateto			   = Request("txtdateto")
	v_txtStatusTmTbm       = Request("txtStatusTmTbm")
	v_txtjobcode     	   = Request("txtjobcode")
	v_txtGroup			   = Request("txtGroup")
	'response.write v_txtGroup	
	if v_txtdateto ="" then
	   v_txtdatefrom = textdate(Now())
	   v_txtdateto = textdate(Now())
    end if	
	
	        v_txtyear 		= Request("txtyear")
			v_txtmonth 		= lpad(Request("txtmonth"),"0",2)

	        'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " TT" & v_txtPlantingDate
			'response.write  "TM/TBM " & v_txtStatusTmTbm & "  jobcode " & v_txtjobcode & "  Group " &  v_txtGroup

					
			Dim selectFields,strFieldCons,selectCons,whereCons,orderCons 
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, BUSINESSUNIT_API.get_OrderData(fcba) norut, fcba, " 

						   
			wherecons = " where company is not null "
			if v_txtGroup <> "--SUMMARY--" And v_txtGroup <> "--ALL--" And v_txtGroup <> ""  then
			   whereCons = wherecons & " And param_2= '" & v_txtGroup & "' "  
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
	        
			if  v_txtStatusTmTbm <> "--SUMMARY--" AND v_txtStatusTmTbm <> "" then
				if v_txtStatusTmTbm <> "--ALL--" then
					whereCons = whereCons &  " AND STATUS = '" &  v_txtStatusTmTbm & "'"
				end if 
                selectCons = selectCons &  " , STATUS "	
				orderCons  = orderCons & ",STATUS"
			else
			   nSpanTotal = nSpanTotal - 1 
			end if	
			if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",fcba"
			    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba) "
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
			if  v_txtDivision <> "--SUMMARY--" AND v_txtDivision <> "" then
			      if v_txtDivision <>  "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      end if
                  selectCons = selectCons &  " , DIVISION "
				  orderCons  = orderCons & ",DIVISION "
			else
			   nSpanTotal = nSpanTotal - 1 
			end if
            
			if v_txtPlantingDate <> "--SUMMARY--"  AND  v_txtPlantingDate <> "" then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   end if
				   selectCons = selectCons &  " , PLANTINGYEAR "
				   orderCons  = orderCons & ",PLANTINGYEAR "
			else
			   nSpanTotal = nSpanTotal - 1 
			end if
			
			if  v_txtBlok <> "--SUMMARY--" AND v_txtBlok <> "" then
				   if v_txtBlok <>  "--ALL--" then
					  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
				   end if
				   selectCons = selectCons & " , FIELDCODE "
                   orderCons  = orderCons & ",FIELDCODE "				   
			else
			   nSpanTotal = nSpanTotal - 1 
            end if
			
	        
			if v_txtjobcode <> "--SUMMARY--" AND v_txtjobcode <> "" then
				   if v_txtjobcode <> "--ALL--" then
					  whereCons = whereCons &  " AND  jobcode = '" & v_txtjobcode & "'"
				   end if
			      selectCons = selectCons &  " , jobcode"				   
			else
			   nSpanTotal = nSpanTotal - 1 
			end if	
	
    'strMerumputTM = " SELECT c.param_2, a.status "
    ' strMerumputTM = " SELECT c.param_2, c.param_2_name "
	
	' supportExtendedQuery = 0

	

	strMerumputTM 		=     " select * from MERUMPUT_MONTHLY_REPORT_TAB t " & _ 
                              " where t.year = '" & v_txtyear &"' and t.month = '"& v_txtmonth &"' and USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) ='TRUE'  " 
							  
	
                      strEPS2    =  " ,Sum(total_01) total_01,sum(total_bo_01) total_bo_01,decode(sum(total_bo_01),0,0,Sum(total_01)/sum(total_bo_01)) hk_ha_01,sum(Amount_01) Amount_01,decode(sum(total_bo_01),0,0,sum(Amount_01)/sum(total_bo_01)) Amount_ha_01 " & _
									" ,Sum(total_02) total_02,sum(total_bo_02) total_bo_02,decode(sum(total_bo_02),0,0,Sum(total_02)/sum(total_bo_02)) hk_ha_02,sum(Amount_02) Amount_02,decode(sum(total_bo_02),0,0,sum(Amount_02)/sum(total_bo_02)) Amount_ha_02 " & _
                                    " ,Sum(total_03) total_03,sum(total_bo_03) total_bo_03,decode(sum(total_bo_03),0,0,Sum(total_03)/sum(total_bo_03)) hk_ha_03,sum(Amount_03) Amount_03,decode(sum(total_bo_03),0,0,sum(Amount_03)/sum(total_bo_03)) Amount_ha_03 " & _
                                    " ,Sum(total_04) total_04,sum(total_bo_04) total_bo_04,decode(sum(total_bo_04),0,0,Sum(total_04)/sum(total_bo_04)) hk_ha_04,sum(Amount_04) Amount_04,decode(sum(total_bo_04),0,0,sum(Amount_04)/sum(total_bo_04)) Amount_ha_04 " & _
                                    " ,Sum(total_05) total_05,sum(total_bo_05) total_bo_05,decode(sum(total_bo_05),0,0,Sum(total_05)/sum(total_bo_05)) hk_ha_05,sum(Amount_05) Amount_05,decode(sum(total_bo_05),0,0,sum(Amount_05)/sum(total_bo_05)) Amount_ha_05 " & _
                                    " ,Sum(total_06) total_06,sum(total_bo_06) total_bo_06,decode(sum(total_bo_06),0,0,Sum(total_06)/sum(total_bo_06)) hk_ha_06,sum(Amount_06) Amount_06,decode(sum(total_bo_06),0,0,sum(Amount_06)/sum(total_bo_06)) Amount_ha_06 " & _
                                    " ,Sum(total_07) total_07,sum(total_bo_07) total_bo_07,decode(sum(total_bo_07),0,0,Sum(total_07)/sum(total_bo_07)) hk_ha_07,sum(Amount_07) Amount_07,decode(sum(total_bo_07),0,0,sum(Amount_07)/sum(total_bo_07)) Amount_ha_07 " & _
                                    " ,Sum(total_08) total_08,sum(total_bo_08) total_bo_08,decode(sum(total_bo_08),0,0,Sum(total_08)/sum(total_bo_08)) hk_ha_08,sum(Amount_08) Amount_08,decode(sum(total_bo_08),0,0,sum(Amount_08)/sum(total_bo_08)) Amount_ha_08 " & _
                                    " ,Sum(total_09) total_09,sum(total_bo_09) total_bo_09,decode(sum(total_bo_09),0,0,Sum(total_09)/sum(total_bo_09)) hk_ha_09,sum(Amount_09) Amount_09,decode(sum(total_bo_09),0,0,sum(Amount_09)/sum(total_bo_09)) Amount_ha_09 " & _
                                    " ,Sum(total_10) total_10,sum(total_bo_10) total_bo_10,decode(sum(total_bo_10),0,0,Sum(total_10)/sum(total_bo_10)) hk_ha_10,sum(Amount_10) Amount_10,decode(sum(total_bo_10),0,0,sum(Amount_10)/sum(total_bo_10)) Amount_ha_10 " & _
                                    " ,Sum(total_11) total_11,sum(total_bo_11) total_bo_11,decode(sum(total_bo_11),0,0,Sum(total_11)/sum(total_bo_11)) hk_ha_11,sum(Amount_11) Amount_11,decode(sum(total_bo_11),0,0,sum(Amount_11)/sum(total_bo_11)) Amount_ha_11 " & _
                                    " ,Sum(total_12) total_12,sum(total_bo_12) total_bo_12,decode(sum(total_bo_12),0,0,Sum(total_12)/sum(total_bo_12)) hk_ha_12,sum(Amount_12) Amount_12,decode(sum(total_bo_12),0,0,sum(Amount_12)/sum(total_bo_12)) Amount_ha_12 " & _
                                    " ,Sum(total_SBI) total_SBI,sum(total_bo_SBI) total_bo_SBI,decode(sum(total_bo_SBI),0,0,Sum(total_SBI)/sum(total_bo_SBI)) hk_ha_SBI,sum(Amount_SBI) Amount_SBI,decode(sum(total_bo_SBI),0,0,sum(Amount_SBI)/sum(total_bo_SBI)) Amount_ha_SBI " & _
                                    " ,Sum(bud_HK_SBI) bud_HK_SBI,sum(bud_HA_SBI) bud_HA_SBI,decode(sum(bud_HA_SBI),0,0,Sum(bud_HK_SBI)/sum(bud_HA_SBI)) bud_hk_ha_SBI,sum(bud_RP_SBI) bud_Amount_SBI,decode(sum(bud_HA_SBI),0,0,sum(bud_RP_SBI)/sum(bud_HA_SBI)) bud_Amount_ha_SBI " 

									 
	strMerumputTM =  "select param_2,param_2_name,GroupPlantation " & selectCons & strEPS2 & "  From (" & strMerumputTM & " )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by  param_2,param_2_name,GroupPlantation" & selectCons & "  Order by param_2,param_2_name,GroupPlantation" & orderCons  
	
	
	
	
	
	Set rsMerumputTM = Server.CreateObject("ADODB.Recordset")
	' response.write  strMerumputTM
		
	rsMerumputTM.Open strMerumputTM, connstr, adOpenStatic 	
	


	
%>


<html>

<head>
<title>ePlaS Reporting - Laporan Bulanan Merumput</title>
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

<!--<script type="text/javascript" src="http://jqueryjs.googlecode.com/files/jquery-1.3.1.min.js" > </script> -->
<script type="text/javascript">

    function PrintElem(elem)
    {
        Popup($(elem).html());
    }

    function Popup(data) 
    {
        var mywindow = window.open('', 'my div', 'height=400,width=600');
        mywindow.document.write('<html><head><title>my div</title>');
        mywindow.document.write('<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />');
        mywindow.document.write('</head><body >');
        mywindow.document.write(data);
        mywindow.document.write('</body></html>');

        mywindow.print();
        mywindow.close();

        return true;
    }

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
<div class="top_select" style=" position:absolute;  top:0; left:0px; width:100%;"> 
<form name="frmEntry" method="post" action="p_MerumputMonthly.asp">

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
		<th>STATUS</th>
		<th>PEKERJAAN</th>
		<th>AFDELING</th>
		<th>TAHUN TANAM</th>
		<th>BLOK</th>
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
		<select name="txtGroup">	
			<option value="">--ALL--</option>
		<%
			Dim rsGroup
			Set rsGroup = Server.CreateObject("ADODB.Recordset")
			rsGroup.CursorLocation = adUseClient
			
			
			if len(trim(v_txtStatusTmTbm)) > 0 then 
			    strsql =  "  select distinct fccode, fcname from parameterdetail " & _
							" where " & _
							" parcode = 'SR002' and " & _
							" parheadcode = 'GROUPJOB' " & _
							" and fcba = '"& v_txtFcba &"'" 
							
			else
			    strsql =  "  select distinct fccode, fcname from parameterdetail " & _
							" where " & _
							" parcode = 'SR002' and " & _
							" parheadcode = 'GROUPJOB' " & _
							" and fcba = '"& Session("fcba") &"'" 
							
			end if
			
			
			
			rsGroup.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsGroup.eof then
				DO while not rsGroup.eof
					if v_txtGroup = rsGroup("fccode") then
			%>			<option value="<%=rsGroup("fccode")%>" Selected><%=rsGroup("fccode")%> - <%= rsGroup("fcname") %></option>
					<%else%>
						<option value="<%=rsGroup("fccode")%>"><%=rsGroup("fccode")%> - <%= rsGroup("fcname") %></option>
					<%end if%>
		<%			rsGroup.movenext
				loop
			End if
			rsGroup.Close
			set rsGroup = nothing
		%>
		</select>		
		</td>
		
		
		<td align="center">
		<select name="txtStatusTmTbm">	
			<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsStatusTmTbm
			Set rsStatusTmTbm = Server.CreateObject("ADODB.Recordset")
			rsStatusTmTbm.CursorLocation = adUseClient
			
			
			' if v_txtStatusTmTbm <> "--SUMMARY--" AND v_txtStatusTmTbm <> "" then 
				' strsql =  "  SELECT '--ALL--' as STATUSTMTBM from DUAL UNION SELECT FCCODE as STATUSTMTBM from PARAMETER where parcode = 'F0006' and FCBA = '" & v_txtFcba & "'"
			' else
				strsql =  "  SELECT '--ALL--' as STATUSTMTBM from DUAL UNION SELECT DISTINCT FCCODE as STATUSTMTBM from PARAMETER where parcode = 'F0006' " 

			' end if
			
			
			rsStatusTmTbm.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsStatusTmTbm.eof then
				DO while not rsStatusTmTbm.eof
				if v_txtStatusTmTbm = rsStatusTmTbm("STATUSTMTBM") then
		%>			<option value="<%=rsStatusTmTbm("STATUSTMTBM")%>" Selected><%=rsStatusTmTbm("STATUSTMTBM")%></option>
				<%else%>
					<option value="<%=rsStatusTmTbm("STATUSTMTBM")%>"><%=rsStatusTmTbm("STATUSTMTBM")%></option>
				<%end if%>
		<%			rsStatusTmTbm.movenext
				loop
			End if
			rsStatusTmTbm.Close
			set rsStatusTmTbm = nothing
		%>
		</select>		
		</td>		
		
		
		<td>
				<select name="txtjobcode">	
				<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsJob
			Set rsJob = Server.CreateObject("ADODB.Recordset")
			rsJob.CursorLocation = adUseClient
			
			' Get the Activity from ParameterDetail Table
			
			if len(trim(v_txtBlok)) > 0 then 
			    strsql = " SELECT '--ALL--' as fccode, '' as fcname  from DUAL UNION SELECT FCCODE, FCNAME FROM Job WHERE FCCODE IN ( " & _
						 " select param_1 from parameterdetail " & _
						 " where " & _
						 " parcode = 'SR002' and " & _
						 " parheadcode in ( 'ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 " ) ORDER BY FCCODE ASC " 
			
			else
			    strsql = " SELECT '--ALL--' as fccode, '' as fcname  from DUAL UNION SELECT DISTINCT FCCODE, FCNAME FROM Job WHERE FCCODE IN ( " & _
						 " select param_1 from parameterdetail " & _
						 " where " & _
						 " parcode = 'SR002' and " & _
						 " parheadcode in ( 'ACTIVITY_TBM', 'ACTIVITY_TM') " & _
						 " )ORDER BY FCCODE ASC " 

			end if		
			
			rsJob.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsJob.eof then
				DO while not rsJob.eof
				if v_txtjobcode = rsJob("FCCODE") then
		%>			<option value="<%=rsJob("FCCODE")%>" Selected><%=rsJob("FCCODE")%> - <%=rsJob("FCNAME")%></option>
				<%else%>
					<option value="<%=rsJob("FCCODE")%>"><%=rsJob("FCCODE")%> - <%=rsJob("FCNAME")%></option>
				<%end if%>
		<%			rsJob.movenext
				loop
			End if
			rsJob.Close
			set rsJob = nothing
		%>
		</select>		
		</td>
		
		
		<td align="center">
		<select name="txtDivision">	
			<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsDivision
			Set rsDivision = Server.CreateObject("ADODB.Recordset")
			rsDivision.CursorLocation = adUseClient
			

		
			
			if len(trim(txtDivision)) > 0 then 
				strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
			else
				strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD "
			end if
			
			rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsDivision.eof then
				DO while not rsDivision.eof
				if v_txtDivision = rsDivision("Division") then
		%>			<option value="<%=rsDivision("Division")%>" Selected><%=rsDivision("Division")%></option>
				<%else%>
					<option value="<%=rsDivision("Division")%>"><%=rsDivision("Division")%></option>
				<%end if%>
		<%			rsDivision.movenext
				loop
			End if
			rsDivision.Close
			set rsDivision = nothing
		%>
		</select>		
		</td>

		
		
		<td align="center">
		<select name="txtPlantingDate">	
				<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsPlantingDate
			Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
			rsPlantingDate.CursorLocation = adUseClient
			

			if len(trim(txtPlantingDate)) > 0 then 
				strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
			else
				strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD "
			end if

			
			
			rsPlantingDate.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsPlantingDate.eof then
				DO while not rsPlantingDate.eof
				if v_txtPlantingDate = rsPlantingDate("PlantingDate") then
		%>			<option value="<%=rsPlantingDate("PlantingDate")%>" Selected><%=rsPlantingDate("PlantingDate")%></option>
				<%else%>
					<option value="<%=rsPlantingDate("PlantingDate")%>"><%=rsPlantingDate("PlantingDate")%></option>
				<%end if%>
		<%			rsPlantingDate.movenext
				loop
			End if
			rsPlantingDate.Close
			set rsPlantingDate = nothing
		%>
		</select>		
		</td>
		

		
		<td align="center">
		<select name="txtBlok">	
				<option value="--SUMMARY--">--SUMMARY--</option>
		<%
			Dim rsBlok
			Set rsBlok = Server.CreateObject("ADODB.Recordset")
			rsBlok.CursorLocation = adUseClient
			
			if len(trim(v_txtBlok)) > 0 then 
				strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
			else
				strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
			end if			
			
			rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsBlok.eof then
				DO while not rsBlok.eof
				if v_txtBlok = rsBlok("Blok") then
		%>			<option value="<%=rsBlok("Blok")%>" Selected><%=rsBlok("Blok")%></option>
				<%else%>
					<option value="<%=rsBlok("Blok")%>"><%=rsBlok("Blok")%></option>
				<%end if%>
		<%			rsBlok.movenext
				loop
			End if
			rsBlok.Close
			set rsBlok = nothing
		%>
		</select>		
		</td>
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
    <tr>
<td colspan="140" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
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
						<h2>LAPORAN BULANAN PERAWATAN</h2>
						<h5>PERIODE : <%=v_txtmonth%> - <%=v_txtyear%> </h5>
					</div>
				</div>

</td>
</tr>
</table>
</div>

<div id="mydiv" style="margin-top:280px;">
<table border="0" width="100%" id="product-table" class="xthf-site" style="background-color:#FFFFFF;">
<thead>
	<tr>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">GROUP</th>
		<% if v_txtStatusTmTbm <> "--SUMMARY--" AND v_txtStatusTmTbm <> "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" >STATUS</th>
		<% end if %>

		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" ><%=titlePlantation%></th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" ><%=titleGroupRegion%></th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;" ><%=titleRegion%></th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="3" style="background-color:#00A300;color:#fff;" ><%=titleCompany%></th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="3" style="background-color:#00A300;color:#fff;"  ><%=titleSite%></th>
		<% end if %>

		
		<% if v_txtDivision <> "--SUMMARY--" AND v_txtDivision <> "" then %>
		    <th rowspan="3" style="background-color:#00A300;color:#fff;"  >AFDELING</th>
		<% end if %>
		
		
		<% if v_txtPlantingDate <> "--SUMMARY--" AND v_txtPlantingDate <> "" then %>
		    <th rowspan="3" style="background-color:#00A300;color:#fff;"  >TT</th>
		<% end if %>


		<% if v_txtBlok <> "--SUMMARY--" AND v_txtBlok <> "" then %>
		    <th rowspan="3" style="background-color:#00A300;color:#fff;"  >BLOK</th>
		<% end if %>
		
	
		
		<th colspan="5" style="background-color:#00A300;color:#fff;" >JAN</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >FEB</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >MAR</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >APR</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >MAY</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >JUN</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >JUL</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >AUG</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >SEP</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >OCT</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >NOV</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >DES</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >SBI</th>
		<th colspan="5" style="background-color:#00A300;color:#fff;" >BGT</th>
	</tr>
	<tr>
	   <% DIM K
		   FOR K = 1 TO 14 %>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >HK</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >HA</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >HK/HA</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >RP</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;" >RP/HA</th>
		<% NEXT %>	
	</tr>
	
</thead>
<tbody>
<%
Dim nsubsku, nsubkhl, nsubtotal, nsubha_harian, nsubbo_harian,nbsubbudHK,nbsubbudHA,nbsubbudRasio,nbsubbudGLI,nbsubbudPAR,nbsubbudMET,nbsubbudFLO
Dim nsubtotal_bo_harian, nsubha_per_hk, nsubgli_ltr, nsubgli_ltr_ha, nsubpar_ltr
Dim nsubpar_ltr_ha, nsubmet_kg, nsubmet_kg_ha, nsubflo_kg, nsubflo_kg_ha, nsubamount_hi_per_ha
Dim ntotal_01,ntotal_bo_01,nhk_ha_01,nAmount_01,nAmount_ha_01,ntotal_02,ntotal_bo_02,nhk_ha_02,nAmount_02,nAmount_ha_02,ntotal_03,ntotal_bo_03,nhk_ha_03,nAmount_03,nAmount_ha_03
Dim	ntotal_04,ntotal_bo_04,nhk_ha_04,nAmount_04,nAmount_ha_04,ntotal_05,ntotal_bo_05,nhk_ha_05,nAmount_05,nAmount_ha_05,ntotal_06,ntotal_bo_06,nhk_ha_06,nAmount_06,nAmount_ha_06
Dim	ntotal_07,ntotal_bo_07,nhk_ha_07,nAmount_07,nAmount_ha_07,ntotal_08,ntotal_bo_08,nhk_ha_08,nAmount_08,nAmount_ha_08,ntotal_09,ntotal_bo_09,nhk_ha_09,nAmount_09,nAmount_ha_09
Dim	ntotal_10,ntotal_bo_10,nhk_ha_10,nAmount_10,nAmount_ha_10,ntotal_11,ntotal_bo_11,nhk_ha_11,nAmount_11,nAmount_ha_11,ntotal_12,ntotal_bo_12,nhk_ha_12,nAmount_12,nAmount_ha_12
Dim	ntotal_sbi,ntotal_bo_sbi,nhk_ha_sbi,nAmount_sbi,nAmount_ha_sbi,ntotal_bud,ntotal_bo_bud,nhk_ha_bud,nAmount_bud,nAmount_ha_bud

	if not rsMerumputTM.EOF then
		NO = 1
		nsubsku = 0
		nsubkhl = 0
		nsubtotal = 0
		nsubha_harian = 0
		nsubbo_harian = 0
		nsubtotal_bo_harian = 0
		nsubha_per_hk = 0
		nsubgli_ltr = 0
		nsubgli_ltr_ha = 0
		nsubpar_ltr = 0
		nsubpar_ltr_ha = 0
		nsubmet_kg = 0
		nsubmet_kg_ha = 0
		nsubflo_kg = 0
		nsubflo_kg_ha = 0
		
		nsubamount_hi_per_ha = 0

		while not rsMerumputTM.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				
				
				<% if v_txtGroup <> "--SUMMARY--"   then %>	
				        <% tempGroup = rsMerumputTM("param_2") 
						   tempStatus = "--ALL--" 
						   tempJobcode = "--SUMMARY--" 
						   tempPlantation = "--SUMMARY--" 
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempDivision = "--SUMMARY--"
						   tempPlantingDate = "--SUMMARY--"
						   tempBlok = "--SUMMARY--"
						  %> 
				
						<% if sprevGroup <> rsMerumputTM("param_2") then %>
							<td align="left" >
								<a href="p_MerumputMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
								<%=rsMerumputTM("param_2_name")%>
								</a>
							</td>
							<% sprevGroup = rsMerumputTM("param_2") %>	
						<% else	%>
							<td align="left" ></td>					
						<% end if  
				     end if 
				  %>	
				
				
				<% if v_txtStatusTmTbm <> "--SUMMARY--" AND v_txtStatusTmTbm <> "" then %>	
                     <%    tempGroup = rsMerumputTM("param_2") 
						   tempStatus = rsMerumputTM("status")
						   tempJobcode = "--SUMMARY--" 
						   tempPlantation = "--ALL--" 
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempDivision = "--SUMMARY--"
						   tempPlantingDate = "--SUMMARY--"
						   tempBlok = "--SUMMARY--"
						  %> 				
					<% if sprevStatus <> rsMerumputTM("status") then %>

						<td align="center">
							<a href="p_MerumputMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("status")%> 
							</a>
						</td>
						<% sprevStatus = rsMerumputTM("status") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>	
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     tempGroup = rsMerumputTM("param_2") 
					       'response.write v_txtStatusTmTbm 
					       if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtStatusTmTbm <>  "" And v_txtStatusTmTbm  <>  "--ALL--" then 
						      tempStatus = rsMerumputTM("status")
						   else
						      tempStatus =""
						   end if
						   tempJobcode = "--SUMMARY--" 
						   tempPlantation = rsMerumputTM("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempDivision = "--SUMMARY--"
						   tempPlantingDate = "--SUMMARY--"
						   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsMerumputTM("Plantation") then %>
						<td align="center">
							<a href="p_MerumputMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsMerumputTM("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  tempGroup = rsMerumputTM("param_2") 
						   if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtStatusTmTbm <>  "" then 
						      tempStatus = rsMerumputTM("status")
						   else
						      tempStatus =""
						   end if
						tempJobcode = "--SUMMARY--" 
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsMerumputTM("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsMerumputTM("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsMerumputTM("GroupRegion") then %>
						<td align="center">
							<a href="p_MerumputMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsMerumputTM("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <%  tempGroup = rsMerumputTM("param_2") 
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtStatusTmTbm <>  "" then 
						      tempStatus = rsMerumputTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--" 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsMerumputTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsMerumputTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsMerumputTM("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsMerumputTM("Region") then %>
						<td align="center">
							<a href="p_MerumputMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("region")%>
							</a>
						</td>
						<% SprevRegion = rsMerumputTM("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					    tempGroup = rsMerumputTM("param_2") 
					    
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtStatusTmTbm <>  "" then 
						      tempStatus = rsMerumputTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--" 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsMerumputTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsMerumputTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsMerumputTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsMerumputTM("Company")
					   tempfcba = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsMerumputTM("Company") then %>
						<td align="center">
							<a href="p_MerumputMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("company")%>
							</a>
						</td>
						<% SprevCompany = rsMerumputTM("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <%  tempGroup = rsMerumputTM("param_2") 
					    
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtStatusTmTbm <>  "" then 
						      tempStatus = rsMerumputTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--" 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsMerumputTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsMerumputTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsMerumputTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsMerumputTM("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsMerumputTM("fcba")
					   tempDivision = "--ALL--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsMerumputTM("fcba") then %>
						<td align="center">
							<a href="p_MerumputMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
						    <%=rsMerumputTM("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsMerumputTM("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> "" then %>
				   <%   tempGroup = rsMerumputTM("param_2") 
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtStatusTmTbm <>  "" then 
						      tempStatus = rsMerumputTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--"  
				       if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsMerumputTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsMerumputTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsMerumputTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsMerumputTM("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsMerumputTM("fcba")
					   else
					       tempfcba = ""
					   end if
					   if v_txtdivision <>  "--SUMMARY--" and v_txtdivision  <>  "" then 
					      tempDivision = rsMerumputTM("division")
					   else
					      tempDivision = ""
                       end if					   
					   tempPlantingDate = "--ALL--"
					   tempBlok = "--SUMMARY--"					   
					 %>
				
				
				    <% if sprevDivision <> rsMerumputTM("division") then %>
					    <td align="center">
							<a href="p_MerumputMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("division")%>
							</a>
						</td>
						<% sprevDivision = rsMerumputTM("division") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>

				<% if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> "" then %>
				   <%   tempGroup = rsMerumputTM("param_2") 
						if v_txtStatusTmTbm  <>  "--SUMMARY--"  and  v_txtStatusTmTbm <>  "" then 
						      tempStatus = rsMerumputTM("status")
						else
						      tempStatus =""
						end if
						tempJobcode = "--SUMMARY--" 
				       if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsMerumputTM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsMerumputTM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsMerumputTM("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsMerumputTM("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsMerumputTM("fcba")
					   else
					       tempfcba = ""
					   end if
					   if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> "" then
					      tempDivision = rsMerumputTM("division")
						else
    						tempDivision = ""
					   end if 		
					   tempPlantingDate = rsMerumputTM("plantingyear")
					   tempBlok = "--ALL--"					   
					 %>
					
					
					
				    <% if sprevPlantingYear <> rsMerumputTM("plantingyear") then %>
				        <td align="center">
							<a href="p_MerumputMonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtStatusTmTbm=<%=tempStatus%>&txtjobcode=<%=tempjobcode%>&txtGroup=<%=tempGroup%>">
							<%=rsMerumputTM("plantingyear")%>
							</a>
						</td>
						<% sprevPlantingYear = rsMerumputTM("plantingyear") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				
                <% if v_txtBlok <> "--SUMMARY--"  and v_txtBlok <> "" then %>
				    <% if sprevFieldCode <> rsMerumputTM("fieldcode") then %>				
					    <td align="center"><%=rsMerumputTM("fieldcode")%></td>				
						<% sprevFieldCode = rsMerumputTM("fieldcode") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>	

				<td align="right"><%=FormatNumber(rsMerumputTM("total_01"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_01"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_01"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_01"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_02"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_02"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_02"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_02"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_03"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_03"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_03"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_03"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_04"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_04"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_04"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_04"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_05"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_05"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_05"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_05"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_06"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_06"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_06"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_06"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_07"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_07"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_07"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_07"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_08"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_08"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_08"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_08"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_09"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_09"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_09"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_09"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_10"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_10"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_10"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_10"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_11"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_11"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_11"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_11"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_12"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_12"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_12"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_12"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("total_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("total_bo_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("hk_ha_sbi"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("Amount_ha_sbi"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsMerumputTM("bud_HK_SBI"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("bud_HA_SBI"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("bud_hk_ha_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("bud_Amount_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsMerumputTM("bud_Amount_ha_sbi"),0)%></td>
			</tr>
<%
			NO = NO + 1			
			ntotal_01 = ntotal_01 + cdbl(rsMerumputTM("total_01"))
			ntotal_bo_01 = ntotal_bo_01 + cdbl(rsMerumputTM("total_bo_01"))
			nhk_ha_01 = nhk_ha_01 + cdbl(rsMerumputTM("hk_ha_01"))
			nAmount_01 = nAmount_01 + cdbl(rsMerumputTM("Amount_01"))
			nAmount_ha_01 = nAmount_ha_01 + cdbl(rsMerumputTM("Amount_ha_01"))
			
			ntotal_02 = ntotal_02 + cdbl(rsMerumputTM("total_02"))
			ntotal_bo_02 = ntotal_bo_02 + cdbl(rsMerumputTM("total_bo_02"))
			nhk_ha_02 = nhk_ha_02 + cdbl(rsMerumputTM("hk_ha_02"))
			nAmount_02 = nAmount_02 + cdbl(rsMerumputTM("Amount_02"))
			nAmount_ha_02 = nAmount_ha_02 + cdbl(rsMerumputTM("Amount_ha_02"))
            
			ntotal_03 = ntotal_03 + cdbl(rsMerumputTM("total_03"))
			ntotal_bo_03 = ntotal_bo_03 + cdbl(rsMerumputTM("total_bo_03"))
			nhk_ha_03 = nhk_ha_03 + cdbl(rsMerumputTM("hk_ha_03"))
			nAmount_03 = nAmount_03 + cdbl(rsMerumputTM("Amount_03"))
			nAmount_ha_03 = nAmount_ha_03 + cdbl(rsMerumputTM("Amount_ha_03"))
			
			ntotal_04 = ntotal_04 + cdbl(rsMerumputTM("total_04"))
			ntotal_bo_04 = ntotal_bo_04 + cdbl(rsMerumputTM("total_bo_04"))
			nhk_ha_04 = nhk_ha_04 + cdbl(rsMerumputTM("hk_ha_04"))
			nAmount_04 = nAmount_04 + cdbl(rsMerumputTM("Amount_04"))
			nAmount_ha_04 = nAmount_ha_04 + cdbl(rsMerumputTM("Amount_ha_04"))
			
			ntotal_05 = ntotal_05 + cdbl(rsMerumputTM("total_05"))
			ntotal_bo_05 = ntotal_bo_05 + cdbl(rsMerumputTM("total_bo_05"))
			nhk_ha_05 = nhk_ha_05 + cdbl(rsMerumputTM("hk_ha_05"))
			nAmount_05 = nAmount_05 + cdbl(rsMerumputTM("Amount_05"))
			nAmount_ha_05 = nAmount_ha_05 + cdbl(rsMerumputTM("Amount_ha_05"))
			
			ntotal_06 = ntotal_06 + cdbl(rsMerumputTM("total_06"))
			ntotal_bo_06 = ntotal_bo_06 + cdbl(rsMerumputTM("total_bo_06"))
			nhk_ha_06 = nhk_ha_06 + cdbl(rsMerumputTM("hk_ha_06"))
			nAmount_06 = nAmount_06 + cdbl(rsMerumputTM("Amount_06"))
			nAmount_ha_06 = nAmount_ha_06 + cdbl(rsMerumputTM("Amount_ha_06"))
			
			ntotal_07 = ntotal_07 + cdbl(rsMerumputTM("total_07"))
			ntotal_bo_07 = ntotal_bo_07 + cdbl(rsMerumputTM("total_bo_07"))
			nhk_ha_07 = nhk_ha_07 + cdbl(rsMerumputTM("hk_ha_07"))
			nAmount_07 = nAmount_07 + cdbl(rsMerumputTM("Amount_07"))
			nAmount_ha_07 = nAmount_ha_07 + cdbl(rsMerumputTM("Amount_ha_07"))
			
			ntotal_08 = ntotal_08 + cdbl(rsMerumputTM("total_08"))
			ntotal_bo_08 = ntotal_bo_08 + cdbl(rsMerumputTM("total_bo_08"))
			nhk_ha_08 = nhk_ha_08 + cdbl(rsMerumputTM("hk_ha_08"))
			nAmount_08 = nAmount_08 + cdbl(rsMerumputTM("Amount_08"))
			nAmount_ha_08 = nAmount_ha_08 + cdbl(rsMerumputTM("Amount_ha_08"))
			
			ntotal_09 = ntotal_09 + cdbl(rsMerumputTM("total_09"))
			ntotal_bo_09 = ntotal_bo_09 + cdbl(rsMerumputTM("total_bo_09"))
			nhk_ha_09 = nhk_ha_09 + cdbl(rsMerumputTM("hk_ha_09"))
			nAmount_09 = nAmount_09 + cdbl(rsMerumputTM("Amount_09"))
			nAmount_ha_09 = nAmount_ha_09 + cdbl(rsMerumputTM("Amount_ha_09"))
			
			ntotal_10 = ntotal_10 + cdbl(rsMerumputTM("total_10"))
			ntotal_bo_10 = ntotal_bo_10 + cdbl(rsMerumputTM("total_bo_10"))
			nhk_ha_10 = nhk_ha_10 + cdbl(rsMerumputTM("hk_ha_10"))
			nAmount_10 = nAmount_10 + cdbl(rsMerumputTM("Amount_10"))
			nAmount_ha_10 = nAmount_ha_10 + cdbl(rsMerumputTM("Amount_ha_10"))
			
			ntotal_11 = ntotal_11 + cdbl(rsMerumputTM("total_11"))
			ntotal_bo_11 = ntotal_bo_11 + cdbl(rsMerumputTM("total_bo_11"))
			nhk_ha_11 = nhk_ha_11 + cdbl(rsMerumputTM("hk_ha_11"))
			nAmount_11 = nAmount_11 + cdbl(rsMerumputTM("Amount_11"))
			nAmount_ha_11 = nAmount_ha_11 + cdbl(rsMerumputTM("Amount_ha_11"))
			
			ntotal_12 = ntotal_12 + cdbl(rsMerumputTM("total_12"))
			ntotal_bo_12 = ntotal_bo_12 + cdbl(rsMerumputTM("total_bo_12"))
			nhk_ha_12 = nhk_ha_12 + cdbl(rsMerumputTM("hk_ha_12"))
			nAmount_12 = nAmount_12 + cdbl(rsMerumputTM("Amount_12"))
			nAmount_ha_12 = nAmount_ha_12 + cdbl(rsMerumputTM("Amount_ha_12"))
			
			ntotal_sbi = ntotal_sbi + cdbl(rsMerumputTM("total_sbi"))
			ntotal_bo_sbi = ntotal_bo_sbi + cdbl(rsMerumputTM("total_bo_sbi"))
			nhk_ha_sbi = nhk_ha_sbi + cdbl(rsMerumputTM("hk_ha_sbi"))
			nAmount_sbi = nAmount_sbi + cdbl(rsMerumputTM("Amount_sbi"))
			nAmount_ha_sbi = nAmount_ha_sbi + cdbl(rsMerumputTM("Amount_ha_sbi"))
			
			ntotal_bud =  ntotal_bud  + cdbl(rsMerumputTM("bud_HK_SBI"))
			ntotal_bo_bud = ntotal_bo_bud + cdbl(rsMerumputTM("bud_HA_SBI"))
			nhk_ha_bud = nhk_ha_bud + cdbl(rsMerumputTM("bud_hk_ha_SBI"))
			nAmount_bud = nAmount_bud + cdbl(rsMerumputTM("bud_Amount_sbi"))
			nAmount_ha_bud = nAmount_ha_bud + cdbl(rsMerumputTM("bud_Amount_ha_sbi"))
			rsMerumputTM.MoveNext
		Wend
	end if
	rsMerumputTM.Close
	set rsMerumputTM = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%= 1+ nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(ntotal_01,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_01,0)%></td>
				<% if cdbl(ntotal_bo_01) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_01/ntotal_bo_01,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_01,0)%></td>
				<% if cdbl(ntotal_bo_01) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_01/ntotal_bo_01,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_02,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_02,0)%></td>
				<% if cdbl(ntotal_bo_02) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_02/ntotal_bo_02,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_02,0)%></td>
				<% if cdbl(ntotal_bo_02) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_02/ntotal_bo_02,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_03,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_03,0)%></td>
				<% if cdbl(ntotal_bo_03) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_03/ntotal_bo_03,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_03,0)%></td>
				<% if cdbl(ntotal_bo_03) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_03/ntotal_bo_03,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_04,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_04,0)%></td>
				<% if cdbl(ntotal_bo_04) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_04/ntotal_bo_04,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_04,0)%></td>
				<% if cdbl(ntotal_bo_04) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_04/ntotal_bo_04,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_05,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_05,0)%></td>
				<% if cdbl(ntotal_bo_05) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_05/ntotal_bo_05,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_05,0)%></td>
				<% if cdbl(ntotal_bo_05) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_05/ntotal_bo_05,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_06,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_06,0)%></td>
				<% if cdbl(ntotal_bo_06) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_06/ntotal_bo_06,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_06,0)%></td>
				<% if cdbl(ntotal_bo_06) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_06/ntotal_bo_06,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_07,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_07,0)%></td>
				<% if cdbl(ntotal_bo_07) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_07/ntotal_bo_07,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_07,0)%></td>
				<% if cdbl(ntotal_bo_07) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_07/ntotal_bo_07,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_08,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_08,0)%></td>
				<% if cdbl(ntotal_bo_08) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_08/ntotal_bo_08,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_08,0)%></td>
				<% if cdbl(ntotal_bo_08) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_08/ntotal_bo_08,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_09,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_09,0)%></td>
				<% if cdbl(ntotal_bo_09) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_09/ntotal_bo_09,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_09,0)%></td>
				<% if cdbl(ntotal_bo_09) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_09/ntotal_bo_09,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_10,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_10,0)%></td>
				<% if cdbl(ntotal_bo_10) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_10/ntotal_bo_10,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_10,0)%></td>
				<% if cdbl(ntotal_bo_10) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_10/ntotal_bo_10,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_11,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_11,0)%></td>
				<% if cdbl(ntotal_bo_11) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_11/ntotal_bo_11,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_11,0)%></td>
				<% if cdbl(ntotal_bo_11) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_11/ntotal_bo_11,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_12,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_12,0)%></td>
				<% if cdbl(ntotal_bo_12) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_12/ntotal_bo_12,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_12,0)%></td>
				<% if cdbl(ntotal_bo_12) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_12/ntotal_bo_12,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				<td align="right"><%=FormatNumber(ntotal_sbi,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_sbi,0)%></td>
				<% if cdbl(ntotal_bo_sbi) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_sbi/ntotal_bo_sbi,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_sbi,0)%></td>
				<% if cdbl(ntotal_bo_sbi) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_sbi/ntotal_bo_sbi,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				
				
				
				<td align="right"><%=FormatNumber(ntotal_bud,0)%></td>
				<td align="right"><%=FormatNumber(ntotal_bo_bud,0)%></td>
				<% if cdbl(ntotal_bo_bud) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(ntotal_bud/ntotal_bo_bud,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
				<td align="right"><%=FormatNumber(nAmount_bud ,0)%></td>
				<% if cdbl(ntotal_bo_bud) > 0 then %>
				    <% if Request("txtGroup") <> ""  then %>
				          <td align="right"><%=FormatNumber(nAmount_bud /ntotal_bo_bud,2)%></td>
					   <% else %>
						   <td align="right"></td>
						<% End if %> 	  
				<% else %>
				   <td align="right"><%=FormatNumber(0,2)%></td>
				<% End if %>
			</tr>
     </tbody>       
</table>

</form>

</div>
</body>
</html>