<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--Modift AMH Mar 21 sesuaikan dengan format dari pa Arif -->
<%	
	Dim rsEPS, strEPS,rscomp , strcomp,	strEPS2
	Dim v_txtyear, v_txtmonth, v_txttodaydate, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass
	Dim strMain, nSpanTotal
	Dim txtyear, txtmonth, txttodaydate, txtDivision, txtBlok, txtPlantingDate,txtComp,v_txtComp
    Dim tempDivision, tempBlok, tempPlantingDate 
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	Dim nSubMandays_Amount_hi, nSubMandays_Amount_shi
	Dim nOutput_Mandays_Amount_hi, nOutput_Mandays_Amount_shi
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType, txtstatus
	
	        v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	= Request("txtPlantation")
			v_txtGroupRegion  	= Request("txtGroupRegion")
			v_txtRegion  		= Request("txtRegion")
			v_txtcompany  		= Request("txtCompany")
			v_txtfcba 			= Request("txtfcba")
			v_txtDivision		= Request("txtDivision")
			v_txtBlok			= Request("txtBlok")
			v_txtPlantingDate	= Request("txtPlantingDate")
			v_txttodaydate		= Request("txttodaydate")
			
			txtstatus = Request("txtstatus")
			
			if "" & txtstatus = "" then txtstatus = "TM"
			
			txtyear 		= right(Request("txttodaydate"),4)
			txtmonth 		= left(Request("txttodaydate"),2)
			txttodaydate 	= Request("txttodaydate")
			
			txtComp         = Request("txtCompany")
			nSpanTotal = 8
			v_txtComp           =   txtComp
			v_txtyear			=	txtyear
			v_txtmonth			=	txtmonth
			
			
			
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
			
			Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut, " 

						   
			wherecons = " where company is not null "
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
			
			
			
			   if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
			      if v_txtDivision <>  "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      end if
                  selectCons = selectCons &  " , DIVISION "
				  orderCons  = orderCons & ",DIVISION "
			   else
			    nSpanTotal = nSpanTotal - 1 	  
			   end if   

            
			
				if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   end if
				   selectCons = selectCons &  " , PLANTINGYEAR "
				   orderCons  = orderCons & ",PLANTINGYEAR "
				else
			    nSpanTotal = nSpanTotal - 1    
				end if
		
			
				if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
				   if v_txtBlok <>  "--ALL--" then
					  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
				   end if
				   selectCons = selectCons & " , FIELDCODE, FIELDHA, FIELDPOKOK "
				   orderCons  = orderCons & ",FIELDCODE, FIELDHA, FIELDPOKOK "
				else
			    nSpanTotal = nSpanTotal - 1    
				end if
            
	Dim wherestatus				
	if "" & txtstatus <> "" and "" & txtstatus <> "--ALL--" then
		wherestatus = " and field_api.GetValid_status(fieldcode,fcba,TO_DATE('" & v_txttodaydate & "', 'mm/dd/yyyy') ) = '" & txtstatus & "'"
	end if
	
	strMain = "select field_api.get_HA(fieldcode,fcba) FieldHA,field_api.get_Pokok(fieldcode,fcba) FieldPokok ,t.* from PRODUCTION_DAILY_REPORT_TAB t " & _
              " where t.fddate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') and USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) ='TRUE' " 
		
	 strEPS2  =  " ,SUM(hi) hi, " & _
				 "  SUM(shi) shi, " & _
				 "  SUM(sbi) sbi, " & _
				 "  SUM(tonase_hi) tonase_hi, " & _
				 "  SUM(tonase_shi) tonase_shi, " & _
				 "  SUM(tonase_sbi) tonase_sbi,  " & _
				 "  SUM(output_hi) output_hi, " & _
				 "  SUM(output_shi) output_shi, " & _
				 "  SUM(output_sbi) output_sbi,  " & _
				 "  SUM(pom_hi) pom_hi, " & _
				 "  SUM(pom_shi) pom_shi, " & _
				 "  SUM(pom_sbi) pom_sbi, " & _
				 "  sum(bud_jjg_shi) bud_jjg_shi,sum(bud_jjg_sbi) bud_jjg_sbi, " & _
				 "  SUM(pom_tonase_hi) pom_tonase_hi, " & _
				 "  SUM(pom_tonase_shi) pom_tonase_shi, " & _
				 "  SUM(pom_tonase_sbi) pom_tonase_sbi, " & _
				 "  sum(bud_qty_shi) bud_qty_shi,sum(bud_qty_sbi) bud_qty_sbi, " & _
				 "  sum(wo_plus_hi)  wo_plus_hi,  " & _ 
                 "  sum(wo_plus_shi) wo_plus_shi,  " & _
                 "  sum(wo_plus_sbi) wo_plus_sbi,  " & _
                 "  sum(wo_min_hi)   wo_min_hi,   " & _
                 "  sum(wo_min_shi)  wo_min_shi,  " & _
                 "  sum(wo_min_sbi)  wo_min_sbi,  " & _
				 "  SUM(restan) restan, " & _
				 "  SUM(mandays_hi) mandays_hi, " & _
				 "  SUM(mandays_shi) mandays_shi, " & _
				 "  SUM(mandays_sbi) mandays_sbi, " & _
				 "  sum(bud_mandays_shi) bud_mandays_shi,sum(bud_mandays_sbi) bud_mandays_sbi, " & _
				 "  SUM(mandays_amount_hi) mandays_amount_hi, " & _
				 "  SUM(mandays_amount_shi) mandays_amount_shi,	" & _	
                 "  sum(bud_output_shi) bud_output_shi,sum(bud_output_sbi) bud_output_sbi " 				 
								 
								 

	strEPS =  "select GroupPlantation " & selectCons & strEPS2 &"  From (" & strMain & " )" & whereCons & wherestatus & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by  GroupPlantation" & selectCons & "  Order by GroupPlantation" & orderCons  	

	

				
Set rsEPS = Server.CreateObject("ADODB.Recordset")
'response.write strEPS

rsEPS.Open strEPS , connstr, adOpenStatic 


%>

<html>

<head>
<title>ePlaS Reporting - Laporan Harian Produksi</title>
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
<form name="frmEntry" method="post" action="p_ProductionDaily.asp">


	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	

<!--#include file="FormSelectNewSum.html"-->
<table border="0" id="form-table" class="centered-margin">
    
	<tr>
		<th>TAHUN</th>
		<th>BULAN</th>
		<th>TANGGAL</th>
		<th>STATUS</th>
		<th>AFDELING</th>
		<th>TAHUN TANAM</th>
		<th>BLOK</th>
		<th>&nbsp;</th>
	</tr>
	<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
	<tr>
	    <!--%response.write v_txtfcba %-->
	    <!--%response.write v_txtyear %-->
	    <!--%response.write v_txtmonth %-->
		<!--%response.write len(trim(request("txtYear"))) %-->
		<td align="center"><input type="text" name="txtyear" value="<%=v_txtyear%>"></td>
		<td align="center"><input type="text" name="txtmonth" value="<%=v_txtmonth%>"></td>
		<td align="center">
		<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
			<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
		<% else%>
			<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<!--<a class="calendar" href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>-->
		</td>
		
		<td align="center">
			<select name="txtstatus">	
				<option value="--ALL--" <%if txtstatus = "--ALL--" then%> Selected <%end if%>>--ALL--</option>
				<option value="TM" <%if txtstatus = "TM" then%> Selected <%end if%>>TM</option>
				<option value="TBM" <%if txtstatus = "TBM" then%> Selected <%end if%>>TBM</option>
			</select>		
		</td>
		
		<td align="center">
			<select name="txtDivision">	
				<option value="--SUMMARY--">--SUMMARY--</option>
				<%
					Dim rsDivision
					Set rsDivision = Server.CreateObject("ADODB.Recordset")
					rsDivision.CursorLocation = adUseClient
					

			
					if len(trim(v_txtDivision)) > 0 then 
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "'"
					else
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELDDETAIL WHERE FCBA = '" & Session("fcBA") & "'"
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
				
						if len(trim(v_txtPlantingDate)) > 0 then 
							strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM Budget_Field WHERE FCBA = '" & v_txtFcba & "'"
						else
							strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM Budget_Field WHERE FCBA = '" & Session("fcBA") & "'"
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
						strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELDDETAIL WHERE FCBA = '" & v_txtFcba & "'"
					else
						strsql =  " SELECT '--ALL--' as Blok from DUAL UNION SELECT DISTINCT fccode as Blok FROM FIELDDETAIL WHERE FCBA = '" & Session("fcBA") & "'"
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
</table>



<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
<tr>
<th colspan="38" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
<div id="header-wrapper">
		<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
		</div>
		<div id="subheader"class="centered">
			<h2>LAPORAN HARIAN PRODUKSI</h2>
			<% if v_txttodaydate <> "" and v_txtFcba <> "" then %>
				<h5>PERIODE : <%=v_txttodaydate%>, KEBUN : <%=v_txtFcba%></h5>
			<% end if %>
			<h3></br></h3>
			<h3></br></h3>
		</div>
	</div>

	<%
		if len(trim(request("txtYear"))) = 0 then
			v_txtyear = year(Now)	 
		end if

		if len(trim(request("txtmonth"))) = 0 then
			v_txtmonth = Month(Now)	 
		end if
	%>

</th>
</tr>
	<tr>
		 <th rowspan="4" style="  background-color:#00A300;color:#fff; "></th>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<% 			
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<% 			
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<% 		
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<% 		
	    end if %>

		
		<% if v_txtDivision <> "--SUMMARY--"   and v_txtDivision <> "" then %>
		    <th rowspan="4" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<% 		
	    end if %>
		
		
		<% if v_txtPlantingDate <> "--SUMMARY--" And v_txtPlantingDate <> "" then %>
		    <th rowspan="4" style="background-color:#00A300;color:#fff;">TT</th>
		<% 		
	    end if %>


		<% if v_txtBlok <> "--SUMMARY--" And v_txtBlok <> "" then %>
		    <th rowspan="4" style="background-color:#00A300;color:#fff;">BLOK</th>
		<% 		
	    end if %>
		
		<th colspan="2" style="background-color:#00A300;color:#fff;">PRODUKSI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">BUDGET (Ton)</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">HA</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">HK</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JJG (Panen)</th>
		<!--<th colspan="3" style="background-color:#00A300;color:#fff;">JJG (Afkir)</th>-->
		<th colspan="3" style="background-color:#00A300;color:#fff;">JJG (PKS)</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">TONASE (PKS)</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">BJR</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">ROTASI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JJG/PKK</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">TON/HK</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">TON/HA</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">HA/HK</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JJG/HK</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">AKP</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">%PENCAPAIAN</th>
	</tr>
	
	
	<tr>
	<!-- Produksi-->
		<th  style="background-color:#00A300;color:#fff;" >Luas</th>
		<th  style="background-color:#00A300;color:#fff;" >Pokok</th>
	<!-- Budget-->	
		<th  style="background-color:#00A300;color:#fff;" >BI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		<th  style="background-color:#00A300;color:#fff;" >THN</th>
	<!-- HA-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- HK-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- JJG Panen-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- JJG Afkir-->	
	<!--	<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>-->
	<!-- JJG PKS-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- TON PKS-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
    <!-- BJR-->		
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- ROTASI-->		
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- JJG/PKK-->
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- TON/HK-->
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- TON/HA-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- HA/HK-->
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- JJG/HK-->
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- AKP-->
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- %-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
		
	</tr>
        </thead>
 <tbody>
<%
Dim nsubhi, nsubshi, nsubbi, nsubsbi, nsubtonase_hi
Dim nsubtonase_shi, nsubtonase_bi, nsubtonase_sbi, nsubpom_hi, nsubpom_shi
Dim nsubpom_bi, nsubpom_sbi, nsubpom_tonase_hi, nsubpom_tonase_shi, nsubpom_tonase_bi, nsubpom_tonase_sbi, nsubrestan
Dim nsubbudjjg_shi,nsubbudjjg_sbi,nsubbudqty_shi,nsubbudqty_sbi,nsubbudmandays_shi,nsubbudmandays_sbi,nsubbudOutput_shi,nsubbudOutput_sbi
Dim nsubmandays_hi, nsubmandays_shi, nsubmandays_bi, nsubmandays_sbi
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim nsubOutput_hi, nsubOutput_shi, nsubOutput_bi, nsubOutput_sbi
Dim nOutput_hi, nOutput_shi, nOutput_bi, nOutput_sbi ,nbudOutput_shi,nbudOutput_sbi
Dim nsubwo_plus_hi, nsubwo_plus_shi, nsubwo_plus_sbi,nsubwo_min_hi, nsubwo_min_shi, nsubwo_min_sbi
Dim nfieldha,nfieldpokok
Dim nsubhahi,nsubhashi,nsubhasbi

	if not rsEPS.EOF then
		NO = 1
		nfieldha = 0
		nfieldpokok = 0
		nsubhahi = 0
		nsubhashi = 0
		nsubhasbi = 0
		nsubhi = 0
		nsubshi = 0
		nsubbi = 0
		nsubsbi = 0
		nsubtonase_hi = 0
		nsubtonase_shi = 0
		nsubtonase_bi = 0
		nsubtonase_sbi = 0
		nsubpom_hi = 0
		nsubpom_shi = 0
		nsubpom_bi = 0
		nsubpom_sbi = 0
		nsubpom_tonase_hi = 0
		nsubpom_tonase_shi = 0
		nsubpom_tonase_bi = 0
		nsubpom_tonase_sbi = 0
		
		nsubwo_plus_hi=0 
		nsubwo_plus_shi=0 
		nsubwo_plus_sbi=0
		nsubwo_min_hi=0
		nsubwo_min_shi=0
		nsubwo_min_sbi=0
		
		nsubrestan = 0
		
		nbudOutput_shi = 0
		nbudOutput_sbi = 0
		
		nsubmandays_hi = 0
		nsubmandays_shi = 0
		nsubmandays_bi = 0
		nsubmandays_sbi = 0
		
		nsubOutput_hi = 0
		nsubOutput_shi = 0
		nsubOutput_bi = 0
		nsubOutput_sbi = 0
		
		nSubMandays_Amount_hi = 0
		nSubMandays_Amount_shi = 0
		
		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		while not rsEPS.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"></td>
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% tempPlantation = rsEPS("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsEPS("Plantation") then %>
						<td align="center">
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsEPS("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsEPS("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsEPS("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsEPS("GroupRegion") then %>
						<td align="center">
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsEPS("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsEPS("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsEPS("Region") then %>
						<td align="center">
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("region")%>
							</a>
						</td>
						<% SprevRegion = rsEPS("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsEPS("Company")
					   tempfcba = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsEPS("Company") then %>
						<td align="center">
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("company")%>
							</a>
						</td>
						<% SprevCompany = rsEPS("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsEPS("fcba")
					   tempDivision = "--ALL--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsEPS("fcba") then %>
						<td align="center">
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
						    <%=rsEPS("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsEPS("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtDivision <> "--SUMMARY--" then %>
				   <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsEPS("fcba")
					   else
					       tempfcba = ""
					   end if
					   tempDivision = rsEPS("division")
					   tempPlantingDate = "--ALL--"
					   tempBlok = "--SUMMARY--"					   
					 %>
				
				
				    <% if sprevDivision <> rsEPS("division") then %>
					    <td align="center">
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("division")%>
							</a>
						</td>
						<% sprevDivision = rsEPS("division") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>

				<% if v_txtPlantingDate <> "--SUMMARY--" then %>
				   <%  if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
					       tempfcba = rsEPS("fcba")
					   else
					       tempfcba = ""
					   end if
					   
					   if v_txtDivision <>  "--SUMMARY--" and v_txtDivision <>  "" then
					       tempDivision = rsEPS("division")
					   else
					       tempDivision = ""
					   end if
					   
					   
					   tempPlantingDate = rsEPS("plantingyear")
					   tempBlok = "--ALL--"					   
					 %>
					
					<% tempPlantingDate = rsEPS("plantingyear") %>
					
					
				    <% if sprevPlantingYear <> rsEPS("plantingyear") then %>
				        <td align="center">
							<a href="p_ProductionDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("plantingyear")%>
							</a>
						</td>
						<% sprevPlantingYear = rsEPS("plantingyear") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>


				<% if v_txtBlok <> "--SUMMARY--" then %>
				    <% if sprevFieldCode <> rsEPS("fieldcode") then %>				
					    <td align="center"><%=rsEPS("fieldcode")%></td>				
						<% sprevFieldCode = rsEPS("fieldcode") %>
						<td align="center"><%=rsEPS("fieldha")%></td>	
						<td align="center"><%=rsEPS("fieldpokok")%></td>	
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>				
				
				<!-- Budget-->
				<td align="right"><%=FormatNumber(cdbl(rsEPS("bud_qty_shi"))/1000,2)%></td>
                <td align="right"><%=FormatNumber(cdbl(rsEPS("bud_qty_sbi"))/1000,2)%></td>
				<td align="right">0</td>
				
				<!-- HA -->
				<td align="right"><%=FormatNumber(rsEPS("output_hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("output_shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("output_sbi"),2)%></td>
				<!-- HK -->
				<td align="right"><%=FormatNumber(rsEPS("mandays_hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("mandays_shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("mandays_sbi"),2)%></td>
				<!-- JJG -->
				<td align="right"><%=FormatNumber(rsEPS("hi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("shi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("sbi"),0)%></td>
				<!-- JJG Afkir -->
				<!--<td align="right">0</td>-->
				<!--<td align="right">0</td>-->
				<!--<td align="right">0</td>-->				
				<!-- JJG PKS -->
				<td align="right"><%=FormatNumber(rsEPS("pom_hi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_shi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_sbi"),0)%></td>
				<!-- Ton PKS -->
				<td align="right"><%=FormatNumber(cdbl(rsEPS("pom_tonase_hi"))/1000,2)%></td>
				<td align="right"><%=FormatNumber(cdbl(rsEPS("pom_tonase_shi"))/1000,2)%></td>
				<td align="right"><%=FormatNumber(cdbl(rsEPS("pom_tonase_sbi"))/1000,2)%></td>
				
				<!-- BJR -->
				<% if cdbl(rsEPS("pom_hi")) <> 0 then %>
				    <% nOutput_hi = cdbl(rsEPS("pom_tonase_hi"))  / cdbl(rsEPS("pom_hi")) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("pom_shi")) <> 0 then %>
					<% nOutput_shi = cdbl(rsEPS("pom_tonase_shi")) / cdbl(rsEPS("pom_shi")) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("pom_sbi")) <> 0 then %>
					<% nOutput_sbi = cdbl(rsEPS("pom_tonase_sbi"))  / cdbl(rsEPS("pom_sbi")) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
				
				<td align="right"><%=FormatNumber(nOutput_hi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,2)%></td>
				
				<!-- Rotasi -->
				<% if cdbl(rsEPS("fieldha")) <> 0 then %>
				    <% nOutput_hi = cdbl(rsEPS("output_hi"))  / cdbl(rsEPS("fieldha")) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("fieldha")) <> 0 then %>
					<% nOutput_shi = cdbl(rsEPS("output_shi")) / cdbl(rsEPS("fieldha")) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("fieldha")) <> 0 then %>
					<% nOutput_sbi = cdbl(rsEPS("output_sbi"))  / cdbl(rsEPS("fieldha")) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
				
				<td align="right"><%=FormatNumber(nOutput_hi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,2)%></td>
				
				<!-- JJG/PKK -->
				<% if cdbl(rsEPS("FIELDPOKOK")) <> 0 then %>
				    <% nOutput_hi = cdbl(rsEPS("hi"))  / cdbl(rsEPS("FIELDPOKOK")) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("FIELDPOKOK")) <> 0 then %>
					<% nOutput_shi = cdbl(rsEPS("shi")) / cdbl(rsEPS("FIELDPOKOK")) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>

				
				<% if cdbl(rsEPS("FIELDPOKOK")) <> 0 then %>
					<% nOutput_sbi = cdbl(rsEPS("sbi"))  / cdbl(rsEPS("FIELDPOKOK")) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
		
				
				<td align="right"><%=FormatNumber(nOutput_hi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,0)%></td>
				
				
				<!-- TON/HK -->
				<% if cdbl(rsEPS("mandays_hi")) <> 0 then %>
				    <!-- Remarked by NAF, to remove tonage of QTY -->
				    <!--% nOutput_hi = cdbl(rsEPS("pom_tonase_hi")) * 1000 / cdbl(rsEPS("mandays_hi")) %-->
				    <% nOutput_hi = (cdbl(rsEPS("pom_tonase_hi")) /1000) / cdbl(rsEPS("mandays_hi")) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("mandays_shi")) <> 0 then %>
					<!--% nOutput_shi = cdbl(rsEPS("pom_tonase_shi")) * 1000 / cdbl(rsEPS("mandays_shi")) %-->
					<% nOutput_shi = (cdbl(rsEPS("pom_tonase_shi"))/1000) / cdbl(rsEPS("mandays_shi")) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>

				
				<% if cdbl(rsEPS("mandays_sbi")) <> 0 then %>
					<!--% nOutput_sbi = cdbl(rsEPS("pom_tonase_sbi")) * 1000 / cdbl(rsEPS("mandays_sbi")) %-->
					<% nOutput_sbi = (cdbl(rsEPS("pom_tonase_sbi"))/1000)  / cdbl(rsEPS("mandays_sbi")) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
		
				
				<td align="right"><%=FormatNumber(nOutput_hi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,2)%></td>
				
				
				<!-- TON/HA -->
				<% if cdbl(rsEPS("FIELDHA")) <> 0 then %>
				    <% nOutput_hi = (cdbl(rsEPS("pom_tonase_hi"))/1000)  / cdbl(rsEPS("FIELDHA")) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("FIELDHA")) <> 0 then %>
					<% nOutput_shi = (cdbl(rsEPS("pom_tonase_shi"))/1000) / cdbl(rsEPS("FIELDHA")) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>

				
				<% if cdbl(rsEPS("FIELDHA")) <> 0 then %>
					<% nOutput_sbi = (cdbl(rsEPS("pom_tonase_sbi"))/1000)  / cdbl(rsEPS("FIELDHA")) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
		
				
				<td align="right"><%=FormatNumber(nOutput_hi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,2)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,2)%></td>
				
				<!-- HA/HK -->
				<% if cdbl(rsEPS("mandays_hi")) <> 0 then %>
				    <!-- Remarked by NAF, to remove tonage of QTY -->
				    <!--% nOutput_hi = cdbl(rsEPS("pom_tonase_hi")) * 1000 / cdbl(rsEPS("mandays_hi")) %-->
				    <% nOutput_hi = cdbl(rsEPS("output_hi"))  / cdbl(rsEPS("mandays_hi")) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("mandays_shi")) <> 0 then %>
					<!--% nOutput_shi = cdbl(rsEPS("pom_tonase_shi")) * 1000 / cdbl(rsEPS("mandays_shi")) %-->
					<% nOutput_shi = cdbl(rsEPS("output_shi")) / cdbl(rsEPS("mandays_shi")) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>

				
				<% if cdbl(rsEPS("mandays_sbi")) <> 0 then %>
					<!--% nOutput_sbi = cdbl(rsEPS("pom_tonase_sbi")) * 1000 / cdbl(rsEPS("mandays_sbi")) %-->
					<% nOutput_sbi = cdbl(rsEPS("output_sbi"))  / cdbl(rsEPS("mandays_sbi")) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
		
				
				<td align="right"><%=FormatNumber(nOutput_hi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,0)%></td>
				
				<!-- JJG/HK -->
				<% if cdbl(rsEPS("mandays_hi")) <> 0 then %>
				    <% nOutput_hi = cdbl(rsEPS("hi"))  / cdbl(rsEPS("mandays_hi")) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("mandays_shi")) <> 0 then %>
					<% nOutput_shi = cdbl(rsEPS("shi")) / cdbl(rsEPS("mandays_shi")) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>

				
				<% if cdbl(rsEPS("mandays_sbi")) <> 0 then %>
					<% nOutput_sbi = cdbl(rsEPS("sbi"))  / cdbl(rsEPS("mandays_sbi")) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
		
				
				<td align="right"><%=FormatNumber(nOutput_hi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,0)%></td>
				
				<!-- AKP -->
				<% if (cdbl(rsEPS("fieldha"))*133) <> 0 then %>
				    <% nOutput_hi = cdbl(rsEPS("hi"))  / (cdbl(rsEPS("fieldha"))*133) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if (cdbl(rsEPS("fieldha"))*133) <> 0 then %>
					<% nOutput_shi = cdbl(rsEPS("shi")) / (cdbl(rsEPS("fieldha"))*133) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>

				
				<% if (cdbl(rsEPS("fieldha"))*133) <> 0 then %>
					<% nOutput_sbi = cdbl(rsEPS("sbi"))  / (cdbl(rsEPS("fieldha"))*133) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
		
				
				<td align="right"><%=FormatNumber(nOutput_hi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,0)%></td>
				
				<!-- % -->
				<% if cdbl(rsEPS("bud_qty_shi")) <> 0 then %>
				    <% nOutput_hi = cdbl(rsEPS("pom_tonase_hi"))  / cdbl(rsEPS("bud_qty_shi")) %>
				<% else %>
				    <% nOutput_hi = 0 %>
				<% end if %>
				
				<% if cdbl(rsEPS("bud_qty_shi")) <> 0 then %>
					<% nOutput_shi = cdbl(rsEPS("pom_tonase_sbi")) / cdbl(rsEPS("bud_qty_shi")) %>
				<% else %>
				    <% nOutput_shi = 0 %>
				<% end if %>

				
				<% if cdbl(rsEPS("bud_qty_sbi")) <> 0 then %>
					<% nOutput_sbi = cdbl(rsEPS("pom_tonase_sbi"))  / cdbl(rsEPS("bud_qty_sbi")) %>
				<% else %>
				    <% nOutput_sbi = 0 %>
				<% end if %>
		
				
				<td align="right"><%=FormatNumber(nOutput_hi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_shi,0)%></td>
				<td align="right"><%=FormatNumber(nOutput_sbi,0)%></td>
				
				
				
			</tr>
<%
			NO = NO + 1
			
			nfieldha= nfieldha + Cdbl(rsEPS("fieldha"))
			nfieldpokok=nfieldpokok+ Cdbl(rsEPS("fieldpokok")) 
			
			nsubbudqty_shi = nsubbudqty_shi + (Cdbl(rsEPS("bud_qty_shi"))/1000)
			nsubbudqty_sbi = nsubbudqty_sbi + (Cdbl(rsEPS("bud_qty_sbi"))/1000)
			
			nsubOutput_hi = nsubOutput_hi + Cdbl(rsEPS("output_hi")) 
            nsubOutput_shi = nsubOutput_shi + Cdbl(rsEPS("output_shi"))
			nsubOutput_bi = nsubOutput_bi + Cdbl(rsEPS("output_sbi")) 
			
			nsubmandays_hi = nsubmandays_hi + Cdbl(rsEPS("mandays_hi"))
			nsubmandays_shi = nsubmandays_shi + Cdbl(rsEPS("mandays_shi"))
			nsubmandays_sbi = nsubmandays_sbi + Cdbl(rsEPS("mandays_sbi"))		
			
			nsubhi = nsubhi + Cdbl(rsEPS("hi"))
			nsubshi = nsubshi + Cdbl(rsEPS("shi"))			
			nsubsbi = nsubsbi + Cdbl(rsEPS("sbi"))
			
			
			
			nsubpom_hi = nsubpom_hi + Cdbl(rsEPS("pom_hi"))
			nsubpom_shi = nsubpom_shi + Cdbl(rsEPS("pom_shi"))
			nsubpom_sbi = nsubpom_sbi + Cdbl(rsEPS("pom_sbi"))
			
						
			nsubpom_tonase_hi = nsubpom_tonase_hi + (Cdbl(rsEPS("pom_tonase_hi"))/1000)
			nsubpom_tonase_shi = nsubpom_tonase_shi + (Cdbl(rsEPS("pom_tonase_shi"))/1000)
			nsubpom_tonase_sbi = nsubpom_tonase_sbi + (Cdbl(rsEPS("pom_tonase_sbi"))/1000)
			
			
			rsEPS.MoveNext			
		Wend
	end if
	rsEPS.Close
	set rsEPS = Nothing
	
	' tambahan akbar
	
	
	
	
	if nsubpom_tonase_hi > 0 then
	   'nSubMandays_Amount_hi = (nsubmandays_amount_hi/(nsubpom_tonase_hi*1000)) 
	   nSubMandays_Amount_hi = (nsubmandays_amount_hi/(nsubpom_tonase_hi)) 
	else
	   nSubMandays_Amount_hi = 0
	end if
	if nsubpom_tonase_shi > 0 then
	   'nSubMandays_Amount_shi = (nsubmandays_amount_shi/(nsubpom_tonase_shi*1000)) 
	   nSubMandays_Amount_shi = (nsubmandays_amount_shi/(nsubpom_tonase_shi)) 
	else
	   nSubMandays_Amount_shi = 0
	end if
	
	' tambahan akbar
%>
			<tr bgcolor="#F4F4A0">
			   
				<td align="center" colspan="<%= 1+ nSpanTotal %>">TOTAL</td>
				<!-- Produksi -->
				<td align="right"><%=FormatNumber(nfieldha,0)%></td>
				<td align="right"><%=FormatNumber(nfieldpokok,0)%></td>
				<!-- Budget -->
				<td align="right"><%=FormatNumber(nsubbudqty_shi,0)%></td>
				<td align="right"><%=FormatNumber(nsubbudqty_sbi,0)%></td>
				<td align="right">0</td>
				<!-- HA -->
				<td align="right"><%=FormatNumber(nsubOutput_hi,0)%></td>
				<td align="right"><%=FormatNumber(nsubOutput_shi,0)%></td>
				<td align="right"><%=FormatNumber(nsubOutput_sbi,0)%></td>
				<!-- HK -->
				<td align="right"><%=FormatNumber(nsubmandays_hi,0)%></td>
				<td align="right"><%=FormatNumber(nsubmandays_shi,0)%></td>
				<td align="right"><%=FormatNumber(nsubmandays_sbi,0)%></td>
				
				<!-- JJG -->
				<td align="right"><%=FormatNumber(nsubhi,0)%></td>
				<td align="right"><%=FormatNumber(nsubshi,0)%></td>
				<td align="right"><%=FormatNumber(nsubsbi,0)%></td>
				<!-- JJG PKS -->
				<td align="right"><%=FormatNumber(nsubpom_hi,0)%></td>
				<td align="right"><%=FormatNumber(nsubpom_shi,0)%></td>
				<td align="right"><%=FormatNumber(nsubpom_sbi,0)%></td>
				<!-- TON PKS -->
				<td align="right"><%=FormatNumber(nsubpom_tonase_hi,0)%></td>
				<td align="right"><%=FormatNumber(nsubpom_tonase_shi,0)%></td>
				<td align="right"><%=FormatNumber(nsubpom_tonase_sbi,0)%></td>
				
				
				
				
				
			</tr>
</tbody>
</table>
</form>
</div>
</body>
</html>

