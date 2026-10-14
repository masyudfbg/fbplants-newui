<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%	
	Dim rsEPS, strEPS,rscomp , strcomp,	strEPS2
	Dim v_txtyear, v_txtmonth, v_txttodaydate, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass
	Dim strMain, nSpanTotal
	Dim txtyear, txtmonth,txtday, txttodaydate, txtDivision, txtBlok, txtPlantingDate,txtComp,v_txtComp
    Dim tempDivision, tempBlok, tempPlantingDate ,tempharvestingsection
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempGroupPlantation
	Dim nSubMandays_Amount_hi, nSubMandays_Amount_shi
	Dim nOutput_Mandays_Amount_hi, nOutput_Mandays_Amount_shi
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode,SprevPlantingDate
	Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation,Sprevharvestingsection,SprevBlok
	
		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		Sprevharvestingsection = ""
		SprevPlantingDate = ""
		SprevBlok = ""
		
		
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
			
			txtyear 		= YEAR(Request("txttodaydate"))
			txtmonth 		= MONTH(Request("txttodaydate"))
			txtday			= DAY(Request("txttodaydate")) 
			txttodaydate 	= Request("txttodaydate")
			
			txtComp         = Request("txtCompany")
			nSpanTotal = 8
			v_txtComp           =   txtComp
			v_txtyear			=	txtyear
			v_txtmonth			=	txtmonth
			
			
			
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
			
			Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons,orderBlok,orderPlantingDate
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
				   orderPlantingDate  = ",PLANTINGYEAR "
				else
				   nSpanTotal = nSpanTotal - 1 
				end if
		
			
				if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
				   if v_txtBlok <>  "--ALL--" then
					  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
				   end if
				   selectCons = selectCons & " , FIELDCODE,FCNAME "
				   orderBlok  = ",FIELDCODE "
				else
					nSpanTotal = nSpanTotal - 1 
				end if
				
				
    strMain = 	"select " & strFieldCons & "f.fccode FIELDCODE,f.fcname," & _
					"field_api.getValid_Afdeling(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) DIVISION," & _
					"field_api.GetValid_Harvestingsection(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) harvestingsection," & _
					"field_api.getvalid_plantingyear(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) plantingyear," & _
					"field_api.GetValid_hectarageplanted(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) hectarageplanted," & _
					"field_api.GetValid_harvestingbased_abw(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) harvestingbased_abw," & _
					"field_api.GetValid_CropType(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) croptype," & _
					"field_api.GetValid_terraintype(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) terraintype," & _
					"opharvrate_api.Get_Basis_Qty(field_api.GetValid_harvestingbased_abw(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY'))," & _
					"field_api.GetValid_CropType(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY'))," & _
					"field_api.GetValid_terraintype(" & _
						"f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY')),f.fcba) basis_qty," & _
					"opharvrate_api.Get_Basis_Rate(" & _
						"field_api.GetValid_harvestingbased_abw(" & _
							"f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY'))," & _
						"field_api.GetValid_CropType(f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY'))," & _
						"field_api.GetValid_terraintype(f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY')),f.fcba) Basis_Rate," & _
					"opharvrate_api.Get_Extra_Basis_Rate(" & _
						"field_api.GetValid_harvestingbased_abw(" & _
							"f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY'))," & _
						"field_api.GetValid_CropType(f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY'))," & _
						"field_api.GetValid_terraintype(f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY')),f.fcba) Extra_Basis_Rate," & _
					"opharvrate_api.Get_Basis_Ha(" & _
						"field_api.GetValid_harvestingbased_abw(" & _
							"f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY'))," & _
						"field_api.GetValid_CropType(f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY'))," & _
						"field_api.GetValid_terraintype(f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY')),f.fcba) Basis_Ha " & _
				"from field f " & _
				"where field_api.GetValid_status(f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) = 'TM' " & _
				"and field_api.GetValid_harvestingbased_abw(f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) is not null " & _
				" And   field_api.GetValid_Activation(f.fccode,f.fcba,to_date('" & txttodaydate & "','MM/DD/YYYY'))='Y' " & _
				"and field_api.GetValid_Harvestingsection(f.fccode,f.fcba,to_date('" & txttodaydate & "', 'MM/DD/YYYY')) is not null"
				
	strEPS2  =  " ,harvestingsection," & _
				"sum(hectarageplanted)  hectarageplanted," & _
				"sum(harvestingbased_abw)  harvestingbased_abw," & _
			    "sum(basis_qty)  basis_qty," & _
				"sum(Basis_Rate)  Basis_Rate," & _
				"sum(Extra_Basis_Rate)  Extra_Basis_Rate," & _
				"sum(Basis_Ha)  Basis_Ha"
				
	strEPS =  "select GroupPlantation " & selectCons & strEPS2 &"  From (" & strMain & " )" & whereCons & " "& _
			  "and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' "& _
			  "group by GroupPlantation" & selectCons & _
			  ",harvestingsection " & _
			  "Order by GroupPlantation" & orderCons & ",harvestingsection" & orderPlantingDate & orderBlok 


'response.write strEPS				
Set rsEPS = Server.CreateObject("ADODB.Recordset")


rsEPS.Open strEPS , connstr, adOpenStatic 


%>

<html>

<head>
<title>ePlaS Reporting - Laporan Harian Produksi</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<style>
.janjang{font-size:10px;background-color:#00A300;color:#fff; }
</style>
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
<form name="frmEntry" method="post" action="p_LaporanPremiPanen.asp">


	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	

<!--#include file="FormSelectNewSum.html"-->
<table border="0" id="form-table" class="centered-margin">
    
	<tr>
		<th>TANGGAL</th>
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
			<select name="txtDivision">	
				<option value="--SUMMARY--">--SUMMARY--</option>
				<%
					Dim rsDivision
					Set rsDivision = Server.CreateObject("ADODB.Recordset")
					rsDivision.CursorLocation = adUseClient
					

			
					if len(trim(v_txtDivision)) > 0 then 
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
					else
						strsql =  " SELECT '--ALL--' as DIVISION from DUAL UNION SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
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
							strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & v_txtFcba & "'"
						else
							strsql =  " SELECT '--ALL--' as PlantingDate from DUAL UNION SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
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
</table>



<table border="1" width="100%" id="product-table" class="xthf-site">
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
			<h2>SISTEM PREMI POTONG BUAH</h2>
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
		
		<%if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<%End if %>
		<th style="background-color:#00A300;color:#fff;">SEKSI PANEN</th>
		<%if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">TAHUN TANAM</th>
		<%End if %>
		
		<%if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then%>
		<th style="background-color:#00A300;color:#fff;">BLOK</th>
		<th style="background-color:#00A300;color:#fff;">BLOK NAME</th>
		<%End if %>
		<th style="background-color:#00A300;color:#fff;">LUAS</th>	
        <th style="background-color:#00A300;color:#fff;">BJR</th>			
		<th style="background-color:#00A300;color:#fff;">BASIS JANJANG <br/><font size="1">(Jjg)</font></th>
		<th style="background-color:#00A300;color:#fff;">PREMI SIAP BASIS<br/><font size="1">(Rp/HK)</font></th>
		<th style="background-color:#00A300;color:#fff;">PREMI LEBIH BASIS<br/><font size="1">(Rp/Jjg)</font></th>
		<th style="background-color:#00A300;color:#fff;">BASIS LUAS HANCAK<br/><font size="1">(Ha)</font></th>
	</tr>
        </thead>
 <tbody>
<%
Set rsEPS = Server.CreateObject("ADODB.Recordset")
rsEPS.Open strEPS , connstr, adOpenStatic 

Dim xNum

if rsEPS.bof and rsEPS.eof then
%>
	<tr>
		<td colspan="<%=5+GetLastDay(v_txtMonth, v_txtYear)%>">NO DATA AVAILABLE</td>
	</tr>
<%end if
	 dim newcolum,tempharvestingsectionnewcolum,tempDivisionnewcolum
     newcolum = 0
		while not rsEPS.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
	<tr class="<%=cClass%>">
		

		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
		
			<%if tempGroupPlantation <> rsEPS("GroupPlantation") then %>	
				<td align="center" style=" border-bottom:none;"><%=rsEPS("GroupPlantation")%></td>
				<% tempGroupPlantation = rsEPS("GroupPlantation")%>
				
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
			
			<%end if %>
			
		<%End If%>

		
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
		
			<%if SprevPlantation <> rsEPS("Plantation") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_LaporanPremiPanen.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("Plantation")%>
							</a>
				</td>
				<% SprevPlantation = rsEPS("Plantation")%>
	
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
		
			<%end if %>
		
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
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
			<%if SprevGroupRegion <> rsEPS("groupregion") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_LaporanPremiPanen.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("groupregion")%>
							</a></td>
				<% SprevGroupRegion = rsEPS("groupregion")%>
			
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
		
			<%end if %>
			
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%> 
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
		
			<%if SprevRegion <> rsEPS("Region") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_LaporanPremiPanen.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("Region")%>
							</a>
				</td>
				<% SprevRegion = rsEPS("Region")%>
			
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
		
			<%end if %>
				
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then

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
					   
					   tempCompany = rsEPS("Company")
					   tempfcba = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   

		%>
			<%if SprevCompany <> rsEPS("company") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_LaporanPremiPanen.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("company")%>
							</a>
				</td>
				<% SprevCompany = rsEPS("company")%>

			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
	
			<%end if %>
		
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
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
			<%if Sprevfcba <> rsEPS("fcba") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_LaporanPremiPanen.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("fcba")%>
							</a>
				</td>
				<% Sprevfcba = rsEPS("fcba")%>
				<% newcolum = 1%>
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
				<% newcolum = 0%>
			<%end if %>
			
		<%End If%>
		
		
		<% if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then%>
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
					   tempBlok = "--ALL--"					   
					 %>
			<%if SprevDivision <> rsEPS("DIVISION") then %>	
				<td align="center" style=" border-bottom:none;">
				<a href="p_LaporanPremiPanen.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempCompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("DIVISION")%>
							</a></td>
				<% SprevDivision = rsEPS("DIVISION")%>
				<% tempDivisionnewcolum = 1%>
			<% elseif tempDivision = rsEPS("DIVISION") and newcolum = 1 then%>
				<td align="center" style=" border-bottom:none;"><%=rsEPS("DIVISION")%></td>
					<% SprevDivision = rsEPS("DIVISION")%>
					<% tempDivisionnewcolum = 1%>
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
				<% tempDivisionnewcolum = 0%>
			<%end if %>
			
		<%End if %>
		
			<%if Sprevharvestingsection <> rsEPS("harvestingsection") then %>	
				<td align="center" style=" border-bottom:none;"><%=rsEPS("harvestingsection")%></td>
				<% Sprevharvestingsection = rsEPS("harvestingsection")%>
				<% tempharvestingsectionnewcolum = 1%>
			<% elseif tempharvestingsection = rsEPS("harvestingsection") and tempDivisionnewcolum = 1 then%>
			<td align="center" style=" border-bottom:none;"><%=rsEPS("harvestingsection")%></td>
				<% Sprevharvestingsection = rsEPS("harvestingsection")%>
				<% tempharvestingsectionnewcolum = 1%>
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
				<% tempharvestingsectionnewcolum = 0%>
			<%end if %>
		
		<%if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then%>
		
			<%if SprevPlantingDate <> rsEPS("plantingyear") then %>	
				<td align="center" style=" border-bottom:none;"><%=rsEPS("plantingyear")%></td>
				<% SprevPlantingDate = rsEPS("plantingyear")%>
			<% elseif tempPlantingDate = rsEPS("plantingyear") and tempharvestingsectionnewcolum = 1 then%>
				<td align="center" style=" border-bottom:none;"><%=rsEPS("plantingyear")%></td>
				<% SprevPlantingDate = rsEPS("plantingyear")%>
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
			<%end if %>
			
		<%End if %>	
		
		
		<%if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then%>
			<%if SprevBlok <> rsEPS("FIELDCODE") then %>	
				<td align="center" style=" border-bottom:none;"><%=rsEPS("FIELDCODE")%></td>
				<td align="center" style=" border-bottom:none;"><%=rsEPS("fcname")%></td>
				<% SprevBlok = rsEPS("FIELDCODE")%>
			<% else	%>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
				<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
			<%end if %>

		<%End if %>

			<td align="center"><%=rsEPS("hectarageplanted")%></td>
			<td align="center"><%=rsEPS("harvestingbased_abw")%></td>
			<td align="center"><%=rsEPS("basis_qty")%></td>
			<td align="center"><%=rsEPS("Basis_Rate")%></td>
			<td align="center"><%=rsEPS("Extra_Basis_Rate")%></td>
			<td align="center"><%=rsEPS("Basis_Ha")%></td>
	</tr>
<%

			rsEPS.MoveNext
		Wend
		
rsEPS.Close
set rsEPS = Nothing

%>
</tbody>
</table>
</form>
</div>
</body>
</html>

