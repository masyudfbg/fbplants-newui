<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtDivision, v_txtPlantingDate, v_txtBlok
Dim v_txtyear,v_txtmonth,no
Dim txtDivision, txtBlok, txtPlantingDate
Dim v_vstatus,v_vCROPTYPE,v_vJenisReport, v_vOwnership
Dim tempyear,tempmonth
Dim tempDivision, tempBlok, tempPlantingDate 
Dim SprevCode,Sprevjobname, sprevdetail
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

Dim v_txttype : v_txttype = Request("txttype")
if len(trim(v_txttype)) = 0 then v_txttype = "rekap"

Dim currjob : currjob = ""
Dim prevjob : prevjob = ""
Dim currjobdesc : currjobdesc = ""
Dim prevjobdesc : prevjobdesc = ""
Dim sub_act_bi : sub_act_bi = 0
Dim sub_act_sbi : sub_act_sbi = 0
Dim sub_bud_bi : sub_bud_bi = 0
Dim sub_bud_sbi : sub_bud_sbi = 0
Dim sub_bud_ann : sub_bud_ann = 0

Dim sub_rpha_act_bi : sub_rpha_act_bi = 0
Dim sub_rpha_act_sbi : sub_rpha_act_sbi = 0
Dim sub_rpha_bud_bi : sub_rpha_bud_bi = 0
Dim sub_rpha_bud_sbi : sub_rpha_bud_sbi = 0
Dim sub_rpha_bud_ann : sub_rpha_bud_ann = 0

Dim sub_var_act_bi : sub_var_act_bi = 0
Dim sub_var_act_sbi : sub_var_act_sbi = 0

Dim addcolspan : addcolspan = 0

            
			 v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	  = Request("txtPlantation")
			v_txtGroupRegion  	  = Request("txtGroupRegion")
			v_txtRegion  		  = Request("txtRegion")
			v_txtcompany  		  = Request("txtCompany")
			v_txtfcba 			  = Request("txtfcba")
			v_txtDivision		  = Request("txtDivision")
			v_txtPlantingDate	  = Request("txtPlantingDate")
			v_txtBlok			  = Request("txtBlok")
			v_txtyear             = Request("txtYear")
			v_txtmonth            = Request("txtMonth")
			' v_vstatus             = Request("vStatus")
			' v_vCROPTYPE           = Request("vCROPTYPE")
			' v_vJenisReport        = request("vJenisReport")
			' v_vOwnership	      = Request("vOwnership")
			nSpanTotal = 5
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " TT" & v_txtPlantingDate
			'response.write  "YEAR " & v_txtyear & "  month " & v_txtmonth& "  status " & v_vstatus & " crop " & v_vCROPTYPE  & " owner " & v_vOwnership & " jenis " & v_vJenisReport  

			
		    Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut, businessunit_api.Isactive(fcba) fcbaActive, " 

						   
			wherecons = " where company is not null and fcbaActive='TRUE'  "
			
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
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
				addcolspan = addcolspan + 1
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
				addcolspan = addcolspan + 1
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
				addcolspan = addcolspan + 1
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
				addcolspan = addcolspan + 1
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
				addcolspan = addcolspan + 1
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
			 
%>
<html>

<head>
	<title>ePlaS Reporting - Analisa Biaya Tidak Langsung</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<link rel="stylesheet" href="/css/font/font-awesome.css"/>
	<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
	xAddEventListener(window, 'load',
	  function() {
		new xTableHeaderFixed('xthf-site', window);
	  }, false
	);
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
		
		<form name="frmBTL" method="post" action="<%=ScriptName%>">
			<div class="printhide">
				<!--#include file="FormSelectNewSum.html"-->
				<table id="form-table" class="centered-margin">	
					<tr>
						<td><font face="verdana" size="1">TAHUN</font></td>
						<td>
							<% 
								MM=v_txtMonth
								yyyy=v_txtYear

								mon = Month(Session("v_Costbook"))
								If len(MM)=0 then MM=mon

								Yr = Year(Session("v_Costbook"))
								MinYr = 2015
								MaxYr = Yr  + 2
								If len(yyyy)= 0 then yyyy = Yr
							%>
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
						<td><font face="verdana" size="1">BULAN</font></td>
						<td>
							<select name="txtMonth">
									<%
										For i = 1 to 12
									%>
											<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=Bulan(i)%></option>
									<%
										Next
									%>
							</select>
						</td>
						<td><font face="verdana" size="1">TYPE</font></td>
						<td>
							<select name="txttype">
								<Option Value="rekap" <%If v_txttype = "rekap" then%> Selected <%End If%>>Rekap</option>
								<Option Value="detail" <%If v_txttype = "detail" then%> Selected <%End If%>>Detail</option>
							</select>
						</td>
						<td><input type="submit" value="View" /></td>
					</tr>
				</table>	
			</div>
		</form>
		
	</div>
	
<%
Dim cur_month,strbud_BI,strbud_SBI,x1
	    strbud_BI = "btm.budget_01" 
	    strbud_SBI = "btm.budget_01"
		
		
		cur_month = v_txtmonth 
		
		IF Len(cur_month) > 0 then

			strbud_BI = "" 
	        strbud_SBI = ""
	
			x1=1
			
			Do While x1<= cint(cur_month) 'Output the values from 0 to 10
			  'response.write(i & "<br \>")
			  'response.write  " Budget_" & Lpad(x1,"0",2) & " +"
			  strbud_SBI = strbud_SBI & " btm.budget_" & Lpad(x1,"0",2) & " +"
			  x1 = x1 + 1 'increment the value of i for next time loop executes
			Loop

			strbud_BI = "btm.budget_" & Lpad(cur_month,"0",2)
	
			strbud_SBI = Mid(strbud_SBI,1,Len(strbud_SBI)-1)
		
			
            ' response.write  strbudTBS_BI & "<br>"
			' response.write strbudCPO_BI & "<br>"
			' response.write strbudKER_BI & "<br>"
			
			
			' response.write strbudTBS_SBI  & "<br>"
			' response.write strbudCPO_SBI  & "<br>"
			' response.write strbudKER_SBI  & "<br>"
		END IF	

Dim strha, rsBTL, LuasHa,budHA
' strha =  " SELECT GroupPlantation,SUM(ha) HA, SUM(bud_ha) bud_ha from ( " & _
         ' "  SELECT " & strFieldCons & " DECODE(SUM (ha), NULL,1,0,1,SUM (ha)) HA,DECODE(SUM (bud_ha), NULL,1,0,1,SUM (bud_ha)) bud_ha FROM   " & _
		 ' "  ( SELECT fcba, croptype, ownership, division, status, TO_CHAR(plantingdate,'YYYY') PLANTINGYEAR, SUM (hectarageplanted) ha,0 bud_ha FROM field   " & _ 
		 ' "	 GROUP BY fcba, croptype, ownership, division, status, plantingdate   " & _
		 ' "	 union all " & _
		 ' "	 select fcba,croptype,ownership,division,status,PLANTINGYEAR,0 ha, sum(hectarageplanted)  bud_ha  from    " & _
		 ' "	 ( select  distinct t.fcba,t.croptype,t.ownership,t.division,t.status,to_char(t.plantingdate,'YYYY') PLANTINGYEAR, hectarageplanted   " & _
		 ' "	   from BUDGET_TRANSACTION_ESTATE t )   " & _
		 ' "	 GROUP BY fcba, croptype, ownership, division, status, PLANTINGYEAR   " & _
		 ' "  ) group by fcba )" & whereCons & " group by GroupPlantation "
		 
		 
If v_txtYear <> "" and v_txtMonth <> "" Then
		 
	strha =  " SELECT GroupPlantation, nvl(SUM(nvl(ha,0)),0) HA, nvl(SUM(nvl(bud_ha,0)),0) bud_ha from ( " & _
			 "  SELECT " & strFieldCons & " SUM (nvl(ha,0)) HA,SUM (nvl(bud_ha,0)) bud_ha FROM   " & _
			 "  ( SELECT fcba, field_api.GetValid_CropType(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )  croptype,field_api.GetValid_ownership(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )  ownership, field_api.getValid_Afdeling(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) division, field_api.GetValid_status(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) status,field_api.getValid_Plantingyear(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) PLANTINGYEAR, SUM (nvl(field_api.getValid_HA(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ),0)) ha,0 bud_ha FROM field   " & _ 
			 "    where  field_api.GetValid_Activation(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) = 'Y' " & _
			 "	  GROUP BY fcba,field_api.GetValid_CropType(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) ,field_api.GetValid_ownership(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ), field_api.getValid_Afdeling(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) , field_api.GetValid_status(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) ,field_api.getValid_Plantingyear(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )    " & _
			 "	 union all " & _
			 "	 select fcba,croptype,ownership,division,status,PLANTINGYEAR,0 ha, sum(hectarageplanted)  bud_ha  from    " & _
			 "		( select  t.fcba,t.croptype,t.ownership,t.division,t.status,to_char(t.plantingdate,'YYYY') PLANTINGYEAR, hectarageplanted " & _
			 "		  from Budget_Field t where   t.budget_year = '" & v_txtyear & "') " & _
			 "	 GROUP BY fcba, croptype, ownership, division, status, PLANTINGYEAR   " & _
			 "  ) group by fcba )" & whereCons & " group by GroupPlantation "		 
			 
			 
	' response.write strha
			 
			
	Set rsBTL = Server.CreateObject("ADODB.Recordset")

	rsBTL.Open strha, connstr, adOpenStatic
	if Not rsBTL.eof then
			LuasHa	=	cdbl(rsBTL("ha"))
			budHA = cdbl(rsBTL("bud_ha"))
	end if
	rsBTL.Close
	Set rsBTL = Nothing
	
	
	%>
	
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
		<tr>
			<th colspan="22" style="background-color:#FFFFFF; border:none;">
				<div id="header-wrapper">
				<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
						<% If  CompanyName <>  "" then %>
							<h4>PT.<%=CompanyName%></h4> 
						<% End if%>
						<% If  FcbaName <> ""  then %>
							   <h5><%=FcbaName%></h5>   
						<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h3>LAPORAN BIAYA TAK LANGSUNG</h3>
					<h4>Periode&nbsp;:&nbsp;<%=Bulan(mm)%>&nbsp;<%=yyyy%></h4>
				</div>
				</div>
			</th>
		</tr>
		<tr>
			<%if v_txttype = "detail" then%>
				<th style="background-color:#00A300;color:#fff;"  colspan="<%=1 + 2 + nSpanTotal%>">&nbsp;</th>
			<%else%>
				<th style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
			<%end if%>
			<th style="background-color:#00A300;color:#fff;"  colspan="5">BULAN INI</th>
			<th style="background-color:#00A300;color:#fff;"  colspan="5">S/D BULAN INI</th>
			<th style="background-color:#00A300;color:#fff;"  colspan="2" rowspan="2">ANGGARAN TAHUNAN</th>
		</tr>
		<tr>
			<%if v_txttype = "detail" then%>
				<th style="background-color:#00A300;color:#fff;"  colspan="<%=1 + 2 + nSpanTotal%>">&nbsp;</th>
			<%else%>
				<th style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
			<%end if%>
			<th style="background-color:#00A300;color:#fff;"  colspan="2">REALISASI</th>
			<th style="background-color:#00A300;color:#fff;"  colspan="2">ANGGARAN</th>
			<th style="background-color:#00A300;color:#fff;"  rowspan="3">%</th>
			<th style="background-color:#00A300;color:#fff;"  colspan="2">REALISASI</th>
			<th style="background-color:#00A300;color:#fff;"  colspan="2">ANGGARAN</th>
			<th style="background-color:#00A300;color:#fff;"  rowspan="3">%</th>
		</tr>
		<tr>
			<%if v_txttype = "detail" then%>
				<th style="background-color:#00A300;color:#fff;"  colspan="<%= 1 + 2 + nSpanTotal %>" align="right">Luas Ha</th>
			<%else%>
				<th style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">Luas Ha</th>
			<%end if%>
			<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(LuasHa,0)%></th>
			<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(budHa,0)%></th>
			<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(LuasHa,0)%></th>
			<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(budHa,0)%></th>
			<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(budHa,0)%></th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;" >CODE</th>
			<th style="background-color:#00A300;color:#fff;" >JENIS PEKERJAAN</th>
			<%if v_txttype = "detail" then%>
				<th style="background-color:#00A300;color:#fff;" >&nbsp;</th>
			<%end if%>
			<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
				  <th style="background-color:#00A300;color:#fff;" >PLANTATION</th>			
			<% end if %>
			
			<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				  <th style="background-color:#00A300;color:#fff;" >GROUP REGION</th>			
			<% end if %>
			
			<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				  <th style="background-color:#00A300;color:#fff;" >REGION</th>		
			<% end if %>
			
			<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				   <th style="background-color:#00A300;color:#fff;" >COMPANY</th>		
			<% end if %>
			
			<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
				   <th style="background-color:#00A300;color:#fff;" >SITE</th>			
			<% end if %>

			
			<% if v_txtDivision <> "--SUMMARY--"  and v_txtDivision <> "" then %>
				<th style="background-color:#00A300;color:#fff;" >AFDELING</th>	
		   <% end if %>
			
			
			<% if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> "" then %>
				<th style="background-color:#00A300;color:#fff;" >TT</th>			
		   <% end if %>
		   
			<th style="background-color:#00A300;color:#fff;" >Total</th>
			<th style="background-color:#00A300;color:#fff;" >Rp/ha</th>
			<th style="background-color:#00A300;color:#fff;" >Total</th>
			<th style="background-color:#00A300;color:#fff;" >Rp/ha</th>
			<th style="background-color:#00A300;color:#fff;" >Total</th>
			<th style="background-color:#00A300;color:#fff;" >Rp/ha</th>
			<th style="background-color:#00A300;color:#fff;" >Total</th>
			<th style="background-color:#00A300;color:#fff;" >Rp/ha</th>
			<th style="background-color:#00A300;color:#fff;" >Total</th>
			<th style="background-color:#00A300;color:#fff;" >Rp/ha</th>
		</tr>
	</thead>
	<tbody>
	<%
	'Dim rsBTL, strGroupBy, rowno, rsalokasiBTL, stralokasiBTL
	Dim  rowno, strBTL
	Dim subBI,subSBI,totsubBI,totsubSBI,budsubBI,budsubSBI,budSubthn,budtotsubBI,budtotsubSBI,budtotsubTHN
	Dim rsjob,strjob,debetjob,kreditjob,notAlokasi

	'strjob= "select report_posting_api.getPostingDataDirect('KSR','ANALISABTL','p_AnalisaBTL.asp','DATADEBET_ESTATE','JOB') debet, report_posting_api.getPostingDataDirect('KSR','ANALISABTL','p_AnalisaBTL.asp','DATAKREDIT_ESTATE','JOB')   kredit  from dual"
	'Set rsjob = Server.CreateObject("ADODB.Recordset")
	'response.write strjob
	'rsjob.Open strjob, connstr, adOpenStatic
	'if Not rsjob.eof then
	'   debetjob = rsjob("debet")
	'   kreditjob = rsjob("kredit")
	'end if


	rowno		=	0
	if v_txttype = "detail" then
		strBTL		=	" select GroupPlantation " & selectCons & "" & _
						"        , grp" & _
						"        , jobcode" & _
						"        , description" & _
						"        , targettype" & _
						"        , targetcode" & _
						"        , detail" & _
						"        , ROUND(sum(bi), 0) BI" & _
						"        , ROUND(sum(sbi), 0) SBI" & _
						"        , ROUND(sum(budBI), 0) budBI" & _
						"        , ROUND(sum(budSBI), 0) budSBI" & _
						"        , ROUND(sum(budthn), 0) budthn" & _
						" from ( " & _
						" select " & strFieldCons & " grp" & _
						"        , jobcode" & _
						"        , description" & _
						"        , targettype" & _
						"        , targetcode" & _
						"        , decode(targettype, 'GC', targetcode||' - '||GCMASTER_API.get_fcname(fcba, targettype, targetcode)" & _
						"                           , 'BM', targetcode||' - '||BUILDING_API.get_FCNAME(targetcode, fcba)" & _
						"                           , null) detail" & _
						"        , sum(bi) BI" & _
						"        , sum(sbi) SBI" & _
						"        , sum(budBI) budBI" & _
						"        , sum(budSBI) budSBI" & _
						"        , sum(budthn) budthn" & _
						" from ( " & _
						"         SELECT data.*, jj.fcname description" & _
						"         FROM (" & _
						"             SELECT fcba, status grp, jobcode, targettype, targetcode, SUM(act_bi) bi, SUM(bud_bi) budbi, SUM(act_sbi) sbi, SUM(bud_sbi) budsbi, SUM(bud_annual) budthn" & _
						"             FROM ar_btl" & _
						"             WHERE year = '" & v_txtyear & "'" & _
						"                   AND month = '" & v_txtmonth & "'" & _
						"             GROUP BY fcba, status, jobcode, targettype, targetcode" & _
						"         ) data" & _
						"         LEFT JOIN job jj ON (data.fcba = jj.fcba AND data.jobcode = jj.fccode)" & _
						"	 ) " & _
						"	group by fcba" & _
						"            , grp" & _
						"            , jobcode" & _
						"            , description" & _
						"            , targettype" & _
						"            , targetcode" & _
						"            , decode(targettype, 'GC', targetcode||' - '||GCMASTER_API.get_fcname(fcba, targettype, targetcode)" & _
						"                               , 'BM', targetcode||' - '||BUILDING_API.get_FCNAME(targetcode, fcba)" & _
						"                               , null)" & _
						" )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by GroupPlantation,grp,jobcode,description, targettype, targetcode, detail" & selectCons & " order by grp,jobcode" & orderCons & ", detail"
					
	else
		strBTL		=	" select GroupPlantation " & selectCons & " , grp,jobcode,description, ROUND(sum(bi), 0) BI, ROUND(sum(sbi), 0) SBI, ROUND(sum(budBI), 0) budBI, ROUND(sum(budSBI), 0) budSBI, ROUND(sum(budthn), 0) budthn  from ( " & _
					" select " & strFieldCons & " grp,jobcode,description, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from ( " & _
					"     SELECT " & strFieldCons & " grp, jobcode, description, SUM(bi) bi, SUM(sbi) sbi, SUM(budBI) budbi, SUM(budSBI) budsbi, SUM(budthn) budthn" & _
					"     FROM ( " & _
					"         SELECT data.*, jj.fcname description" & _
					"         FROM (" & _
					"             SELECT fcba, status grp, jobcode, SUM(act_bi) bi, SUM(bud_bi) budbi, SUM(act_sbi) sbi, SUM(bud_sbi) budsbi, SUM(bud_annual) budthn" & _
					"             FROM ar_btl" & _
					"             WHERE year = '" & v_txtyear & "'" & _
					"                   AND month = '" & v_txtmonth & "'" & _
					"             GROUP BY fcba, status, jobcode" & _
					"         ) data" & _
					"         LEFT JOIN job jj ON (data.fcba = jj.fcba AND data.jobcode = jj.fccode)" & _
					"     ) " & _
					"     GROUP BY fcba, grp, jobcode, description" & _
					" ) " & _
					"	group by fcba,grp,jobcode,description )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by GroupPlantation,grp,jobcode,description " & selectCons & " order by grp,jobcode" & orderCons

	end if


	' response.write strBTL

	Set rsBTL = Server.CreateObject("ADODB.Recordset")
	rsBTL.Open strBTL, connstr, adOpenStatic
	
	if Not rsBTL.eof then

	Dim TotalBTL, TotalTyBTL
	Dim BudBI_BudHa, BudSBI_BudHa, BudTHN_BudHa
	Dim AmtBI_LuasHa, AmtSBI_LuasHa
    
    If isNull(budHA) Then 
	   budHA = 0
	end if
	If isNull(LuasHa) Then 
	   LuasHa = 0
	end if	
	
	TotalBTL = 0
	TotalTyBTL = 0
		While Not rsBTL.EOF			 
				rowno	=	rowno	+	1
				if Not rsBTL.EOF		then
					if Cdbl(budHa) <> 0 then
					    'If  rsBTL("grp") <> "A2" Then
							BudBI_BudHa   = Cdbl(rsBTL("budBI"))/Cdbl(budHa)
							BudSBI_BudHa  = Cdbl(rsBTL("budSBI"))/Cdbl(budHa) 
							BudTHN_BudHa  = Cdbl(rsBTL("budthn"))/Cdbl(budHa)
						'end if
					else
						BudBI_BudHa   = 0
						BudSBI_BudHa  = 0
						BudTHN_BudHa  = 0
					end if
				end if
				
				if Cdbl(LuasHa) <> 0 then
				   AmtBI_LuasHa  = Cdbl(rsBTL("BI"))/Cdbl(LuasHa)
				   AmtSBI_LuasHa = Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)
				else
				   AmtBI_LuasHa  = 0
				   AmtSBI_LuasHa = 0
				end if
				
			If 	rsBTL("grp") <> "A2" Then
				if v_txttype = "detail" then
					currjob = rsBTL("JOBCODE")
					currjobdesc = rsBTL("description")
					if currjob <> prevjob then
						if len(trim(prevjob)) = 0 then
							sub_act_bi = 0
							sub_act_sbi = 0
							sub_bud_bi = 0
							sub_bud_sbi = 0
							sub_bud_ann = 0
							sub_rpha_act_bi = 0
							sub_rpha_act_sbi = 0
							sub_rpha_bud_bi = 0
							sub_rpha_bud_sbi = 0
							sub_rpha_bud_ann = 0
							sub_var_act_bi = 0
							sub_var_act_sbi = 0
						else
							sub_rpha_act_bi = 0
							sub_rpha_act_sbi = 0
							sub_rpha_bud_bi = 0
							sub_rpha_bud_sbi = 0
							sub_rpha_bud_ann = 0
							sub_var_act_bi = 0
							sub_var_act_sbi = 0
							if LuasHa <> 0 then
								sub_rpha_act_bi = sub_act_bi/LuasHa
								sub_rpha_act_sbi = sub_act_sbi/LuasHa
							else
								sub_rpha_act_bi = 0
								sub_rpha_act_sbi = 0
							end if
							if budHa <> 0 then
								sub_rpha_bud_bi = sub_bud_bi/budHa
								sub_rpha_bud_sbi = sub_bud_sbi/budHa
								sub_rpha_bud_ann = sub_bud_ann/budHa
							else
								sub_rpha_bud_bi = 0
								sub_rpha_bud_sbi = 0
								sub_rpha_bud_ann = 0
							end if
							
							if sub_rpha_bud_bi = 0 and sub_rpha_act_bi <> 0 then
								sub_var_act_bi = -100
							elseif sub_rpha_bud_bi <> 0 and sub_rpha_act_bi = 0 then
								sub_var_act_bi = 100
							elseif sub_rpha_bud_bi = 0 and sub_rpha_act_bi = 0 then
								sub_var_act_bi = 0
							elseif sub_rpha_bud_bi <> 0 and sub_rpha_act_bi <> 0 then
								sub_var_act_bi = Round(((sub_rpha_bud_bi - sub_rpha_act_bi)/ sub_rpha_bud_bi)*100, 2)
							else
								sub_var_act_bi = 0
							end if
							
							if sub_rpha_bud_sbi = 0 and sub_rpha_act_sbi <> 0 then
								sub_var_act_sbi = -100
							elseif sub_rpha_bud_sbi <> 0 and sub_rpha_act_sbi = 0 then
								sub_var_act_sbi = 100
							elseif sub_rpha_bud_sbi = 0 and sub_rpha_act_sbi = 0 then
								sub_var_act_sbi = 0
							elseif sub_rpha_bud_sbi <> 0 and sub_rpha_act_sbi <> 0 then
								sub_var_act_sbi = Round(((sub_rpha_bud_sbi - sub_rpha_act_sbi)/ sub_rpha_bud_sbi)*100, 2)
							else
								sub_var_act_sbi = 0
							end if
			%>
							<tr bgcolor="#F4F4A0">
								<td colspan="<%=3 + addcolspan%>"><strong>Subtotal <%=prevjobdesc%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_act_bi, 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpha_act_bi, 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_bud_bi, 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpha_bud_bi, 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_var_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_act_sbi, 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpha_act_sbi, 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_bud_sbi, 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpha_bud_sbi, 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_var_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_bud_ann, 0)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpha_bud_ann, 0)%></strong></td>
							</tr>
			<%
							sub_act_bi = 0
							sub_act_sbi = 0
							sub_bud_bi = 0
							sub_bud_sbi = 0
							sub_bud_ann = 0
							sub_rpha_act_bi = 0
							sub_rpha_act_sbi = 0
							sub_rpha_bud_bi = 0
							sub_rpha_bud_sbi = 0
							sub_rpha_bud_ann = 0
							sub_var_act_bi = 0
							sub_var_act_sbi = 0
						end if
						sub_act_bi = sub_act_bi + cdbl(rsBTL("bi"))
						sub_act_sbi = sub_act_sbi + cdbl(rsBTL("sbi"))
						sub_bud_bi = sub_bud_bi + cdbl(rsBTL("budBI"))
						sub_bud_sbi = sub_bud_sbi + cdbl(rsBTL("budSBI"))
						sub_bud_ann = sub_bud_ann + cdbl(rsBTL("budthn"))
					else
						sub_act_bi = sub_act_bi + cdbl(rsBTL("bi"))
						sub_act_sbi = sub_act_sbi + cdbl(rsBTL("sbi"))
						sub_bud_bi = sub_bud_bi + cdbl(rsBTL("budBI"))
						sub_bud_sbi = sub_bud_sbi + cdbl(rsBTL("budSBI"))
						sub_bud_ann = sub_bud_ann + cdbl(rsBTL("budthn"))
					end if
				end if
			%>
				<tr>
					
					
							   <% If SprevCode <> rsBTL("JOBCODE") then %>
									 <td align="center"><%=rsBTL("JOBCODE")%></td>
							   <% Else %>
									 <td align="center"></td>
							   <% End If %>				   
								<% If Sprevjobname <> rsBTL("description") then %>
									 <td><%=rsBTL("description")%></td>
							   <% Else %>
									 <td align="center"></td>
							   <% End If %>
							   <%if v_txttype = "detail" then%>
									<td><%=rsBTL("detail")%></td>
							   <% End If %>
							   <%
							   SprevCode   = rsBTL("JOBCODE")
							   Sprevjobname = rsBTL("description")
							   if v_txttype = "detail" then sprevdetail = rsBTL("detail")
							   notAlokasi = "T"
							   %>
							   <% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										 tempPlantation = rsBTL("Plantation")
										 tempGroupRegion = "--ALL--"
										 tempRegion = "--SUMMARY--"
										 tempCompany ="--SUMMARY--"
										 tempfcba = "--SUMMARY--"
										 tempDivision = "--SUMMARY--"
										 tempPlantingDate = "--SUMMARY--"
										 IF SprevPlantation <> rsBTL("Plantation") then %>
											   <td align="center">																																																																																									
													<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsBTL("Plantation")%>
													</a>
											   </td>
											   <% SprevPlantation = rsBTL("Plantation") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
								<% end if %>
								
								<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
										 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
											  tempPlantation = rsBTL("Plantation")
										 else
											  tempPlantation = ""
										 end if
										 tempGroupRegion = rsBTL("GroupRegion")
										 tempRegion = "--ALL--"
										 tempCompany ="--SUMMARY--"
										 tempfcba = "--SUMMARY--"
										 tempDivision = "--SUMMARY--"
										 tempPlantingDate = "--SUMMARY--"
										 IF SprevGroupRegion <> rsBTL("GroupRegion") then %>
											   <td align="center">
													<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsBTL("GroupRegion")%>
													</a>
											   </td>
											   <% SprevGroupRegion = rsBTL("GroupRegion") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
								<% end if %>
								
								<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
										   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
											  tempPlantation = rsBTL("Plantation")
										   else
											  tempPlantation = ""
										   end if
										   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
											   tempGroupRegion = rsBTL("GroupRegion")
										   else
											   tempGroupRegion = ""
										   end if
										   
										   tempRegion = rsBTL("Region") 
										   tempCompany ="--ALL--"
										   tempfcba = "--SUMMARY--"
										   tempDivision = "--SUMMARY--"
										   tempPlantingDate = "--SUMMARY--"
										   tempBlok = "--SUMMARY--"
											 IF SprevRegion <> rsBTL("Region") then %>
												   <td align="center">
														<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
														<%=rsBTL("Region")%>
														</a>
												   </td>
												   <% SprevRegion = rsBTL("Region") %>	
										  <% else	%>
												   <td align="center"></td>					
										   <% end if %>									   
								<% end if %>
								
								<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
										   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
											  tempPlantation = rsBTL("Plantation")
										   else
											  tempPlantation = ""
										   end if
										   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
											   tempGroupRegion = rsBTL("GroupRegion")
										   else
											   tempGroupRegion = ""
										   end if
										   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
											   tempRegion = rsBTL("Region") 
										   else
											   tempRegion = ""
										   end if								   
										   tempCompany =rsBTL("Company") 
										   tempfcba = "--ALL--"
										   tempDivision = "--SUMMARY--"
										   tempPlantingDate = "--SUMMARY--"
										   tempBlok = "--SUMMARY--"
											 IF SprevCompany <> rsBTL("Company") then %>
												   <td align="center">
														<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
														  <%=rsBTL("Company")%>
														</a>
												   </td>
												   <% SprevCompany = rsBTL("Company") %>	
										  <% else	%>
												   <td align="center"></td>					
										   <% end if %>									   
								<% end if %>
								
								<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
										   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
											  tempPlantation = rsBTL("Plantation")
										   else
											  tempPlantation = ""
										   end if
										   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
											   tempGroupRegion = rsBTL("GroupRegion")
										   else
											   tempGroupRegion = ""
										   end if
										   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
											   tempRegion = rsBTL("Region") 
										   else
											   tempRegion = ""
										   end if
										   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
											   tempCompany =rsBTL("Company")
										   else
											   tempCompany =""
										   end if							   
										   tempfcba = rsBTL("fcba")
										   tempDivision = "--ALL--"
										   tempPlantingDate = "--SUMMARY--"
										   tempBlok = "--SUMMARY--"
											 IF Sprevfcba <> rsBTL("fcba") then %>
												   <td align="center">
														<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
														<%=rsBTL("fcba")%>
														</a>
												   </td>
												   <% Sprevfcba = rsBTL("fcba") %>	
										  <% else	%>
												   <td align="center"></td>					
										   <% end if %>									   
								<% end if %>
								
								<% if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> ""  then
										   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
											  tempPlantation = rsBTL("Plantation")
										   else
											  tempPlantation = ""
										   end if
										   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
											   tempGroupRegion = rsBTL("GroupRegion")
										   else
											   tempGroupRegion = ""
										   end if
										   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
											   tempRegion = rsBTL("Region") 
										   else
											   tempRegion = ""
										   end if
										   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
											   tempCompany =rsBTL("Company")
										   else
											   tempCompany =""
										   end if
										   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
											   tempfcba = rsBTL("fcba")
										   else
											   tempfcba = ""
										   end if
										   tempDivision = rsBTL("Division")
										   tempPlantingDate = "--ALL--"
										   tempBlok = "--SUMMARY--"
											 IF SprevDivision <> rsBTL("Division") then %>
												   <td align="center">
														<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
														<%=rsBTL("Division")%>
														</a>
												   </td>
												   <% SprevDivision = rsBTL("Division") %>	
										  <% else	%>
												   <td align="center"></td>					
										   <% end if %>									   
								<% end if %>
								
													
								<% if v_txtPlantingDate <> "--SUMMARY--" AND v_txtPlantingDate <> "" then %>
									<td><%=rsBTL("PlantingYear")%></td>			
							   <% end if %>
					
					
					<% ' BI 
						 
					
					%>
					<td align="right"><%=FormatNumber(rsBTL("BI"),0)%></td>
					<% IF Cdbl(LuasHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("BI"))/Cdbl(LuasHa)),0)%></td>
					<% END IF %>
					
					<td align="right"><%=FormatNumber(rsBTL("budBI"),0)%></td>
					<% IF Cdbl(budHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budBI"))/Cdbl(budHa)),0)%></td>
					<% END IF %>
					
					
					
					<% IF cdbl(BudBI_BudHa) = 0 and Cdbl(AmtBI_LuasHa) <> 0 then %>
					  <td align="right">-100</td>
					<% ELSEIF cdbl(BudBI_BudHa) <> 0 and Cdbl(AmtBI_LuasHa) = 0 then %>
					  <td align="right">100</td>
					<% ELSEIF cdbl(BudBI_BudHa) = 0 and Cdbl(AmtBI_LuasHa) = 0 then %>
					  <td align="right">0</td>
					<% ELSEIF cdbl(BudBI_BudHa)  <> 0 and Cdbl(AmtBI_LuasHa) <> 0 then %>
					   <td align="right"><%=FormatNumber(((cdbl(BudBI_BudHa) - Cdbl(AmtBI_LuasHa))/cdbl(BudBI_BudHa))*100,2)%></td>
					<% END IF %>
					
					<% 
					' SBI 
					%>
					<td align="right"><%=FormatNumber(rsBTL("SBI"),0)%></td>
					<% IF Cdbl(LuasHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)),0)%></td>
					<% END IF %>
					
					<td align="right"><%=FormatNumber(rsBTL("budSBI"),0)%></td>
					<% IF Cdbl(budHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budSBI"))/Cdbl(budHa)),0)%></td>
					<% END IF %>
					
					<% IF cdbl(BudSBI_BudHa) = 0 and Cdbl(AmtSBI_LuasHa) <> 0 then %>
					  <td align="right">-100</td>
					<% ELSEIF cdbl(BudSBI_BudHa) <> 0 and Cdbl(AmtSBI_LuasHa) = 0 then %>
					  <td align="right">100</td>
					<% ELSEIF cdbl(BudSBI_BudHa) = 0 and Cdbl(AmtSBI_LuasHa) = 0 then %>
					  <td align="right">0</td>
					<% ELSEIF cdbl(BudSBI_BudHa)  <> 0 and Cdbl(AmtSBI_LuasHa) <> 0 then %>
					   <td align="right"><%=FormatNumber(((cdbl(BudSBI_BudHa) - Cdbl(AmtSBI_LuasHa))/cdbl(BudSBI_BudHa))*100,2)%></td>
					<% END IF %>
					
					<td align="right"><%=FormatNumber(rsBTL("budthn"),0)%></td>
					<% IF Cdbl(budHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budthn"))/Cdbl(budHa)),0)%></td>
					<% END IF %>
					
				</tr>
			<%
					subBI     = subBI + Cdbl(rsBTL("BI"))
					subSBI    = subSBI + Cdbl(rsBTL("SBI"))
					budsubBI  = budsubBI + Cdbl(rsBTL("budBI"))
					budsubSBI = budsubSBI + Cdbl(rsBTL("budSBI"))
					budSubthn = budSubthn + Cdbl(rsBTL("budTHN")) 
			Else 
					  if notAlokasi = "T"   then   %>
					  
						<%
							if v_txttype="detail" then
								sub_rpha_act_bi = 0
								sub_rpha_act_sbi = 0
								sub_rpha_bud_bi = 0
								sub_rpha_bud_sbi = 0
								sub_rpha_bud_ann = 0
								sub_var_act_bi = 0
								sub_var_act_sbi = 0
								if LuasHa <> 0 then
									sub_rpha_act_bi = sub_act_bi/LuasHa
									sub_rpha_act_sbi = sub_act_sbi/LuasHa
								else
									sub_rpha_act_bi = 0
									sub_rpha_act_sbi = 0
								end if
								if budHa <> 0 then
									sub_rpha_bud_bi = sub_bud_bi/budHa
									sub_rpha_bud_sbi = sub_bud_sbi/budHa
									sub_rpha_bud_ann = sub_bud_ann/budHa
								else
									sub_rpha_bud_bi = 0
									sub_rpha_bud_sbi = 0
									sub_rpha_bud_ann = 0
								end if
								
								if sub_rpha_bud_bi = 0 and sub_rpha_act_bi <> 0 then
									sub_var_act_bi = -100
								elseif sub_rpha_bud_bi <> 0 and sub_rpha_act_bi = 0 then
									sub_var_act_bi = 100
								elseif sub_rpha_bud_bi = 0 and sub_rpha_act_bi = 0 then
									sub_var_act_bi = 0
								elseif sub_rpha_bud_bi <> 0 and sub_rpha_act_bi <> 0 then
									sub_var_act_bi = Round(((sub_rpha_bud_bi - sub_rpha_act_bi)/ sub_rpha_bud_bi)*100, 2)
								else
									sub_var_act_bi = 0
								end if
								
								if sub_rpha_bud_sbi = 0 and sub_rpha_act_sbi <> 0 then
									sub_var_act_sbi = -100
								elseif sub_rpha_bud_sbi <> 0 and sub_rpha_act_sbi = 0 then
									sub_var_act_sbi = 100
								elseif sub_rpha_bud_sbi = 0 and sub_rpha_act_sbi = 0 then
									sub_var_act_sbi = 0
								elseif sub_rpha_bud_sbi <> 0 and sub_rpha_act_sbi <> 0 then
									sub_var_act_sbi = Round(((sub_rpha_bud_sbi - sub_rpha_act_sbi)/ sub_rpha_bud_sbi)*100, 2)
								else
									sub_var_act_sbi = 0
							    end if
					  %>
								<tr bgcolor="#F4F4A0">
									<td colspan="<%=3 + addcolspan%>"><strong>Subtotal <%=prevjobdesc%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_act_bi, 0)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_rpha_act_bi, 0)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_bud_bi, 0)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_rpha_bud_bi, 0)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_var_act_bi, 2)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_act_sbi, 0)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_rpha_act_sbi, 0)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_bud_sbi, 0)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_rpha_bud_sbi, 0)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_var_act_sbi, 2)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_bud_ann, 0)%></strong></td>
									<td align="right"><strong><%=FormatNumber(sub_rpha_bud_ann, 0)%></strong></td>
								</tr>
							<%end if%>
							<tr bgcolor="#ffff01">
								<%if v_txttype = "detail" then%>
									<td colspan="<%= 1 + 2 + nSpanTotal %>" align="left"><strong>Total Biaya Tak Langsung</strong></td>
								<%else%>
									<td colspan="<%= 2 + nSpanTotal %>" align="left"><strong>Total Biaya Tak Langsung</strong></td>
								<%end if%>
								
								<td align="right"><strong><%=FormatNumber(subBI,0)%></strong></td>
								<% IF Cdbl(budHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>						
								   <td align="right"><strong><%=FormatNumber((Cdbl(subBI)/Cdbl(LuasHa)),0)%></strong></td>	
								<% END IF %>   
								<td align="right"><strong><%=FormatNumber(budsubBI,0)%></strong></td>
								<% IF Cdbl(budHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>	
								   <td align="right"><strong><%=FormatNumber(cdbl(budsubBI)/cdbl(budHa),0)%></strong></td>
								<% END IF %>
								
								<% IF cdbl(budsubBI) = 0 and Cdbl(subBI) <> 0 then %>
									  <td align="right"><strong>-100</strong></td>
								<% ELSEIF cdbl(budsubBI) <> 0 and Cdbl(subBI) = 0 then %>	  
									  <td align="right"><strong>100</strong></td>
								<% ELSEIF cdbl(budsubBI) = 0 and Cdbl(subBI) = 0 then %>
									  <td align="right"><strong>0</strong></td>
								<% ELSEIF cdbl(budsubBI) <> 0 and Cdbl(subBI) <> 0 then %>
									  <td align="right"><strong><%=FormatNumber((cdbl(budsubBI) - Cdbl(subBI)) / cdbl(budsubBI) * 100,2)%></strong></td>
								<% END IF %>
								<td align="right"><strong><%=FormatNumber(subSBI,0)%></strong></td>	
								<% IF Cdbl(LuasHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>	
								   <td align="right"><strong><%=FormatNumber((Cdbl(subSBI)/Cdbl(LuasHa)),0)%></strong></td>	
								<% END IF %>   
								<td align="right"><strong><%=FormatNumber(budsubSBI,0)%></strong></td>
								<% IF Cdbl(budHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>	
								   <td align="right"><strong><%=FormatNumber(cdbl(budsubSBI)/cdbl(budHa),0)%></strong></td>
								<% END IF %>
								<% IF cdbl(budsubSBI) = 0 and Cdbl(subSBI) <> 0 then %>
									  <td align="right"><strong>-100</strong></td>
								<% ELSEIF cdbl(budsubSBI) <> 0 and Cdbl(subSBI) = 0 then %>	  
									  <td align="right"><strong>100</strong></td>
								<% ELSEIF cdbl(budsubSBI) = 0 and Cdbl(subSBI) = 0 then %>
									  <td align="right"><strong>0</strong></td>
								<% ELSEIF cdbl(budsubSBI) <> 0 and Cdbl(subSBI) <> 0 then %>
									  <td align="right"><strong><%=FormatNumber((cdbl(budsubSBI) - Cdbl(subSBI)) / cdbl(budsubSBI) * 100,2)%></strong></td>
								<% END IF %>
								<td align="right"><strong><%=FormatNumber(budsubTHN,0)%></strong></td>
								<% IF Cdbl(budHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>	
								   <td align="right"><strong><%=FormatNumber(cdbl(budsubTHN)/cdbl(budHa),0)%></strong></td>
								<% END IF %>
							</tr>
					 <%             notAlokasi = "F"
									subBI      = 0
									subSBI     = 0 
									budsubBI   = 0 
									budsubSBI  = 0 
									budSubthn  = 0  
						End if 
						
						subBI     = subBI + Cdbl(rsBTL("BI"))
						subSBI    = subSBI + Cdbl(rsBTL("SBI"))
						budsubBI  = budsubBI + Cdbl(rsBTL("budBI"))
						budsubSBI = budsubSBI + Cdbl(rsBTL("budSBI"))
						budSubthn = budSubthn + Cdbl(rsBTL("budTHN"))
						 'If  rsBTL("grp") <> "A2" Then
							'budsubBI  = budsubBI + Cdbl(rsBTL("budBI"))
							'budsubSBI = budsubSBI + Cdbl(rsBTL("budSBI"))
							'budSubthn = budSubthn + Cdbl(rsBTL("budTHN")) 
						 'end if 	
					 
					%>
					<tr>
							   <% If SprevCode <> rsBTL("JOBCODE") then %>
									 <td align="center"><%=rsBTL("JOBCODE")%></td>
							   <% Else %>
									 <td align="center"></td>
							   <% End If %>				   
								<% If Sprevjobname <> rsBTL("description") then %>
									 <td><%=rsBTL("description")%></td>
							   <% Else %>
									 <td align="center"></td>
							   <% End If %>
							   <%if v_txttype = "detail" then%>
									<td><%=rsBTL("detail")%></td>
							   <% End If %>
							   
							   <%
							   SprevCode   = rsBTL("JOBCODE")
							   Sprevjobname = rsBTL("description")
							   if v_txttype = "detail" then sprevdetail = rsBTL("detail")
							   %>
							   <% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										 tempPlantation = rsBTL("Plantation")
										 tempGroupRegion = "--ALL--"
										 tempRegion = "--SUMMARY--"
										 tempCompany ="--SUMMARY--"
										 tempfcba = "--SUMMARY--"
										 tempDivision = "--SUMMARY--"
										 tempPlantingDate = "--SUMMARY--"
										 IF SprevPlantation <> rsBTL("Plantation") then %>
											   <td align="center">																																																																																									
													<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsBTL("Plantation")%>
													</a>
											   </td>
											   <% SprevPlantation = rsBTL("Plantation") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
								<% end if %>
								
								<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
										 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
											  tempPlantation = rsBTL("Plantation")
										 else
											  tempPlantation = ""
										 end if
										 tempGroupRegion = rsBTL("GroupRegion")
										 tempRegion = "--ALL--"
										 tempCompany ="--SUMMARY--"
										 tempfcba = "--SUMMARY--"
										 tempDivision = "--SUMMARY--"
										 tempPlantingDate = "--SUMMARY--"
										 IF SprevGroupRegion <> rsBTL("GroupRegion") then %>
											   <td align="center">
													<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsBTL("GroupRegion")%>
													</a>
											   </td>
											   <% SprevGroupRegion = rsBTL("GroupRegion") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
								<% end if %>
								
								<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
										   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
											  tempPlantation = rsBTL("Plantation")
										   else
											  tempPlantation = ""
										   end if
										   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
											   tempGroupRegion = rsBTL("GroupRegion")
										   else
											   tempGroupRegion = ""
										   end if
										   
										   tempRegion = rsBTL("Region") 
										   tempCompany ="--ALL--"
										   tempfcba = "--SUMMARY--"
										   tempDivision = "--SUMMARY--"
										   tempPlantingDate = "--SUMMARY--"
										   tempBlok = "--SUMMARY--"
											 IF SprevRegion <> rsBTL("Region") then %>
												   <td align="center">
														<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
														<%=rsBTL("Region")%>
														</a>
												   </td>
												   <% SprevRegion = rsBTL("Region") %>	
										  <% else	%>
												   <td align="center"></td>					
										   <% end if %>									   
								<% end if %>
								
								<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
										   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
											  tempPlantation = rsBTL("Plantation")
										   else
											  tempPlantation = ""
										   end if
										   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
											   tempGroupRegion = rsBTL("GroupRegion")
										   else
											   tempGroupRegion = ""
										   end if
										   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
											   tempRegion = rsBTL("Region") 
										   else
											   tempRegion = ""
										   end if								   
										   tempCompany =rsBTL("Company") 
										   tempfcba = "--ALL--"
										   tempDivision = "--SUMMARY--"
										   tempPlantingDate = "--SUMMARY--"
										   tempBlok = "--SUMMARY--"
											 IF SprevCompany <> rsBTL("Company") then %>
												   <td align="center">
														<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
														  <%=rsBTL("Company")%>
														</a>
												   </td>
												   <% SprevCompany = rsBTL("Company") %>	
										  <% else	%>
												   <td align="center"></td>					
										   <% end if %>									   
								<% end if %>
								
								<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
										   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
											  tempPlantation = rsBTL("Plantation")
										   else
											  tempPlantation = ""
										   end if
										   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
											   tempGroupRegion = rsBTL("GroupRegion")
										   else
											   tempGroupRegion = ""
										   end if
										   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
											   tempRegion = rsBTL("Region") 
										   else
											   tempRegion = ""
										   end if
										   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
											   tempCompany =rsBTL("Company")
										   else
											   tempCompany =""
										   end if							   
										   tempfcba = rsBTL("fcba")
										   tempDivision = "--ALL--"
										   tempPlantingDate = "--SUMMARY--"
										   tempBlok = "--SUMMARY--"
											 IF Sprevfcba <> rsBTL("fcba") then %>
												   <td align="center">
														<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
														<%=rsBTL("fcba")%>
														</a>
												   </td>
												   <% Sprevfcba = rsBTL("fcba") %>	
										  <% else	%>
												   <td align="center"></td>					
										   <% end if %>									   
								<% end if %>
								
								<% if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> ""  then
										   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
											  tempPlantation = rsBTL("Plantation")
										   else
											  tempPlantation = ""
										   end if
										   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
											   tempGroupRegion = rsBTL("GroupRegion")
										   else
											   tempGroupRegion = ""
										   end if
										   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
											   tempRegion = rsBTL("Region") 
										   else
											   tempRegion = ""
										   end if
										   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
											   tempCompany =rsBTL("Company")
										   else
											   tempCompany =""
										   end if
										   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
											   tempfcba = rsBTL("fcba")
										   else
											   tempfcba = ""
										   end if
										   tempDivision = rsBTL("Division")
										   tempPlantingDate = "--ALL--"
										   tempBlok = "--SUMMARY--"
											 IF SprevDivision <> rsBTL("Division") then %>
												   <td align="center">
														<a href="p_AnalisaBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
														<%=rsBTL("Division")%>
														</a>
												   </td>
												   <% SprevDivision = rsBTL("Division") %>	
										  <% else	%>
												   <td align="center"></td>					
										   <% end if %>									   
								<% end if %>
								
													
								<% if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <>  "" then %>
									<td><%=rsBTL("PlantingYear")%></td>			
							   <% end if %>
						
						
						<% 
						' BI 
						%>
						<td align="right"><%=FormatNumber(rsBTL("BI"),0)%></td>
						<% IF Cdbl(LuasHa) = 0  then  %>
							<td align="right">0</td>
						<% ELSE %>		
						   <td align="right"><%=FormatNumber((Cdbl(rsBTL("BI"))/Cdbl(LuasHa)),0)%></td>
						<% END IF 

						%>
						
						<td align="right"><%=FormatNumber(rsBTL("BUDBI"), 0)%></td>
						<% IF Cdbl(budHa) = 0  then  %>
							<td align="right">0</td>
						<% ELSE %>		
						   <td align="right"><%=FormatNumber((Cdbl(rsBTL("BUDBI"))/Cdbl(budHa)),0)%></td>
						<% END IF %>
						
						<% IF cdbl(BudBI_BudHa) = 0 and Cdbl(AmtBI_LuasHa) <> 0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(BudBI_BudHa) <> 0 and Cdbl(AmtBI_LuasHa) = 0 then %>
						  <td align="right">100</td>
						<% ELSEIF cdbl(BudBI_BudHa) = 0 and Cdbl(AmtBI_LuasHa) = 0 then %>
						  <td align="right">0</td>
						<% ELSEIF cdbl(BudBI_BudHa)  <> 0 and Cdbl(AmtBI_LuasHa) <> 0 then %>
						   <td align="right"><%=FormatNumber(((cdbl(BudBI_BudHa) - Cdbl(AmtBI_LuasHa))/cdbl(BudBI_BudHa))*100,2)%></td>
						<% END IF %>
						
						<% 
						' SBI 
						%>
						<td align="right"><%=FormatNumber(rsBTL("SBI"),0)%></td>
						<% IF Cdbl(LuasHa) = 0  then  %>
							<td align="right">0</td>
						<% ELSE %>		
						   <td align="right"><%=FormatNumber((Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)),0)%></td>
						<% END IF %>
						
						<td align="right"><%=FormatNumber(rsBTL("budSBI"),0)%></td>
						<% IF Cdbl(budHa) = 0  then  %>
							<td align="right">0</td>
						<% ELSE %>		
						   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budSBI"))/Cdbl(budHa)),0)%></td>
						<% END IF %>
						
						<% IF cdbl(BudSBI_BudHa) = 0 and Cdbl(AmtSBI_LuasHa) <> 0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(BudSBI_BudHa) <> 0 and Cdbl(AmtSBI_LuasHa) = 0 then %>
						  <td align="right">100</td>
						<% ELSEIF cdbl(BudSBI_BudHa) = 0 and Cdbl(AmtSBI_LuasHa) = 0 then %>
						  <td align="right">0</td>
						<% ELSEIF cdbl(BudSBI_BudHa)  <> 0 and Cdbl(AmtSBI_LuasHa) <> 0 then %>
						   <td align="right"><%=FormatNumber(((cdbl(BudSBI_BudHa) - Cdbl(AmtSBI_LuasHa))/cdbl(BudSBI_BudHa))*100,2)%></td>
						<% END IF %>
						
						<td align="right"><%=FormatNumber(Cdbl(rsBTL("budthn")), 0)%></td>
						<% IF Cdbl(budHa) = 0  then  %>
							<td align="right">0</td>
						<% ELSE %>		
						   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budthn"))/Cdbl(budHa)),0)%></td>
						<% END IF %>
						
					</tr>
					<%
					   
				
			End If	
		
			
			IF Not rsBTL.EOF  then 
					totsubBI  = totsubBI + Cdbl(rsBTL("BI"))
					totsubSBI = totsubSBI + Cdbl(rsBTL("SBI"))
						
					budtotsubBI = budtotsubBI + Cdbl(rsBTL("budBI"))
					budtotsubSBI = budtotsubSBI + Cdbl(rsBTL("budSBI"))
					budtotsubTHN = budtotsubTHN + Cdbl(rsBTL("budTHN"))
			END IF
			
			if v_txttype = "detail" then
				prevjob = currjob
				prevjobdesc = currjobdesc
			end if
			
			rsBTL.MoveNext
		Wend
	end if
		%>
		
					  <tr bgcolor="#ffff01">
						<%if v_txttype = "detail" then%>
							<td colspan="<%= 1 + 2 + nSpanTotal %>" align="left"><strong>Total Alokasi Biaya Tak Langsung</strong></td>
						<%else%>
							<td colspan="<%= 2 + nSpanTotal %>" align="left"><strong>Total Alokasi Biaya Tak Langsung</strong></td>
						<%end if%>
								
								<td align="right"><strong><%=FormatNumber(subBI,0)%></strong></td>
								<% IF Cdbl(budHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>						
								   <td align="right"><strong><%=FormatNumber((Cdbl(subBI)/Cdbl(LuasHa)),0)%></strong></td>	
								<% END IF %>   
								<td align="right"><strong><%=FormatNumber(budsubBI,0)%></strong></td>
								<% IF Cdbl(budHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>	
								   <td align="right"><strong><%=FormatNumber(cdbl(budsubBI)/cdbl(budHa),0)%></strong></td>
								<% END IF %>
								<% IF cdbl(budsubBI) = 0 and Cdbl(subBI) <> 0 then %>
									  <td align="right"><strong>-100</strong></td>
								<% ELSEIF cdbl(budsubBI) <> 0 and Cdbl(subBI) = 0 then %>	  
									  <td align="right"><strong>100</strong></td>
								<% ELSEIF cdbl(budsubBI) = 0 and Cdbl(subBI) = 0 then %>
									  <td align="right"><strong>0</strong></td>
								<% ELSEIF cdbl(budsubBI) <> 0 and Cdbl(subBI) <> 0 then %>
									  <td align="right"><strong><%=FormatNumber((cdbl(budsubBI) - Cdbl(subBI)) / cdbl(budsubBI) * 100,2)%></strong></td>
								<% END IF %>
								<td align="right"><strong><%=FormatNumber(subSBI,0)%></strong></td>	
								<% IF Cdbl(LuasHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>	
								   <td align="right"><strong><%=FormatNumber((Cdbl(subSBI)/Cdbl(LuasHa)),0)%></strong></td>	
								<% END IF %>   
								<td align="right"><strong><%=FormatNumber(budsubSBI,0)%></strong></td>
								<% IF Cdbl(budHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>	
								   <td align="right"><strong><%=FormatNumber(cdbl(budsubSBI)/cdbl(budHa),0)%></strong></td>
								<% END IF %>
								<% IF cdbl(budsubSBI) = 0 and Cdbl(subSBI) <> 0 then %>
									  <td align="right"><strong>-100</strong></td>
								<% ELSEIF cdbl(budsubSBI) <> 0 and Cdbl(subSBI) = 0 then %>	  
									  <td align="right"><strong>100</strong></td>
								<% ELSEIF cdbl(budsubSBI) = 0 and Cdbl(subSBI) = 0 then %>
									  <td align="right"><strong>0</strong></td>
								<% ELSEIF cdbl(budsubSBI) <> 0 and Cdbl(subSBI) <> 0 then %>
									  <td align="right"><strong><%=FormatNumber((cdbl(budsubSBI) - Cdbl(subSBI)) / cdbl(budsubSBI) * 100,2)%></strong></td>
								<% END IF %>
								<td align="right"><strong><%=FormatNumber(budsubTHN,0)%></strong></td>
								<% IF Cdbl(budHa) = 0  then  %>
								   <td align="right">0</td>
								<% ELSE %>	
								   <td align="right"><strong><%=FormatNumber(cdbl(budsubTHN)/cdbl(budHa),0)%></strong></td>
								<% END IF %>
						   </tr>
		
				 <% 
					Dim tot_rp_ha_bi, tot_rp_ha_sbi, tot_rp_ha_budbi, tot_rp_ha_budsbi, tot_rp_ha_budTHN
					IF Cdbl(budHa) = 0  THEN  				       
					   tot_rp_ha_bi = 0
					   tot_rp_ha_sbi = 0
					ELSE 						
					   tot_rp_ha_bi  =  Cdbl(totsubBI)/Cdbl(LuasHa)
					   tot_rp_ha_sbi =  Cdbl(totsubSBI)/Cdbl(LuasHa)
					END IF 
					IF Cdbl(budHa) = 0  THEN 
					  tot_rp_ha_budbi = 0
					  tot_rp_ha_budsbi = 0
					  tot_rp_ha_budTHN = 0
					ELSE
					  tot_rp_ha_budbi  = cdbl(budtotsubBI)/cdbl(budHa)
					  tot_rp_ha_budsbi = cdbl(budtotsubSBI)/cdbl(budHa)
					  tot_rp_ha_budTHN = cdbl(budtotsubTHN)/cdbl(budHa)
					END IF
				 %>   
				   

					<tr bgcolor="#ffff01">
						<%if v_txttype = "detail" then%>
							<td colspan="<%= 1 + 2 + nSpanTotal %>" align="left"><strong>Total </strong></td>
						<%else%>
							<td colspan="<%= 2 + nSpanTotal %>" align="left"><strong>Total </strong></td>
						<%end if%>
						
						<td align="right"><strong><%=FormatNumber(totsubBI,0)%></strong></td>				
						<td align="right"><strong><%=FormatNumber(Cdbl(tot_rp_ha_bi),0)%></strong></td>	 
						<td align="right"><strong><%=FormatNumber(budtotsubBI,0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_ha_budbi),0)%></strong></td>
						<% IF cdbl(tot_rp_ha_budbi) = 0 and cdbl(tot_rp_ha_bi) <> 0 then %>
							  <td align="right"><strong>-100</strong></td>
						<% ELSEIF cdbl(tot_rp_ha_budbi) <> 0 and cdbl(tot_rp_ha_bi) = 0 then %>	  
							  <td align="right"><strong>100</strong></td>
						<% ELSEIF cdbl(tot_rp_ha_budbi) = 0 and cdbl(tot_rp_ha_bi) = 0 then %>
							  <td align="right"><strong>0</strong></td>
						<% ELSEIF cdbl(tot_rp_ha_budbi) <> 0 and cdbl(tot_rp_ha_bi) <> 0 then %>
							  <td align="right"><strong><%=FormatNumber((cdbl(tot_rp_ha_budbi) - cdbl(tot_rp_ha_bi)) / cdbl(tot_rp_ha_budbi) * 100,2)%></strong></td>
						<% END IF %>
						<td align="right"><strong><%=FormatNumber(totsubSBI,0)%></strong></td>	
						<td align="right"><strong><%=FormatNumber(Cdbl(tot_rp_ha_sbi),0)%></strong></td>		
						<td align="right"><strong><%=FormatNumber(budtotsubSBI,0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_ha_budsbi),0)%></strong></td>
				
						<% IF cdbl(tot_rp_ha_budsbi) = 0 and cdbl(tot_rp_ha_sbi) <> 0 then %>
							  <td align="right"><strong>-100</strong></td>
						<% ELSEIF cdbl(tot_rp_ha_budsbi) <> 0 and cdbl(tot_rp_ha_sbi) = 0 then %>	  
							  <td align="right"><strong>100</strong></td>
						<% ELSEIF cdbl(tot_rp_ha_budsbi) = 0 and cdbl(tot_rp_ha_sbi) = 0 then %>
							  <td align="right"><strong>0</strong></td>
						<% ELSEIF cdbl(tot_rp_ha_budsbi) <> 0 and cdbl(tot_rp_ha_sbi) <> 0 then %>
							  <td align="right"><strong><%=FormatNumber((cdbl(tot_rp_ha_budsbi) - cdbl(tot_rp_ha_sbi)) / cdbl(tot_rp_ha_budsbi) * 100,2)%></strong></td>
						<% END IF %>
						<td align="right"><strong><%=FormatNumber(budtotsubTHN ,0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_ha_budTHN),0)%></strong></td>	
					</tr>
		</tbody>
	</table>
<%
end if	
%>
	
<% 
if  v_txtfcba <> "--ALL--" and v_txtfcba <> ""  and v_txtfcba <>  "--SUMMARY--" then %>
   <table width="90%" class="data">
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
	    
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=MgrName%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=KTUName%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
 <% End if %>  
</div>
</body>
</html>