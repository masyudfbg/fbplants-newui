<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtDivision, v_txtPlantingDate, v_txtBlok
Dim v_txtyear,v_txtmonth
Dim txtDivision, txtBlok, txtPlantingDate
Dim v_vstatus,v_vCROPTYPE,v_vJenisReport, v_vOwnership
Dim tempyear,tempmonth
Dim tempDivision, tempBlok, tempPlantingDate 
Dim SprevCode,Sprevjobname
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
            
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
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut, " 

						   
			wherecons = " where company is not null "
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
			
			' if len(trim(v_txtDivision)) > 0 then 
			   ' if v_txtDivision <> "--SUMMARY--" then
			      ' if v_txtDivision <>  "--ALL--"  then
					' whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      ' end if
                  ' selectCons = selectCons &  " , DIVISION "
			   ' else
			      ' nSpanTotal = nSpanTotal - 1 	  
			   ' end if
			' end if
            
			' if len(trim(v_txtPlantingDate)) > 0 then 
				' if v_txtPlantingDate <> "--SUMMARY--"   then
				   ' if v_txtPlantingDate <>  "--ALL--" then
					  ' whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   ' end if
				   ' selectCons = selectCons &  " , PLANTINGYEAR "
				' else
			      ' nSpanTotal = nSpanTotal - 1 	  
			   ' end if
				
			' end if
			 
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Analisa Biaya Tidak Langsung</title>
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
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<table id="form-table" class="centered-margin">
		<form name="frmBTL" method="post" action="<%=ScriptName%>">
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
							MinYr = Yr - 2
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
					
					<td><input type="submit" value="View" /></td>
				</tr>
			</table>	
		</form>
	</table>
	
	<div id="header">
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
strha =  " SELECT GroupPlantation,SUM(ha) HA, SUM(bud_ha) bud_ha from ( " & _
         "  SELECT " & strFieldCons & " DECODE(SUM (ha), NULL,1,0,1,SUM (ha)) HA,DECODE(SUM (bud_ha), NULL,1,0,1,SUM (bud_ha)) bud_ha FROM   " & _
		 "  ( SELECT fcba, croptype, ownership, division, status, TO_CHAR(plantingdate,'YYYY') PLANTINGYEAR, SUM (hectarageplanted) ha,0 bud_ha FROM field   " & _ 
		 "	 GROUP BY fcba, croptype, ownership, division, status, plantingdate   " & _
		 "	 UNION ALL  " & _
		 "	 select fcba,croptype,ownership,division,status,PLANTINGYEAR,0 ha, sum(hectarageplanted)  bud_ha  from    " & _
		 "	 ( select  distinct t.fcba,t.croptype,t.ownership,t.division,t.status,to_char(t.plantingdate,'YYYY') PLANTINGYEAR, hectarageplanted   " & _
		 "	   from BUDGET_TRANSACTION_ESTATE t )   " & _
		 "	 GROUP BY fcba, croptype, ownership, division, status, PLANTINGYEAR   " & _
		 "  ) group by fcba )" & whereCons & " group by GroupPlantation "
		
Set rsBTL = Server.CreateObject("ADODB.Recordset")
rsBTL.Open strha, connstr, adOpenStatic
if Not rsBTL.eof then
		LuasHa	=	rsBTL("ha")
		budHA = rsBTL("bud_ha")
end if
rsBTL.Close
Set rsBTL = Nothing
%>
<table border="0" width="100%" id="product-table">
	<tr>
		<th style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="5">BULAN INI</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="5">S/D BULAN INI</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2">REALISASI</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;"  rowspan="3">%</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2">REALISASI</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;"  rowspan="3">%</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">Luas Ha</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(LuasHa,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(budHa,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(LuasHa,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(budHa,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2" align="right"><%=FormatNumber(budHa,2)%></th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" >CODE</th>
		<th style="background-color:#00A300;color:#fff;" >JENIS PEKERJAAN</th>
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
<%
'Dim rsBTL, strGroupBy, rowno, rsalokasiBTL, stralokasiBTL
Dim  rowno, strBTL
Dim subBI,subSBI,totsubBI,totsubSBI,budsubBI,budsubSBI,budSubthn,budtotsubBI,budtotsubSBI,budtotsubTHN
Dim rsjob,strjob,debetjob,kreditjob,notAlokasi

strjob= "select report_posting_api.getPostingDataDirect('KSR','ANALISABTL','p_AnalisaBTL.asp','DATADEBET_ESTATE','JOB') debet, report_posting_api.getPostingDataDirect('KSR','ANALISABTL','p_AnalisaBTL.asp','DATAKREDIT_ESTATE','JOB')   kredit  from dual"
Set rsjob = Server.CreateObject("ADODB.Recordset")
'response.write strjob
rsjob.Open strjob, connstr, adOpenStatic
if Not rsjob.eof then
   debetjob = rsjob("debet")
   kreditjob = rsjob("kredit")
end if


rowno		=	0
strBTL		=	" select GroupPlantation " & selectCons & " , grp,jobcode,description, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from ( " & _
                " select " & strFieldCons & " grp,jobcode,description, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from ( " & _
				"	select t.fcba,'A1' grp, t.jobcode,job_api.get_jobcode_name(t.jobcode,t.fcba) Description, sum(t.amount) BI, 0 SBI,0 budBI,0 budSBI,0 budthn " & _
				"	from cost_activity t where instr('" & debetjob & "',t.jobcode) <> 0  and to_char(t.tdate,'YYYY')=  '" & v_txtyear & "' and to_char(t.tdate,'MM')=  lpad('" & v_txtmonth & "',2,0) " & _
				"	group by t.fcba, t.jobcode " & _
				"	union " & _
				"	select t.fcba,'A1' grp, t.jobcode,job_api.get_jobcode_name(t.jobcode,t.fcba) Description,0 BI, sum(t.amount) SBI,0 budBI,0 budSBI,0 budthn " & _
				"	from cost_activity t where instr('" & debetjob & "',t.jobcode) <> 0  and to_char(t.tdate,'YYYY')=  '" & v_txtyear& "' and to_char(t.tdate,'MM') <=   lpad('" & v_txtmonth & "',2,0) " & _
				"	group by t.fcba, t.jobcode " & _
				"	union " & _
				"	select t.fcba,'A2' grp, t.jobcode,job_api.get_jobcode_name(t.jobcode,t.fcba) Description, sum(t.amount)* -1 BI, 0 SBI,0 budBI,0 budSBI,0 budthn " & _
				"	from cost_activity t where instr('" & kreditjob & "',t.jobcode)  <> 0  and to_char(t.tdate,'YYYY')=  '" & v_txtyear & "' and to_char(t.tdate,'MM')=  lpad('" & v_txtmonth & "',2,0) " & _
				"	group by t.fcba, t.jobcode " & _
				"	union " & _
				"	select t.fcba, 'A2' grp, t.jobcode,job_api.get_jobcode_name(t.jobcode,t.fcba) Description,0 BI, sum(t.amount)* -1  SBI,0 budBI,0 budSBI,0 budthn " & _
				"	from cost_activity t where instr('" & kreditjob & "',t.jobcode)  <> 0  and to_char(t.tdate,'YYYY')=  '" & v_txtyear & "' and to_char(t.tdate,'MM') <=   lpad('" & v_txtmonth & "',2,0) " & _
				"	group by t.fcba, t.jobcode " & _
				"	union " & _
				"	select btm.fcba,'A2' grp, btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,sum(" & strbud_BI & ")* -1 budBI,0 budSBI,0 budthn  from BUDGET_TRANSACTION_ESTATE btm " & _
				"	where btm.budget_year ='" & v_txtyear & "' and instr('" & kreditjob & "',btm.jobcode) <> 0 " & _
				"	group by btm.fcba,btm.jobcode " & _
				"	union " & _
				"	select btm.fcba,'A2' grp,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,0 budBI,sum(" & strbud_SBI & ")* -1 budSBI,0 budthn  from BUDGET_TRANSACTION_ESTATE btm " & _
				"	where btm.budget_year ='" & v_txtyear & "' and instr('" & kreditjob & "',btm.jobcode) <> 0 " & _
				"	group by btm.fcba,btm.jobcode " & _
				"	union " & _
				"	select btm.fcba,'A2' grp,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,0 budBI,0 budSBI,sum(btm.budget_01+btm.budget_02+btm.budget_03+btm.budget_04+btm.budget_05+btm.budget_06+btm.budget_07+btm.budget_08+btm.budget_09+btm.budget_10+btm.budget_11+btm.budget_12)* -1 budthn  from BUDGET_TRANSACTION_ESTATE btm " & _
				"	where btm.budget_year ='" & v_txtyear & "' and  instr('" & kreditjob & "',btm.jobcode) <> 0 " & _
				"	group by btm.fcba,btm.jobcode " & _
				"	union " & _
				"	select btm.fcba,'A1' grp,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,sum(" & strbud_BI & ") budBI,0 budSBI,0 budthn  from BUDGET_TRANSACTION_INDIRCOST btm " & _
				"	where btm.budget_year ='" & v_txtyear & "'  and  instr('" & debetjob & "',btm.jobcode) <> 0  " & _
				"	group by btm.fcba,btm.jobcode " & _
				"	union " & _
				"	select btm.fcba,'A1' grp,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,0 budBI,sum(" & strbud_SBI & ") budSBI,0 budthn  from BUDGET_TRANSACTION_INDIRCOST btm " & _
				"	where btm.budget_year ='" & v_txtyear & "' and  instr('" & debetjob & "',btm.jobcode) <> 0 " & _
				"	group by btm.fcba,btm.jobcode" & _
				"	union " & _
				"	select btm.fcba,'A1' grp,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,0 budBI,0 budSBI,sum(btm.budget_01+btm.budget_02+btm.budget_03+btm.budget_04+btm.budget_05+btm.budget_06+btm.budget_07+btm.budget_08+btm.budget_09+btm.budget_10+btm.budget_11+btm.budget_12) budthn  from BUDGET_TRANSACTION_INDIRCOST btm " & _ 
				"	where btm.budget_year ='" & v_txtyear & "' and  instr('" & debetjob & "',btm.jobcode) <> 0 " & _
				"	group by btm.fcba,btm.jobcode " & _
				"	 ) " & _
				"	group by fcba,grp,jobcode,description )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by GroupPlantation,grp,jobcode,description " & selectCons & " order by grp,jobcode" & orderCons 




Set rsBTL = Server.CreateObject("ADODB.Recordset")
rsBTL.Open strBTL, connstr, adOpenStatic
response.write strBTL
if Not rsBTL.eof then
Dim TotalBTL, TotalTyBTL
TotalBTL = 0
TotalTyBTL = 0
	While Not rsBTL.eof
			rowno	=	rowno	+	1
		If 	rsBTL("grp") <> "A2" Then
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
						   <%
						   SprevCode   = rsBTL("JOBCODE")
						   Sprevjobname = rsBTL("description")
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
				
				
				<% ' BI %>
				<td align="right"><%=FormatNumber(rsBTL("BI"),2)%></td>
				<% IF Cdbl(LuasHa) = 0  then  %>
				    <td align="right">0</td>
			    <% ELSE %>		
				   <td align="right"><%=FormatNumber((Cdbl(rsBTL("BI"))/Cdbl(LuasHa)),2)%></td>
				<% END IF %>
				
				<td align="right"><%=FormatNumber(rsBTL("budBI"),2)%></td>
				<% IF Cdbl(budHa) = 0  then  %>
				    <td align="right">0</td>
			    <% ELSE %>		
				   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budBI"))/Cdbl(budHa)),2)%></td>
				<% END IF %>
				
				<% IF (Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) = 0 and (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)) > 0 then %>
				  <td align="right">-100</td>
				<% ELSEIF (Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) <> 0 and (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)) = 0 then %>
				  <td align="right">100</td>
				<% ELSEIF (Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) = 0 and (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)) = 0 then %>
				  <td align="right">0</td>
				<% ELSEIF (Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) <> 0 and (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)) <> 0 then %>
				   <td align="right"><%=FormatNumber(((Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) - (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)))/(Cdbl(rsBTL("budBI"))/Cdbl(budHa) )*100,2)%></td>
				<% END IF %>
				
				<% ' SBI %>
				<td align="right"><%=FormatNumber(rsBTL("SBI"),2)%></td>
				<% IF Cdbl(LuasHa) = 0  then  %>
				    <td align="right">0</td>
			    <% ELSE %>		
				   <td align="right"><%=FormatNumber((Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)),2)%></td>
				<% END IF %>
				
				<td align="right"><%=FormatNumber(rsBTL("budSBI"),2)%></td>
				<% IF Cdbl(budHa) = 0  then  %>
				    <td align="right">0</td>
			    <% ELSE %>		
				   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budSBI"))/Cdbl(budHa)),2)%></td>
				<% END IF %>
				
				<% IF (Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) = 0 and (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)) <> 0 then %>
				  <td align="right">-100</td>
				<% ELSEIF (Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) <> 0 and (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)) = 0 then %>
				  <td align="right">100</td>
				<% ELSEIF (Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) = 0 and (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)) = 0 then %>
				  <td align="right">0</td>
				<% ELSEIF (Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) <> 0 and (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)) <> 0 then %>
				   <td align="right"><%=FormatNumber(((Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) - (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)))/(Cdbl(rsBTL("budSBI"))/Cdbl(budHa) )*100,2)%></td>
				<% END IF %>
				
				<td align="right"><%=FormatNumber(rsBTL("budthn"),2)%></td>
				<% IF Cdbl(budHa) = 0  then  %>
				    <td align="right">0</td>
			    <% ELSE %>		
				   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budthn"))/Cdbl(budHa)),2)%></td>
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
						<tr bgcolor="#F4F4A0">
							<td colspan="<%= 2 + nSpanTotal %>" align="left"><strong>Total Biaya Tak Langsung</strong></td>
							<td align="right"><strong><%=FormatNumber(subBI,2)%></strong></td>
							<% IF Cdbl(budHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>						
							   <td align="right"><strong><%=FormatNumber((Cdbl(subBI)/Cdbl(LuasHa)),2)%></strong></td>	
							<% END IF %>   
							<td align="right"><strong><%=FormatNumber(budsubBI,2)%></strong></td>
							<% IF Cdbl(budHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>	
							   <td align="right"><strong><%=FormatNumber(cdbl(budsubBI)/cdbl(budHa),2)%></strong></td>
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
							<td align="right"><strong><%=FormatNumber(subSBI,2)%></strong></td>	
							<% IF Cdbl(LuasHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>	
							   <td align="right"><strong><%=FormatNumber((Cdbl(subSBI)/Cdbl(LuasHa)),2)%></strong></td>	
							<% END IF %>   
							<td align="right"><strong><%=FormatNumber(budsubSBI,2)%></strong></td>
							<% IF Cdbl(budHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>	
							   <td align="right"><strong><%=FormatNumber(cdbl(budsubSBI)/cdbl(budHa),2)%></strong></td>
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
							<td align="right"><strong><%=FormatNumber(budsubTHN,2)%></strong></td>
							<% IF Cdbl(budHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>	
							   <td align="right"><strong><%=FormatNumber(cdbl(budsubTHN)/cdbl(budHa),2)%></strong></td>
							<% END IF %>
						</tr>
				 <%            notAlokasi = "F"
		            End if 
					 subBI     = Cdbl(rsBTL("BI"))
					 subSBI    = Cdbl(rsBTL("SBI"))
					 budsubBI  = Cdbl(rsBTL("budBI"))
					 budsubSBI = Cdbl(rsBTL("budSBI"))
					 budSubthn = Cdbl(rsBTL("budTHN"))
				 
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
						   <%
						   SprevCode   = rsBTL("JOBCODE")
						   Sprevjobname = rsBTL("description")
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
					
					
					<% ' BI %>
					<td align="right"><%=FormatNumber(rsBTL("BI"),2)%></td>
					<% IF Cdbl(LuasHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("BI"))/Cdbl(LuasHa)),2)%></td>
					<% END IF %>
					
					<td align="right"><%=FormatNumber(rsBTL("budBI"),2)%></td>
					<% IF Cdbl(budHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budBI"))/Cdbl(budHa)),2)%></td>
					<% END IF %>
					
					<% IF (Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) = 0 and (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)) <> 0 then %>
					  <td align="right">-100</td>
					<% ELSEIF (Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) <> 0 and (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)) = 0 then %>
					  <td align="right">100</td>
					<% ELSEIF (Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) = 0 and (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)) = 0 then %>
					  <td align="right">0</td>
					<% ELSEIF (Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) <> 0 and (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)) <> 0 then %>
					   <td align="right"><%=FormatNumber(((Cdbl(rsBTL("budBI"))/Cdbl(budHa) ) - (Cdbl(rsBTL("BI"))/Cdbl(LuasHa)))/(Cdbl(rsBTL("budBI"))/Cdbl(budHa) )*100,2)%></td>
					<% END IF %>
					
					<% ' SBI %>
					<td align="right"><%=FormatNumber(rsBTL("SBI"),2)%></td>
					<% IF Cdbl(LuasHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)),2)%></td>
					<% END IF %>
					
					<td align="right"><%=FormatNumber(rsBTL("budSBI"),2)%></td>
					<% IF Cdbl(budHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budSBI"))/Cdbl(budHa)),2)%></td>
					<% END IF %>
					
					<% IF (Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) = 0 and (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)) <> 0 then %>
					  <td align="right">-100</td>
					<% ELSEIF (Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) <> 0 and (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)) = 0 then %>
					  <td align="right">100</td>
					<% ELSEIF (Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) = 0 and (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)) = 0 then %>
					  <td align="right">0</td>
					<% ELSEIF (Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) <> 0 and (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)) <> 0 then %>
					   <td align="right"><%=FormatNumber(((Cdbl(rsBTL("budSBI"))/Cdbl(budHa) ) - (Cdbl(rsBTL("SBI"))/Cdbl(LuasHa)))/(Cdbl(rsBTL("budSBI"))/Cdbl(budHa) )*100,2)%></td>
					<% END IF %>
					
					<td align="right"><%=FormatNumber(rsBTL("budthn"),2)%></td>
					<% IF Cdbl(budHa) = 0  then  %>
						<td align="right">0</td>
					<% ELSE %>		
					   <td align="right"><%=FormatNumber((Cdbl(rsBTL("budthn"))/Cdbl(budHa)),2)%></td>
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
		
		rsBTL.MoveNext
	Wend
end if
	%>
	
	              <tr bgcolor="#F4F4A0">
							<td colspan="<%= 2 + nSpanTotal %>" align="left"><strong>Total Alokasi Biaya Tak Langsung</strong></td>
							<td align="right"><strong><%=FormatNumber(subBI,2)%></strong></td>
							<% IF Cdbl(budHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>						
							   <td align="right"><strong><%=FormatNumber((Cdbl(subBI)/Cdbl(LuasHa)),2)%></strong></td>	
							<% END IF %>   
							<td align="right"><strong><%=FormatNumber(budsubBI,2)%></strong></td>
							<% IF Cdbl(budHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>	
							   <td align="right"><strong><%=FormatNumber(cdbl(budsubBI)/cdbl(budHa),2)%></strong></td>
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
							<td align="right"><strong><%=FormatNumber(subSBI,2)%></strong></td>	
							<% IF Cdbl(LuasHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>	
							   <td align="right"><strong><%=FormatNumber((Cdbl(subSBI)/Cdbl(LuasHa)),2)%></strong></td>	
							<% END IF %>   
							<td align="right"><strong><%=FormatNumber(budsubSBI,2)%></strong></td>
							<% IF Cdbl(budHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>	
							   <td align="right"><strong><%=FormatNumber(cdbl(budsubSBI)/cdbl(budHa),2)%></strong></td>
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
							<td align="right"><strong><%=FormatNumber(budsubTHN,2)%></strong></td>
							<% IF Cdbl(budHa) = 0  then  %>
							   <td align="right">0</td>
							<% ELSE %>	
							   <td align="right"><strong><%=FormatNumber(cdbl(budsubTHN)/cdbl(budHa),2)%></strong></td>
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
	           

                <tr bgcolor="#F4F4A0">
					<td colspan="<%= 2 + nSpanTotal %>" align="left"><strong>Total </strong></td>
					<td align="right"><strong><%=FormatNumber(totsubBI,2)%></strong></td>				
					<td align="right"><strong><%=FormatNumber(Cdbl(tot_rp_ha_bi),2)%></strong></td>	 
					<td align="right"><strong><%=FormatNumber(budtotsubBI,2)%></strong></td>
					<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_ha_budbi),2)%></strong></td>
					<% IF cdbl(tot_rp_ha_budbi) = 0 and cdbl(tot_rp_ha_bi) <> 0 then %>
					      <td align="right"><strong>-100</strong></td>
					<% ELSEIF cdbl(tot_rp_ha_budbi) <> 0 and cdbl(tot_rp_ha_bi) = 0 then %>	  
					      <td align="right"><strong>100</strong></td>
					<% ELSEIF cdbl(tot_rp_ha_budbi) = 0 and cdbl(tot_rp_ha_bi) = 0 then %>
					      <td align="right"><strong>0</strong></td>
					<% ELSEIF cdbl(tot_rp_ha_budbi) <> 0 and cdbl(tot_rp_ha_bi) <> 0 then %>
					      <td align="right"><strong><%=FormatNumber((cdbl(tot_rp_ha_budbi) - cdbl(tot_rp_ha_bi)) / cdbl(tot_rp_ha_budbi) * 100,2)%></strong></td>
					<% END IF %>
					<td align="right"><strong><%=FormatNumber(totsubSBI,2)%></strong></td>	
			  		<td align="right"><strong><%=FormatNumber(Cdbl(tot_rp_ha_sbi),2)%></strong></td>		
					<td align="right"><strong><%=FormatNumber(budtotsubSBI,2)%></strong></td>
					<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_ha_budsbi),2)%></strong></td>
			
					<% IF cdbl(tot_rp_ha_budsbi) = 0 and cdbl(tot_rp_ha_sbi) <> 0 then %>
					      <td align="right"><strong>-100</strong></td>
					<% ELSEIF cdbl(tot_rp_ha_budsbi) <> 0 and cdbl(tot_rp_ha_sbi) = 0 then %>	  
					      <td align="right"><strong>100</strong></td>
					<% ELSEIF cdbl(tot_rp_ha_budsbi) = 0 and cdbl(tot_rp_ha_sbi) = 0 then %>
					      <td align="right"><strong>0</strong></td>
					<% ELSEIF cdbl(tot_rp_ha_budsbi) <> 0 and cdbl(tot_rp_ha_sbi) <> 0 then %>
					      <td align="right"><strong><%=FormatNumber((cdbl(tot_rp_ha_budsbi) - cdbl(tot_rp_ha_sbi)) / cdbl(tot_rp_ha_budsbi) * 100,2)%></strong></td>
					<% END IF %>
					<td align="right"><strong><%=FormatNumber(budtotsubTHN ,2)%></strong></td>
					<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_ha_budTHN),2)%></strong></td>	
				</tr>
	
</table>
<% if  v_txtfcba <> "--ALL--" and v_txtfcba <> ""  and v_txtfcba <>  "--SUMMARY--" then %>
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