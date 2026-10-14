<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa
Dim v_xBatch,v_txtjob
Dim v_txtDivision, v_txtPlantingDate, v_txtBlok
Dim v_txtyear,v_txtmonth
Dim txtDivision, txtBlok, txtPlantingDate
Dim v_vstatus,v_vCROPTYPE,v_vJenisReport, v_vOwnership
Dim tempDivision, tempBlok, tempPlantingDate 
Dim SprevCode,Sprevjobname
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim tempjob,tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
Dim ScriptName,nSpanTotal
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
			v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	  = Request("txtPlantation")
			v_txtGroupRegion  	  = Request("txtGroupRegion")
			v_txtRegion  		  = Request("txtRegion")
			v_txtcompany  		  = Request("txtCompany")
			v_txtfcba 			  = Request("txtfcba")
			v_txtjob			  = Request("txtjob")
			v_xBatch              = Request("xBatch")
			v_txtBlok			  = Request("txtBlok")
			v_txtyear             = Request("txtYear")
			v_txtmonth            = Request("txtMonth")
			nSpanTotal = 5
			if v_txtyear ="" then
			   v_txtyear  = right(textdate(Now()),4)
			   v_txtmonth = left(textdate(Now()),2)
			end if   
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " TT" & v_txtPlantingDate
			'response.write  "YEAR " & v_txtyear & "  month " & v_txtmonth& "  status " & v_vstatus & " crop " & v_vCROPTYPE  & " owner " & v_vOwnership & " jenis " & v_vJenisReport  
			
			Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
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
			
			
			Dim rsJob, strjob,bibitjob

			strjob= " select report_posting_api.getPostingDataDirect('KSR','PEMBIBITAN','p_AnalisaBibit.Asp','PRE_AREA','JOB')||report_posting_api.getPostingDataDirect('KSR','PEMBIBITAN','p_AnalisaBibit.Asp','PRE_NURSERY','JOB')||report_posting_api.getPostingDataDirect('KSR','PEMBIBITAN','p_AnalisaBibit.Asp','MAIN_NURSERY','JOB') job from dual "
			Set rsjob = Server.CreateObject("ADODB.Recordset")
			'response.write strjob
			rsjob.Open strjob, connstr, adOpenStatic
			if Not rsjob.eof then
			   bibitjob = rsjob("job")
			end if
			
%>
<html>

<head>
<title>ePlaS Reporting - Laporan Pembiayaan Bibit</title>
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
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<table id="form-table" class="centered-margin">
		<form name="frmBibit" method="post" action="<%=ScriptName%>">
		
		    <!--#include file="FormSelectNewSum.html"-->
				
			<table id="form-table" class="centered-margin">	
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
			<td><font face="verdana" size="1">JENIS PEKERJAAN</font></td>
			<td>
				<select name="txtJob">
					<option value="--ALL--" <%if request("v_txtJob")="--ALL--" then%> selected <%End if%>>--ALL--</option>
					<%
					
					Set rsJob = Server.CreateObject("ADODB.Recordset")
					rsJob.CursorLocation = adUseClient
					strsql =  " select distinct a.fccode,a.fcname from job a where instr('" & bibitjob & "',a.fccode) <> 0 and user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null ORDER BY a.fccode ASC"
					rsJob.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsJob.eof then
					DO while not rsJob.eof
					  if v_txtJob= rsJob("FCCODE") then
						%>			
						<option value="<%=rsJob("FCCODE")%>" Selected><%=rsJob("FCCODE")%> - <%=rsJob("fcname")%></option>
				   <% Else %>	
						<option value="<%=rsJob("FCCODE")%>"><%=rsJob("FCCODE")%> - <%=rsJob("fcname")%></option>
				   <% End If
					  rsJob.movenext
					loop
					End if
					rsJob.Close
					set rsJob = Nothing
					%>
					
				</select>
			</td>
			<td><font face="verdana" size="1">BATCH</font></td>
			<td>
				<select name="xBatch">
				<option value="--ALL--" <%if request("v_xBatch")="--ALL--" then%> selected <%End if%>>--ALL--</option>
				<%
				Dim rsBatch
				Set rsBatch = Server.CreateObject("ADODB.Recordset")
				rsBatch.CursorLocation = adUseClient
				strsql =  " SELECT DISTINCT NURSERYCODE FCCODE, DESCRIPTION FROM NURSERY where user_company_api.GetCompanyCode('" & Session("CodeName") & "',fcba) is not null ORDER BY NURSERYCODE ASC"
				rsBatch.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
				if not rsBatch.eof then
				DO while not rsBatch.eof
				  if v_xBatch= rsBatch("FCCODE") then
					%>			
					<option value="<%=rsBatch("FCCODE")%>" Selected><%=rsBatch("FCCODE")%></option>
			   <% Else %>	
					<option value="<%=rsBatch("FCCODE")%>"><%=rsBatch("FCCODE")%></option>
			   <% End If
			      rsBatch.movenext
				loop
				End if
				rsBatch.Close
				set rsBatch = Nothing
				%>
				</select>
			</td>
			<td><input type="submit" value="View" /></td>
		</tr>
		</table>
	</form>
</table>
<%  Dim cur_month,strbud_BI,strbud_SBI,x1
	    strbud_BI = "bp.budget_01" 
	    strbud_SBI = "bp.budget_01"
		
		
		cur_month = request.Form("vMonth") 
		
		IF Len(cur_month) > 0 then

			strbud_BI = "" 
	        strbud_SBI = ""
	
			x1=1
			
			Do While x1<= cint(cur_month) 'Output the values from 0 to 10
			  'response.write(i & "<br \>")
			  'response.write  " Budget_" & Lpad(x1,"0",2) & " +"
			  strbud_SBI = strbud_SBI & " bp.budget_" & Lpad(x1,"0",2) & " +"
			  x1 = x1 + 1 'increment the value of i for next time loop executes
			Loop

			strbud_BI = "bp.budget_" & Lpad(cur_month,"0",2)
	
			strbud_SBI = Mid(strbud_SBI,1,Len(strbud_SBI)-1)
		
			
            ' response.write  strbudTBS_BI & "<br>"
			' response.write strbudCPO_BI & "<br>"
			' response.write strbudKER_BI & "<br>"
			
			
			' response.write strbudTBS_SBI  & "<br>"
			' response.write strbudCPO_SBI  & "<br>"
			' response.write strbudKER_SBI  & "<br>"
		END IF	


	
	if  v_xBatch <> "--ALL--"  and v_xBatch <> "--SUMMARY--" then
        if	v_xBatch <> ""  then
	       strActualBibit = strActualBibit & " AND NURSERYCODE = '" & v_xBatch  & "'"
	       strBudgetBibit = strBudgetBibit & " AND NURSERYCODE = '" & v_xBatch  & "'"
		 end if  
	end if
	
	Dim rsActualBibit, strActualBibit, vActualBibit, vBudgetBibit, rsBudgetBibit, strBudgetBibit
	strActualBibit	=	" SELECT GroupPlantation,sum(ACTUALBIBIT) ACTUALBIBIT From ( SELECT " & strFieldCons & " NVL(((SUM(OPENINGQTY)+SUM(QTYRECEIVED))-(SUM(QTYCULLED)+SUM(QTYISSUED))),0) ACTUALBIBIT" & _
						" FROM NURSERYLEDGER WHERE  YEAR = " & yyyy & " AND MONTH = " & MM & strActualBibit & " Group By Fcba )  " & whereCons & " group by GroupPlantation "
						
	strBudgetBibit	=	" SELECT GroupPlantation,nvl(SUM (budgetbibit),0) budgetbibit" & _
						"   FROM (  SELECT " & strFieldCons & " nyear," & _
						"                  nurserycode," & _
						"                  SUM (budgetbibit) budgetbibit" & _
						"             FROM (  SELECT BUDGET_YEAR nyear," & _
						"                            BATCHCODE nurserycode," & _
						"                            QTYNURESRY budgetbibit," & _
						"                            FCBA" & _
						"                       FROM BUDGET_TRANSACTION_NURSERY" & _
						"                   GROUP BY BUDGET_YEAR," & _
						"                            BATCHCODE," & _
						"                            QTYNURESRY," & _
						"                            FCBA)" & _
						"  WHERE nYEAR = " & yyyy & " " & _
						"         GROUP BY nyear, nurserycode, fcba) " & whereCons & " group by GroupPlantation "
						
    
	
	
	
	'response.write strActualBibit
	Set rsActualBibit = Server.CreateObject("ADODB.Recordset")
	rsActualBibit.Open strActualBibit, connstr, adOpenStatic
    'RESPONSE.WRITE  strBudgetBibit
	Set rsBudgetBibit = Server.CreateObject("ADODB.Recordset")
	rsBudgetBibit.Open strBudgetBibit, connstr, adOpenStatic
	
	if Not rsActualBibit.eof then
		vActualBibit	=	rsActualBibit("ACTUALBIBIT")
	end if
	rsActualBibit.Close
	set rsActualBibit = Nothing
	
	if Not rsBudgetBibit.eof then
		vBudgetBibit	=	Cdbl(rsBudgetBibit("BUDGETBIBIT"))
	end if
	rsBudgetBibit.Close
	set rsBudgetBibit = Nothing
	
			
%>
<table border="0" id="product-table" class="xthf-site" width="100%">
	<thead>
        <tr>
          <th colspan="19" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>	
		</div>
		<div id="subheader" class="centered">
			<h3>LAPORAN PEMBIAYAAN BIBITAN</h3>
			<h4>Periode&nbsp;:&nbsp;<%=Bulan(mm)%>&nbsp;<%=yyyy%></h4>
		</div>		
	</div>

           </th>
     </tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
		<th style="background-color:#00A300;color:#fff;" colspan="5">BULAN INI</th>
		<th style="background-color:#00A300;color:#fff;" colspan="5">S/D BULAN INI</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2" rowspan="3">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">REALISASI</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="3">%</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">REALISASI</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="3">%</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">JUMLAH BIBIT</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2" align="right"><%=FormatNumber(vActualBibit,0)%></th>
		<th style="background-color:#00A300;color:#fff;" colspan="2" align="right"><%=FormatNumber(vBudgetBibit,0)%></th>
		<th style="background-color:#00A300;color:#fff;" colspan="2" align="right"><%=FormatNumber(vActualBibit,0)%></th>
		<th style="background-color:#00A300;color:#fff;" colspan="2" align="right"><%=FormatNumber(vBudgetBibit,0)%></th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">CODE</th>
		<th style="background-color:#00A300;color:#fff;">JENIS PEKERJAAN</th>
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
		
		<th style="background-color:#00A300;color:#fff;">Total</th>
		<th style="background-color:#00A300;color:#fff;">Rp/Pkk</th>
		<th style="background-color:#00A300;color:#fff;">Total</th>
		<th style="background-color:#00A300;color:#fff;">Rp/Pkk</th>
		<th style="background-color:#00A300;color:#fff;">Total</th>
		<th style="background-color:#00A300;color:#fff;">Rp/Pkk</th>
		<th style="background-color:#00A300;color:#fff;">Total</th>
		<th style="background-color:#00A300;color:#fff;">Rp/Pkk</th>
		<th style="background-color:#00A300;color:#fff;">Total</th>
		<th style="background-color:#00A300;color:#fff;">Rp/Pkk</th>
	</tr>
  </th>  
 <thead>
 </tbody>
	<%
Dim rsBibit, strBibit,codegroup,codegroupName,rp_pkk_bi,rp_pkk_sbi,rp_pkk_budbi,rp_pkk_budsbi,sub_rp_pkk_bi,sub_rp_pkk_sbi,sub_rp_pkk_budbi,sub_rp_pkk_budsbi,tot_rp_pkk_bi,tot_rp_pkk_sbi,tot_rp_pkk_budbi,tot_rp_pkk_budsbi
Dim subBI,subSBI,totsubBI,totsubSBI,budsubBI,budsubSBI,budSubthn,budtotsubBI,budtotsubSBI,budtotsubTHN
Dim strsql1,strsql2  

'
strsql1 = ""
strsql2 = ""

if  v_xBatch <> "--ALL--"  And v_xBatch <> "--SUMMARY--" then
        if	v_xBatch <> ""  then
		strsql1 = " and ca.targetcode ='" & v_xBatch & "'" 
		strsql2 = " and bp.batchcode  ='" & v_xBatch & "'"
		end if
end if
if v_txtjob <> "--SUMMARY--"  and v_txtjob <> "--ALL--" and v_txtjob <> "" then
	whereCons = wherecons & " And jobcode = '" & v_txtjob & "' "
end if
strBibit	=   " SELECT GroupPlantation " & selectCons & ",jobgroupcode,jobgroup_name,jobcode,Description,sum(BI) BI,sum(SBI) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn from ( " & _
                " select " & strFieldCons & " job_api.get_jobgroupCode(jobcode,fcba) jobgroupcode,jobgroup_api.get_jobgroup_name(job_api.get_jobgroupCode(jobcode,fcba),fcba)   jobgroup_name ,jobcode,job_api.get_jobcode_name(jobcode,fcba) Description,sum(BI) BI,sum(SBI) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn from  " & _
				"	( select ca.fcba,ca.jobcode,sum(ca.amount) BI,0 SBI,0 budBI,0 budSBI,0 budthn FROM COST_ACTIVITY ca " & _
				"	  where  instr('" & bibitjob & "',ca.jobcode) <> 0  and to_char(ca.tdate,'YYYY')='" & v_txtyear & "' and to_char(ca.tdate,'MM')=lpad('" & v_txtmonth  & "',2,0) " & strsql1 &  "  " & _
				"	  group by ca.fcba,ca.jobcode " & _
				"	  union " & _
				"	  select ca.fcba,ca.jobcode,0 BI,sum(ca.amount) SBI,0 budBI,0 budSBI,0 budthn FROM COST_ACTIVITY ca " & _
				"	  where  instr('" & bibitjob & "',ca.jobcode) <> 0   and to_char(ca.tdate,'YYYY')='" & v_txtyear & "' and to_char(ca.tdate,'MM') <= lpad('" & v_txtmonth  & "',2,0) " & strsql1 &  "  " & _
				"	  group by ca.fcba,ca.jobcode " & _
				"	  union " & _
				"	  select bp.fcba,bp.jobcode,0 BI,0 SBI,sum(" & strbud_BI & ") budBI,0 budSBI,0 budTHN from BUDGET_TRANSACTION_NURSERY bp  " & _
				"	  where bp.budget_year='" & v_txtyear & "' " & strsql2 &  "  " & _
				"	  group by bp.fcba,bp.jobcode " & _
				"	  union " & _
				"	  select bp.fcba,bp.jobcode,0 BI,0 SBI,0 budBI,sum(" & strbud_SBI & ") budSBI,0 budTHN from BUDGET_TRANSACTION_NURSERY bp  " & _
				"	  where bp.budget_year='" & v_txtyear & "'  " & strsql2 &  "  " & _
				"	  group by bp.fcba,bp.jobcode " & _
				"	  union " & _
				"	  select bp.fcba,bp.jobcode,0 BI,0 SBI,0 budBI,0 budSBI,sum(bp.budget_01+bp.budget_02+bp.budget_03+bp.budget_04+bp.budget_05+bp.budget_06+bp.budget_07+bp.budget_08+bp.budget_09+bp.budget_10+bp.budget_11+bp.budget_12) budTHN from BUDGET_TRANSACTION_NURSERY bp  " & _
				"	  where bp.budget_year='" & v_txtyear & "'  " & strsql2 &  "  " & _
				"	  group by bp.fcba,bp.jobcode " & _
				"	)  group by BUSINESSUNIT_API.get_OrderData(fcba),fcba,jobcode  ) " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E'  group by GroupPlantation,jobgroupcode,jobgroup_name,jobcode,Description " & selectCons & " order by jobgroupcode,jobcode" & orderCons  

'response.write strBibit
Set rsBibit = Server.CreateObject("ADODB.Recordset")
rsBibit.Open strBibit, connstr, adOpenStatic

if Not rsBibit.eof then
	While Not rsBibit.eof
			 ' cari rupiah per PKKw
			 IF cdbl(vActualBibit) <> 0 then
						rp_pkk_bi = cdbl(rsBibit("BI"))/cdbl(vActualBibit)
						rp_pkk_sbi = cdbl(rsBibit("SBI"))/cdbl(vActualBibit)
			 ELSE
						rp_pkk_bi  = 0
						rp_pkk_sbi = 0
			 END IF  
			 IF cdbl(vBudgetBibit) <> 0 then
				rp_pkk_budbi = cdbl(rsBibit("budBI"))/cdbl(vBudgetBibit)
				rp_pkk_budsbi = cdbl(rsBibit("budSBI"))/cdbl(vBudgetBibit)
			 ELSE
				rp_pkk_budbi  = 0
				rp_pkk_budsbi = 0
			 END IF 
	 
		If 	rsBibit("jobgroupcode") = codegroup Then
		%>
		           <tr>
						<% 
								tempjob = rsBibit("jobcode") 
								tempPlantation = "--ALL--"
								tempGroupRegion = "--SUMMARY--"
								tempRegion = "--SUMMARY--"
								tempCompany ="--SUMMARY--"
								tempfcba = "--SUMMARY--"
								tempDivision = "--SUMMARY--"
								tempPlantingDate = "--SUMMARY--"
							    If SprevCode <> rsBibit("jobcode") then %>
									 <td align="center">
										<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
										<%=rsBibit("jobcode")%>
										</a>
									 </td>
									 <% SprevCode   = rsBibit("jobcode") %>
							   <% Else %>
									 <td align="center"></td>
							   <% End If %>		
                          							   
							<% If Sprevjobname <> rsBibit("Description") then %>
								 <td><%=rsBibit("Description")%></td>
								 <% Sprevjobname = rsBibit("Description") %>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>	
						   
						
						<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
						             if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
									    tempjob = rsBibit("jobcode")
									 else
									    tempjob = ""
                                     end if									 
									 tempPlantation = rsBibit("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevPlantation <> rsBibit("Plantation") then %>
										   <td align="center">																																																																																									
												<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
												<%=rsBibit("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsBibit("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
									 if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
									    tempjob = rsBibit("jobcode")
									 else
									    tempjob = ""
                                     end if	
									 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										  tempPlantation = rsBibit("Plantation")
									 else
										  tempPlantation = ""
									 end if
									 tempGroupRegion = rsBibit("GroupRegion")
									 tempRegion = "--ALL--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevGroupRegion <> rsBibit("GroupRegion") then %>
										   <td align="center">
												<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
												<%=rsBibit("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsBibit("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsBibit("jobcode")
										 else
											tempjob = ""
										 end if	
									   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
										  tempPlantation = rsBibit("Plantation")
									   else
										  tempPlantation = ""
									   end if
									   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
										   tempGroupRegion = rsBibit("GroupRegion")
									   else
										   tempGroupRegion = ""
									   end if
									   
									   tempRegion = rsBibit("Region") 
									   tempCompany ="--ALL--"
									   tempfcba = "--SUMMARY--"
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevRegion <> rsBibit("Region") then %>
											   <td align="center">
													<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsBibit("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsBibit("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsBibit("jobcode")
										 else
											tempjob = ""
										 end if	
									   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
										  tempPlantation = rsBibit("Plantation")
									   else
										  tempPlantation = ""
									   end if
									   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
										   tempGroupRegion = rsBibit("GroupRegion")
									   else
										   tempGroupRegion = ""
									   end if
									   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
										   tempRegion = rsBibit("Region") 
									   else
										   tempRegion = ""
									   end if								   
									   tempCompany =rsBibit("Company") 
									   tempfcba = "--ALL--"
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevCompany <> rsBibit("Company") then %>
											   <td align="center">
												    <a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													  <%=rsBibit("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsBibit("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsBibit("jobcode")
										 else
											tempjob = ""
										 end if	
									   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
										  tempPlantation = rsBibit("Plantation")
									   else
										  tempPlantation = ""
									   end if
									   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
										   tempGroupRegion = rsBibit("GroupRegion")
									   else
										   tempGroupRegion = ""
									   end if
									   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
										   tempRegion = rsBibit("Region") 
									   else
										   tempRegion = ""
									   end if
									   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
										   tempCompany =rsBibit("Company")
									   else
										   tempCompany =""
									   end if							   
									   tempfcba = rsBibit("fcba")
									   tempDivision = "--ALL--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF Sprevfcba <> rsBibit("fcba") then %>
											   <td align="center">
													<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsBibit("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsBibit("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
						<td align="right"><%=FormatNumber(rsBibit("BI"),0)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_bi),0)%></td>
						<td align="right"><%=FormatNumber(rsBibit("budBI"),0)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_budbi),0)%></td>
						<% IF cdbl(rp_pkk_budbi) <> 0 and cdbl(rp_pkk_bi)=0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(rp_pkk_budbi) = 0 and cdbl(rp_pkk_bi)<> 0 then %>  
						  <td align="right">100</td>
						<% ELSEIF cdbl(rp_pkk_budbi) = 0 and cdbl(rp_pkk_bi) = 0 then %>  
						   <td align="right">0</td>
						<% ELSEIF cdbl(rp_pkk_budbi) <> 0 and cdbl(rp_pkk_bi) <> 0 then %>  
						    <td align="right"><%=FormatNumber((cdbl(rp_pkk_budbi)-cdbl(rp_pkk_bi))/cdbl(rp_pkk_budbi)*100,0)%></td>
						<% END IF %>  
						<td align="right"><%=FormatNumber(rsBibit("SBI"),0)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_sbi),0)%></td>
						<td align="right"><%=FormatNumber(rsBibit("budBI"),0)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_budsbi),0)%></td>
						<% IF cdbl(rp_pkk_budsbi) <> 0 and cdbl(rp_pkk_sbi)=0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) = 0 and cdbl(rp_pkk_sbi)<> 0 then %>  
						  <td align="right">100</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) = 0 and cdbl(rp_pkk_sbi) = 0 then %>  
						   <td align="right">0</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) <> 0 and cdbl(rp_pkk_sbi) <> 0 then %>  
						    <td align="right"><%=FormatNumber((cdbl(rp_pkk_budsbi)-cdbl(rp_pkk_sbi))/cdbl(rp_pkk_budsbi)*100,0)%></td>
						<% END IF %> 
						<td align="right"><%=FormatNumber(rsBibit("budTHN"),0)%></td>
						<td align="right"><%=FormatNumber(0,0)%></td>
					 </tr> 
	<%
		        subBI     = subBI + Cdbl(rsBibit("BI"))
				subSBI    = subSBI + Cdbl(rsBibit("SBI"))
				budsubBI  = budsubBI + Cdbl(rsBibit("budBI"))
				budsubSBI = budsubSBI + Cdbl(rsBibit("budSBI"))
				budSubthn = budSubthn + Cdbl(rsBibit("budTHN")) 
		Else  %>
	                <% IF codegroup <> "" Then 
					      ' cari rupiah per PKK
							 IF cdbl(vActualBibit) <> 0 then
										sub_rp_pkk_bi = cdbl(subBI)/cdbl(vActualBibit)
										sub_rp_pkk_sbi = cdbl(subSBI)/cdbl(vActualBibit)
							 ELSE
										sub_rp_pkk_bi  = 0
										sub_rp_pkk_sbi = 0
							 END IF  
							 IF cdbl(vBudgetBibit) <> 0 then
								sub_rp_pkk_budbi = cdbl(budsubBI)/cdbl(vBudgetBibit)
								sub_rp_pkk_budsbi = cdbl(budsubSBI)/cdbl(vBudgetBibit)
							 ELSE
								sub_rp_pkk_budbi  = 0
								sub_rp_pkk_budsbi = 0
							 END IF 
										%>
						<tr bgcolor="#F4F4A0">
							<td colspan="<%= 2 + nSpanTotal %>"  align="left"><strong>Total <%=codegroupName%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(subBI),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_bi),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(budsubBI),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_budbi),0)%></strong></td>
							<% IF cdbl(sub_rp_pkk_budbi) = 0 and cdbl(sub_rp_pkk_bi) <> 0  then %>
							      <td align="right"><strong>-100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) <> 0 and cdbl(sub_rp_pkk_bi) = 0  then %>
							      <td align="right"><strong>100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) = 0 and cdbl(sub_rp_pkk_bi) = 0  then %>
							      <td align="right"><strong>0</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) <> 0 and cdbl(sub_rp_pkk_bi) <> 0  then %>
							      <td align="right"><strong><%=FormatNumber((cdbl(sub_rp_pkk_budbi)-cdbl(sub_rp_pkk_bi))/cdbl(sub_rp_pkk_budbi)*100,0)%></strong></td>
							<% END IF %>	  
							<td align="right"><strong><%=FormatNumber(cdbl(subSBI),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_sbi),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(budsubSBI),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_budsbi),0)%></strong></td>
							<% IF cdbl(sub_rp_pkk_budsbi) = 0 and cdbl(sub_rp_pkk_sbi) <> 0  then %>
							      <td align="right"><strong>-100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) <> 0 and cdbl(sub_rp_pkk_sbi) = 0  then %>
							      <td align="right"><strong>100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) = 0 and cdbl(sub_rp_pkk_sbi) = 0  then %>
							      <td align="right"><strong>0</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) <> 0 and cdbl(sub_rp_pkk_sbi) <> 0  then %>
							      <td align="right"><strong><%=FormatNumber((cdbl(sub_rp_pkk_budsbi)-cdbl(sub_rp_pkk_sbi))/cdbl(sub_rp_pkk_budsbi)*100,0)%></strong></td>
							<% END IF %>	
							<td align="right"><strong><%=FormatNumber(cdbl(budSubthn),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
						</tr>
					<% END IF %> 
					   
					 <tr>
						 <% 
								tempjob = rsBibit("jobcode") 
								tempPlantation = "--ALL--"
								tempGroupRegion = "--SUMMARY--"
								tempRegion = "--SUMMARY--"
								tempCompany ="--SUMMARY--"
								tempfcba = "--SUMMARY--"
								tempDivision = "--SUMMARY--"
								tempPlantingDate = "--SUMMARY--"
							    If SprevCode <> rsBibit("jobcode") then %>
									 <td align="center">
										<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
										<%=rsBibit("jobcode")%>
										</a>
									 </td>
									 <% SprevCode   = rsBibit("jobcode") %>
							   <% Else %>
									 <td align="center"></td>
							   <% End If %>		
                          							   
							<% If Sprevjobname <> rsBibit("Description") then %>
								 <td><%=rsBibit("Description")%></td>
								 <% Sprevjobname = rsBibit("Description") %>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>	
						   
						
						<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
						             if v_txtjob <> "----" and  v_txtjob <> ""  then
									    tempjob = rsBibit("jobcode")
									 else
									    tempjob = ""
                                     end if									 
									 tempPlantation = rsBibit("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevPlantation <> rsBibit("Plantation") then %>
										   <td align="center">																																																																																									
												<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
												<%=rsBibit("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsBibit("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
									 if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
									    tempjob = rsBibit("jobcode")
									 else
									    tempjob = ""
                                     end if	
									 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										  tempPlantation = rsBibit("Plantation")
									 else
										  tempPlantation = ""
									 end if
									 tempGroupRegion = rsBibit("GroupRegion")
									 tempRegion = "--ALL--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevGroupRegion <> rsBibit("GroupRegion") then %>
										   <td align="center">
												<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
												<%=rsBibit("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsBibit("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsBibit("jobcode")
										 else
											tempjob = ""
										 end if	
									   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
										  tempPlantation = rsBibit("Plantation")
									   else
										  tempPlantation = ""
									   end if
									   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
										   tempGroupRegion = rsBibit("GroupRegion")
									   else
										   tempGroupRegion = ""
									   end if
									   
									   tempRegion = rsBibit("Region") 
									   tempCompany ="--ALL--"
									   tempfcba = "--SUMMARY--"
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevRegion <> rsBibit("Region") then %>
											   <td align="center">
													<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsBibit("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsBibit("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsBibit("jobcode")
										 else
											tempjob = ""
										 end if	
									   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
										  tempPlantation = rsBibit("Plantation")
									   else
										  tempPlantation = ""
									   end if
									   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
										   tempGroupRegion = rsBibit("GroupRegion")
									   else
										   tempGroupRegion = ""
									   end if
									   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
										   tempRegion = rsBibit("Region") 
									   else
										   tempRegion = ""
									   end if								   
									   tempCompany =rsBibit("Company") 
									   tempfcba = "--ALL--"
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevCompany <> rsBibit("Company") then %>
											   <td align="center">
												    <a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													  <%=rsBibit("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsBibit("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsBibit("jobcode")
										 else
											tempjob = ""
										 end if	
									   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
										  tempPlantation = rsBibit("Plantation")
									   else
										  tempPlantation = ""
									   end if
									   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
										   tempGroupRegion = rsBibit("GroupRegion")
									   else
										   tempGroupRegion = ""
									   end if
									   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
										   tempRegion = rsBibit("Region") 
									   else
										   tempRegion = ""
									   end if
									   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
										   tempCompany =rsBibit("Company")
									   else
										   tempCompany =""
									   end if							   
									   tempfcba = rsBibit("fcba")
									   tempDivision = "--ALL--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF Sprevfcba <> rsBibit("fcba") then %>
											   <td align="center">
													<a href="p_AnalisaBibit.Asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsBibit("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsBibit("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
						<td align="right"><%=FormatNumber(rsBibit("BI"),0)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_bi),0)%></td>
						<td align="right"><%=FormatNumber(rsBibit("budBI"),0)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_budbi),0)%></td>
						<% IF cdbl(rp_pkk_budbi) <> 0 and cdbl(rp_pkk_bi)=0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(rp_pkk_budbi) = 0 and cdbl(rp_pkk_bi)<> 0 then %>  
						  <td align="right">100</td>
						<% ELSEIF cdbl(rp_pkk_budbi) = 0 and cdbl(rp_pkk_bi) = 0 then %>  
						   <td align="right">0</td>
						<% ELSEIF cdbl(rp_pkk_budbi) <> 0 and cdbl(rp_pkk_bi) <> 0 then %>  
						    <td align="right"><%=FormatNumber((cdbl(rp_pkk_budbi)-cdbl(rp_pkk_bi))/cdbl(rp_pkk_budbi)*100,0)%></td>
						<% END IF %>  
						<td align="right"><%=FormatNumber(rsBibit("SBI"),0)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_sbi),0)%></td>
						<td align="right"><%=FormatNumber(rsBibit("budBI"),0)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_budsbi),0)%></td>
						<% IF cdbl(rp_pkk_budsbi) <> 0 and cdbl(rp_pkk_sbi)=0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) = 0 and cdbl(rp_pkk_sbi)<> 0 then %>  
						  <td align="right">100</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) = 0 and cdbl(rp_pkk_sbi) = 0 then %>  
						   <td align="right">0</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) <> 0 and cdbl(rp_pkk_sbi) <> 0 then %>  
						    <td align="right"><%=FormatNumber((cdbl(rp_pkk_budsbi)-cdbl(rp_pkk_sbi))/cdbl(rp_pkk_budsbi)*100,0)%></td>
						<% END IF %> 
						<td align="right"><%=FormatNumber(rsBibit("budTHN"),0)%></td>
						<td align="right"><%=FormatNumber(0,0)%></td>
					 </tr> 
					  <%
							    subBI     =  Cdbl(rsBibit("BI"))
								subSBI    =  Cdbl(rsBibit("SBI"))
								budsubBI  =  Cdbl(rsBibit("budBI"))
								budsubSBI = Cdbl(rsBibit("budSBI"))
								budSubthn =  Cdbl(rsBibit("budTHN")) 
					  %>
	<%
		End If	
	
	    
		totsubBI  = totsubBI + Cdbl(rsBibit("BI"))
		totsubSBI = totsubSBI + Cdbl(rsBibit("SBI"))
		
		budtotsubBI = budtotsubBI + Cdbl(rsBibit("budBI"))
		budtotsubSBI = budtotsubSBI + Cdbl(rsBibit("budSBI"))
		budtotsubTHN = budtotsubTHN + Cdbl(rsBibit("budTHN"))
		
		codegroup = rsBibit("jobgroupcode")
		codegroupName = rsBibit("jobgroup_name")
		
		rsBibit.MoveNext
	Wend
end if
	%>
      <%	      
		                   IF cdbl(vActualBibit) <> 0 then
										sub_rp_pkk_bi = cdbl(subBI)/cdbl(vActualBibit)
										sub_rp_pkk_sbi = cdbl(subSBI)/cdbl(vActualBibit)
							 ELSE
										sub_rp_pkk_bi  = 0
										sub_rp_pkk_sbi = 0
							 END IF  
							 IF cdbl(vBudgetBibit) <> 0 then
								sub_rp_pkk_budbi = cdbl(budsubBI)/cdbl(vBudgetBibit)
								sub_rp_pkk_budsbi = cdbl(budsubSBI)/cdbl(vBudgetBibit)
							 ELSE
								sub_rp_pkk_budbi  = 0
								sub_rp_pkk_budsbi = 0
							 END IF 
										%>
	               <tr bgcolor="#F4F4A0">
				            <% if  v_txtjob = "--ALL--"  then %>
								   <td colspan="<%= 2 + nSpanTotal %>" align="left"><strong>Total  <%=codegroupName%></strong></td>
							
							<td align="right"><strong><%=FormatNumber(cdbl(subBI),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_bi),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(budsubBI),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_budbi),0)%></strong></td>
							<% IF cdbl(sub_rp_pkk_budbi) = 0 and cdbl(sub_rp_pkk_bi) <> 0  then %>
							      <td align="right"><strong>-100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) <> 0 and cdbl(sub_rp_pkk_bi) = 0  then %>
							      <td align="right"><strong>100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) = 0 and cdbl(sub_rp_pkk_bi) = 0  then %>
							      <td align="right"><strong>0</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) <> 0 and cdbl(sub_rp_pkk_bi) <> 0  then %>
							      <td align="right"><strong><%=FormatNumber((cdbl(sub_rp_pkk_budbi)-cdbl(sub_rp_pkk_bi))/cdbl(sub_rp_pkk_budbi)*100,0)%></strong></td>
							<% END IF %>	  
							<td align="right"><strong><%=FormatNumber(cdbl(subSBI),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_sbi),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(budsubSBI),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_budsbi),0)%></strong></td>
							<% IF cdbl(sub_rp_pkk_budsbi) = 0 and cdbl(sub_rp_pkk_sbi) <> 0  then %>
							      <td align="right"><strong>-100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) <> 0 and cdbl(sub_rp_pkk_sbi) = 0  then %>
							      <td align="right"><strong>100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) = 0 and cdbl(sub_rp_pkk_sbi) = 0  then %>
							      <td align="right"><strong>0</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) <> 0 and cdbl(sub_rp_pkk_sbi) <> 0  then %>
							      <td align="right"><strong><%=FormatNumber((cdbl(sub_rp_pkk_budsbi)-cdbl(sub_rp_pkk_sbi))/cdbl(sub_rp_pkk_budsbi)*100,0)%></strong></td>
							<% END IF %>	
							<td align="right"><strong><%=FormatNumber(cdbl(budSubthn),0)%></strong></td>
							<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
							<% End if %>	
						</tr>
						<%
		                    ' cari rupiah per PKK
							 IF cdbl(vActualBibit) <> 0 then
										tot_rp_pkk_bi = cdbl(totsubBI)/cdbl(vActualBibit)
										tot_rp_pkk_sbi = cdbl(totsubSBI)/cdbl(vActualBibit)
							 ELSE
										tot_rp_pkk_bi  = 0
										tot_rp_pkk_sbi = 0
							 END IF  
							 IF cdbl(vBudgetBibit) <> 0 then
								tot_rp_pkk_budbi = cdbl(budtotsubBI)/cdbl(vBudgetBibit)
								tot_rp_pkk_budsbi = cdbl(budtotsubSBI)/cdbl(vBudgetBibit)
							 ELSE
								tot_rp_pkk_budbi  = 0
								tot_rp_pkk_budsbi = 0
							 END IF 
										%>
		<tr bgcolor="#F4F4A0">
		    <% if  v_txtjob <> "--ALL--"  then %>
			        <td colspan="<%= 2 + nSpanTotal %>" align="left"><strong>Total</strong></td>
			<% Else %>
			       <td colspan="<%= 2 + nSpanTotal %>" align="left"><strong>Total Biaya Pembibitan</strong></td>
		    <% End if %>		   
			<td align="right"><strong><%=FormatNumber(cdbl(totsubBI),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_pkk_bi),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(budtotsubBI),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_pkk_budbi),0)%></strong></td>
			<% IF cdbl(tot_rp_pkk_budbi) = 0 and cdbl(tot_rp_pkk_bi) <> 0  then %>
				<td align="right"><strong>-100</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budbi) <> 0 and cdbl(tot_rp_pkk_bi) = 0  then %>
				<td align="right"><strong>100</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budbi) = 0 and cdbl(tot_rp_pkk_bi) = 0  then %>
				<td align="right"><strong>0</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budbi) <> 0 and cdbl(tot_rp_pkk_bi) <> 0  then %>
				<td align="right"><strong><%=FormatNumber((cdbl(tot_rp_pkk_budbi)-cdbl(tot_rp_pkk_bi))/cdbl(tot_rp_pkk_budbi)*100,0)%></strong></td>
			<% END IF %>
			<td align="right"><strong><%=FormatNumber(cdbl(totsubSBI),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_pkk_sbi),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(budtotsubSBI),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_pkk_budsbi),0)%></strong></td>
			<% IF cdbl(tot_rp_pkk_budsbi) = 0 and cdbl(tot_rp_pkk_sbi) <> 0  then %>
				<td align="right"><strong>-100</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budsbi) <> 0 and cdbl(tot_rp_pkk_sbi) = 0  then %>
				<td align="right"><strong>100</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budsbi) = 0 and cdbl(tot_rp_pkk_sbi) = 0  then %>
				<td align="right"><strong>0</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budsbi) <> 0 and cdbl(tot_rp_pkk_sbi) <> 0  then %>
				<td align="right"><strong><%=FormatNumber((cdbl(tot_rp_pkk_budsbi)-cdbl(tot_rp_pkk_sbi))/cdbl(tot_rp_pkk_budsbi)*100,0)%></strong></td>
			<% END IF %>
			<td align="right"><strong><%=FormatNumber(cdbl(budtotsubTHN),0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		</tr>
	</thead>
</table>
<%


%>
</div>
</body>
</html>