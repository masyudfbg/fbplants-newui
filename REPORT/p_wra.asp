<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal
Dim vyear,vmonth,v_vJenisReport,v_txtJob,v_TxtTypeFCBA 
Dim SprevCode,Sprevjobname,tempjob
Dim sprevFcba, SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation,SprevVH
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempVH
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType,rsJob
            v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	  = Request("txtPlantation")
			v_txtGroupRegion  	  = Request("txtGroupRegion")
			v_txtRegion  		  = Request("txtRegion")
			v_txtcompany  		  = Request("txtCompany")
			v_txtfcba 			  = Request("txtfcba")
			vyear             = Request("vYear")
			vmonth            = Request("vMonth")
            v_txtjob			  = Request("txtjob")
			v_TxtTypeFCBA         = Request("TxtTypeFCBA")
			
			
            nSpanTotal = 5 
			
			Dim selectFields,strFieldCons,selectCons,whereCons,orderCons 
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, BUSINESSUNIT_API.get_OrderData(fcba) norut, fcba, " 

						   
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
	       
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Laporan Workshop Running Account (WRA)</title>
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
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<form name="frmEntry" method="post" action="p_wra.asp">
	<input type="hidden" name="TxtTypeFCBA" value="<%=v_TxtTypeFCBA%>" >
	<!--#include file="FormSelectNewSumME.html"-->
	<!--#include file="FromSelectHDRWRA.html"-->
	</form>
	<div id="header">
		<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
		 <% End if%>
		 <% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
		 <% End if%>
	</div>
		<div id="subheader" class="centered">
			<% Dim ia
			   ia = request("vMonth")
			%>   
			<h3>LAPORAN WORKSHOP RUNNING ACCOUNT (WRA)</h3>
			<% IF request("vMonth") <> "" then %>
			<h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=request("vYear")%></h5>
			<% END IF %>
			
		</div>
	</div>
<%
  Dim cur_month,strbud_BI,strbud_SBI,x1,strbudprod_BI,strbudprod_SBI
	    strbud_BI = "t.Budget_01" 
	    strbud_SBI = "t.Budget_01"
		strbudprod_BI = "t.qty_01"
		strbudprod_SBI = "t.qty_01"
		
		cur_month = request("vMonth") 
		
		IF Len(cur_month) > 0 then

			strbud_BI = "" 
	        strbud_SBI = ""
			strbudprod_BI = ""
		    strbudprod_SBI = ""
			x1=1
			
			Do While x1<= cint(cur_month) 'Output the values from 0 to 10
			  'response.write(i & "<br \>")
			  'response.write  " Budget_" & Lpad(x1,"0",2) & " +"
			  strbud_SBI = strbud_SBI & " t.Budget_" & Lpad(x1,"0",2) & " +"
		      strbudprod_SBI = strbudprod_SBI &  "t.qty_" & Lpad(x1,"0",2) & " +"
			  x1 = x1 + 1 'increment the value of i for next time loop executes
			Loop

			strbud_BI = "t.Budget_" & Lpad(cur_month,"0",2)
			strbud_SBI = Mid(strbud_SBI,1,Len(strbud_SBI)-1)
			strbudprod_BI = "t.qty_" & Lpad(cur_month,"0",2)
			strbudprod_SBI = Mid(strbudprod_SBI,1,Len(strbudprod_SBI)-1)
            'response.write  strbud_SBI	& "<br \>"		
	        'response.write  strbud_BI	& "<br \>"	
		END IF
		
  Dim rshdr, strhdr ,jamBI,jamSBI,jambudBI,jambudSBI,jambudthn
   jamBI = "0"
   jamSBI = "0"
  strhdr  = " select " & strFieldCons & " sum(JK_BI) JK_BI,sum(JK_SBI)  JK_SBI,sum(JK_bud_BI) JK_bud_BI,sum(JK_bud_SBI) JK_bud_SBI,sum(JK_bud_THN) JK_bud_THN from ( " & _
			"	select wa.fcba,sum(wa.unit) JK_BI,0 JK_SBI,0 JK_bud_BI,0 JK_bud_SBI,0 JK_bud_THN from workshopactivity wa  " & _
			"	where to_char(wa.workshopdate,'YYYY')='" & request("vYear") & "' and to_char(wa.workshopdate,'MM')=lpad('" & request("vMonth") & "',2,0) " & _
			"	group by wa.fcba " & _
			"	union " & _
			"	select wa.fcba,0 JK_BI,sum(wa.unit)  JK_SBI,0 JK_bud_BI,0 JK_bud_SBI,0 JK_bud_THN from workshopactivity wa   " & _
			"	where to_char(wa.workshopdate,'YYYY')='" & request("vYear") & "' and to_char(wa.workshopdate,'MM')<=lpad('" & request("vMonth") & "',2,0) " & _
			"	group by wa.fcba " & _
			"	union " & _
			"	select t.fcba,0 JK_BI,0  JK_SBI,sum(" & strbudprod_BI & ") JK_bud_BI,0 JK_bud_SBI,0 JK_bud_THN from BUDGET_TRANSACTION_WSHOUR t " & _
			"	where  t.budget_year='" & request("vYear") & "' " & _
			"	group by t.fcba " & _
			"	union " & _
			"	select t.fcba,0 JK_BI,0  JK_SBI,0 JK_bud_BI,sum(" & strbudprod_SBI & ") JK_bud_SBI,0 JK_bud_THN from BUDGET_TRANSACTION_WSHOUR t " & _
			"	where  t.budget_year='" & request("vYear") & "' " & _
			"	group by t.fcba " & _
			"	union " & _
			"	select t.fcba,0 JK_BI,0  JK_SBI,0 JK_bud_BI,0 JK_bud_SBI,sum(t.qty_01+t.qty_02+t.qty_03+t.qty_04+t.qty_05+t.qty_06+t.qty_07+t.qty_08+t.qty_09+t.qty_10+t.qty_11+t.qty_12) JK_bud_THN from BUDGET_TRANSACTION_WSHOUR t " & _
			"	where  t.budget_year='" & request("vYear") & "' " & _
			"	group by t.fcba " & _
			"	 )  where USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = '" & Request("TxtTypeFCBA") &"'   group by fcba "
			
			strhdr = "select GroupPlantation ,sum(JK_BI) JK_BI,sum(JK_SBI)  JK_SBI,sum(JK_bud_BI) JK_bud_BI,sum(JK_bud_SBI) JK_bud_SBI,sum(JK_bud_THN) JK_bud_THN  from ( " & strhdr & " )" & whereCons & "  group by  GroupPlantation  " 
'response.write strhdr
Set rshdr = Server.CreateObject("ADODB.Recordset") 
'response.write strhdr
rshdr.Open strhdr, connstr, adOpenStatic
	
if Not rshdr.eof then
   jamBI = rshdr("JK_BI")
   jamSBI = rshdr("JK_SBI")
   jambudBI = rshdr("JK_bud_BI")
   jambudSBI = rshdr("JK_bud_SBI")
   jambudthn = rshdr("JK_bud_THN")
end if
rshdr.close

%>

<table border="0" width="100%" id="product-table">
	<tr>
		<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" rowspan="2"></th>
		<th  style="background-color:#00A300;color:#fff;" colspan="5" >BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="5" >S/D BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="2"rowspan="2" >ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
	    <th  style="background-color:#00A300;color:#fff;" colspan="2">REALISASI</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="2">ANGGARAN</th>
		<th  style="background-color:#00A300;color:#fff;" >%</th>
		
		<th  style="background-color:#00A300;color:#fff;" colspan="2">REALISASI</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="2">ANGGARAN</th>
		<th  style="background-color:#00A300;color:#fff;" >%</th>
	</tr>

	<tr>
		<th style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>">JAM KERJA BENGKEL (JAM)</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2"><%=formatnumber(cdbl(jamBI),2)%></th>
		<th style="background-color:#00A300;color:#fff;" colspan="2"><%=formatnumber(cdbl(jambudBI),2)%></th>
		<% IF cdbl(jamBI) = 0 THEN %>
		    <th style="background-color:#00A300;color:#fff;" >100</th>
		<% ELSEIF cdbl(jambudBI) = 0 THEN %>
		    <th style="background-color:#00A300;color:#fff;" >-100</th>
		<% ELSEIF cdbl(jambudBI) = 0 and cdbl(jamBI) = 0 THEN %>
		    <th style="background-color:#00A300;color:#fff;" >0</th>
        <% ELSEIF cdbl(jambudBI) > 0 and cdbl(jamBI) > 0 THEN %>
		    <th style="background-color:#00A300;color:#fff;" ><%=formatnumber((cdbl(jamBI) - cdbl(jambudBI))/cdbl(jambudBI)*100,,,-1)%></th>
        <% END IF %>			
		<th style="background-color:#00A300;color:#fff;" colspan="2"><%=formatnumber(cdbl(jamSBI),2)%></th>
		<th style="background-color:#00A300;color:#fff;" colspan="2"><%=formatnumber(cdbl(jambudSBI),2)%></th>
		<% IF cdbl(jamSBI) = 0 THEN %>
		    <th style="background-color:#00A300;color:#fff;" >100</th>
		<% ELSEIF cdbl(jambudSBI) = 0 THEN %>
		    <th style="background-color:#00A300;color:#fff;" >-100</th>
		<% ELSEIF cdbl(jambudSBI) = 0 and cdbl(jamSBI) = 0 THEN %>
		    <th style="background-color:#00A300;color:#fff;" >0</th>
        <% ELSEIF cdbl(jambudSBI) > 0 and cdbl(jamSBI) > 0 THEN %>
		    <th style="background-color:#00A300;color:#fff;" ><%=formatnumber((cdbl(jamSBI) - cdbl(jambudSBI))/cdbl(jambudSBI)*100,,,-1)%></th>
        <% END IF %>
		
		<th style="background-color:#00A300;color:#fff;" colspan="2"><%=formatnumber(cdbl(jambudthn),2)%></th>
	</tr>
	<tr>
		
		<th style="background-color:#00A300;color:#fff;" >KODE</th>
		<th style="background-color:#00A300;color:#fff;" >NAMA PEKERJAAN</th>
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
		
		<% 
		   if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;" >SITE</th>			
	    <% end if %>
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Jam</th>		
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Jam</th>
		<th style="background-color:#00A300;color:#fff;" ></th>
		
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Jam</th>		
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Jam</th>
		<th style="background-color:#00A300;color:#fff;" ></th>
		
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Jam</th>
		
		
	</tr>
	<%
	Dim rsha, strdtl ,rpperjamBI,rpperjamSBI,rpperjambudBI,rpperjambudSBI,rpperjambudTHN
	Dim subtotBI, subtotSBI,subrpperjamBI,subrpperjamSBI, subtotbudBI, subtotbudSBI,subrpperjambudBI,subrpperjambudSBI,subtotbudTHN,subrpperjambudTHN
	Dim subrasioBI,subactualBI ,subactualSBI ,subtotactualBI ,subtotactualSBI 
	
	      if v_txtjob <> "--SUMMARY--"  and v_txtjob <> "--ALL--" and v_txtjob <> "" then
				whereCons = wherecons & " And jobcode = '" & v_txtjob & "' "
			end if
	
			strdtl = " select " & strFieldCons & " jobcode,job_api.get_jobcode_name(jobcode,fcba) JobName, sum(totBI) totBI,sum(totSBI) totSBI,sum(totbudBI) totbudBI,sum(totbudSBI) totbudSBI ,sum(totbudTHN) totbudTHN  from ( " & _
					 "	select ca.fcba,ca.jobcode, sum(ca.amountdebit)+Sum(ca.amountkredit) totBI,0 totSBI,0 totbudBI,0 totbudSBI,0 totbudTHN  from bkmjournaldtl ca INNER JOIN bkmjournalhdr hdr ON (ca.docno=hdr.docno AND ca.fcba=hdr.fcba) " & _
					 "	where ca.locationtype='WS' and to_char(hdr.docdate,'YYYY')='" & request("vYear") & "' and to_char(hdr.docdate,'MM')=lpad('" & request("vMonth") & "',2,0) " & _
					 "	group by ca.fcba,ca.jobcode " & _
					 "	union " & _
					 "	select ca.fcba,ca.jobcode, 0 totBI,sum(ca.amountdebit)+Sum(ca.amountkredit) totSBI,0 totbudBI,0 totbudSBI,0 totbudTHN  from bkmjournaldtl ca INNER JOIN bkmjournalhdr hdr ON (ca.docno=hdr.docno AND ca.fcba=hdr.fcba) " & _
					 "	where ca.locationtype='WS' and to_char(hdr.docdate,'YYYY')='" & request("vYear")  & "' and to_char(hdr.docdate,'MM')<=lpad('" & request("vMonth") & "',2,0)  " & _
					 "	group by ca.fcba,ca.jobcode  " & _
					 "	union " & _
					 "	select t.fcba,t.jobcode ,0 totBI,0 totSBI,sum( " & strbud_BI & ") totbudBI,0 totbudSBI,0 totbudTHN from BUDGET_TRANSACTION_WORKSHOP t " & _
					 "	where t.budget_year='" & request("vYear") & "'  " & _
					 "	group by t.fcba,t.jobcode " & _
					 "	union " & _
					 "	select t.fcba,t.jobcode ,0 totBI,0 totSBI,0 totbudBI,sum(" & strbud_SBI & ") totbudSBI,0 totbudTHN from BUDGET_TRANSACTION_WORKSHOP t " & _
					 "	where t.budget_year='" & request("vYear")  & "'  " & _
					 "	group by t.fcba,t.jobcode " & _
					 "	union " & _
					 "	select t.fcba,t.jobcode ,0 totBI,0 totSBI,0 totbudBI,0 totbudSBI,sum(nvl(t.budget_01,0)+nvl(t.budget_02,0)+nvl(t.budget_03,0)+nvl(t.budget_04,0)+nvl(t.budget_05,0)+nvl(t.budget_06,0)+nvl(t.budget_07,0)+nvl(t.budget_08,0)+nvl(t.budget_09,0)+nvl(t.budget_10,0)+nvl(t.budget_11,0)+nvl(t.budget_12,0)) totbudTHN from BUDGET_TRANSACTION_WORKSHOP t " & _
					 "	where t.budget_year='" & request("vYear")  & "' " & _
					 "	group by t.fcba,t.jobcode " & _
					 "	) where USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = '" & Request("TxtTypeFCBA") &"'  group by fcba,jobcode "	
		
		strdtl = "select jobcode,JobName,GroupPlantation " & selectCons & " ,sum(totBI) totBI,sum(totSBI) totSBI,sum(totbudBI) totbudBI,sum(totbudSBI) totbudSBI ,sum(totbudTHN) totbudTHN from ( " & strdtl & " )" & whereCons & "  group by  jobcode,JobName,GroupPlantation  " & selectCons & "  Order by jobcode,JobName,GroupPlantation " & orderCons  

	
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
	'response.write strdtl
    'response.write request("vMonth") &  "  " & request("vYear") & " " & session("fcba") 
	if Not rsha.eof then
		  while Not rsha.eof
		    
			  %>
				<tr>  
					<%   tempjob = rsha("jobcode") 
					     tempPlantation = "--ALL--"
						 tempGroupRegion = "--SUMMARY--"
						 tempRegion = "--SUMMARY--"
						 tempCompany ="--SUMMARY--"
						 tempfcba = "--SUMMARY--"
						 tempVH ="--SUMMARY--"
					   If SprevCode <> rsha("jobcode") then %>
					      <td align="CENTER">
						         <a href="p_wra.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">				
						         <%=rsha("jobcode")%>
								 </a>
								<% SprevCode   = rsha("jobcode") %>
						  </td>
				     <% Else %>
								 <td align="center"></td>
					 <% End If %>		 		  
					     <% If Sprevjobname <> rsha("jobname") then %>
								 <td><%=rsha("jobname")%></td>
								<% Sprevjobname = rsha("jobname") %>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>	
					<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     
						   tempPlantation = rsha("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempVH ="--SUMMARY--"
					%>
					<% if SprevPlantation <> rsha("Plantation")   then %>
						<td align="center">
							<a href="p_wra.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsha("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  
					    
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsha("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsha("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempVH ="--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsha("GroupRegion") then %>
						<td align="center">
							<a href="p_wra.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsha("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsha("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsha("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsha("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempVH = "--SUMMARY--"
					   
					%>
					
					<% if SprevRegion <> rsha("Region") then %>
						<td align="center">
							<a href="p_wra.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("region")%>
							</a>
						</td>
						<% SprevRegion = rsha("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsha("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsha("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsha("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsha("Region") 
					   else
					       tempRegion = ""
					   end if	
					   tempCompany =rsha("Company")
					   tempfcba = "--ALL--"	
					   tempVH ="--SUMMARY--"
					 %>
					
					<% if SprevCompany <> rsha("Company") then %>
						<td align="center">
							<a href="p_wra.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("company")%>
							</a>
						</td>
						<% SprevCompany = rsha("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
					 <% 
					   
					   					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsha("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsha("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsha("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtcompany <>  "--SUMMARY--" And v_txtcompany <>  "" then
					      tempCompany = rsha("Company")
					   else
					      tempCompany = ""
                       end if				   
					   tempfcba = rsha("fcba") 
					   tempVH ="--ALL--"					   
					 %>
					<% if sprevFcba <> rsha("fcba") then %>
						<td align="center">
						    <a href="p_wra.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("fcba")%>
						</td>
						<% sprevFcba = rsha("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
					
					<td align="right"><%=formatnumber(cdbl(rsha("totBI")),2)%></td>
					<% if cdbl(jamBI) > 0 then
					      rpperjamBI = cdbl(rsha("totBI"))/ cdbl(jamBI)
					   else 
						  rpperjamBI = 0
					   end if %>
					<td align="right"><%=formatnumber(cdbl(rpperjamBI),1)%></td>
					   
					<td align="right"><%=formatnumber(cdbl(rsha("totbudBI")),2)%></td>
					
					<% if cdbl(jambudBI) > 0 then
					      rpperjambudBI = cdbl(rsha("totbudBI"))/ cdbl(jambudBI)
					  else 
						  rpperjambudBI = 0
					  end if %>
					<td align="right"><%=formatnumber(cdbl(rpperjambudBI),3)%></td>
					
					<% 'persen BI %>
					<% IF cdbl(rpperjamBI) = 0 and cdbl(rpperjambudBI) > 0 THEN %>
					  <td align="right">100</td>
					<% ELSEIF cdbl(rpperjamBI) > 0 and cdbl(rpperjambudBI) = 0 THEN %>
					  <td align="right">-100</td>
					<% ELSEIF cdbl(rpperjamBI) = 0 and cdbl(rpperjambudBI) = 0 THEN %>
					  <td align="right">0</td>
					<% ELSEIF cdbl(rpperjamBI) > 0 and cdbl(rpperjambudBI) > 0 THEN %>
					  <td align="right"><%=formatnumber((cdbl(rpperjambudBI) - cdbl(rpperjamBI)) / cdbl(rpperjambudBI)*100,,,-1)%></td>
                    <% END IF %> 	
					
					<td align="right"><%=formatnumber(cdbl(rsha("totSBI")),2)%></td>
					<% if cdbl(jamSBI) > 0 then
					      rpperjamSBI = cdbl(rsha("totSBI"))/ cdbl(jamSBI)
					   else 
						  rpperjamBI = "0" 
					end if %>
					<td align="right"><%=formatnumber(cdbl(rpperjamSBI),2)%></td>
					<td align="right"><%=formatnumber(cdbl(rsha("totbudSBI")),2)%></td>
					<% if cdbl(jambudSBI) > 0 then
					      rpperjambudSBI = cdbl(rsha("totbudSBI"))/ cdbl(jambudSBI)
					  else 
						  rpperjambudSBI = "0" 
					  end if %>
					<td align="right"><%=formatnumber(cdbl(rpperjambudSBI),2)%></td>
					
					<% 'persen SBI %>
					<% IF cdbl(rpperjamSBI) = 0 and cdbl(rpperjambudSBI) > 0 THEN %>
					  <td align="right">100</td>
					<% ELSEIF cdbl(rpperjamSBI) > 0 and cdbl(rpperjambudSBI) = 0 THEN %>
					  <td align="right">-100</td>
					<% ELSEIF cdbl(rpperjamSBI) = 0 and cdbl(rpperjambudSBI) = 0 THEN %>
					  <td align="right">0</td>
					<% ELSEIF cdbl(rpperjamSBI) > 0 and cdbl(rpperjambudSBI) > 0 THEN %>
					  <td align="right"><%=formatnumber((cdbl(rpperjambudSBI) - cdbl(rpperjamSBI)) / cdbl(rpperjambudSBI) *100,,,-1)%></td>
                    <% END IF %> 
					
					<td align="right"><%=formatnumber(cdbl(rsha("totbudTHN")),2)%> </td>
					<% if cdbl(jambudTHN) > 0 then
					      rpperjambudTHN = cdbl(rsha("totbudTHN"))/ cdbl(jambudTHN)
						  
					  else 
						  rpperjambudTHN = 0 
						 
					  end if %>
					<td align="right"><%=formatnumber(cdbl(rpperjambudTHN),2)%></td>
				<tr>	
          <%
		     subtotBI = cdbl(subtotBI) + cdbl(rsha("totBI"))
			 subtotSBI = cdbl(subtotSBI) + cdbl(rsha("totSBI"))
			 subrpperjamBI = cdbl(subrpperjamBI) + cdbl(rpperjamBI)
			 subrpperjamSBI = cdbl(subrpperjamSBI) + cdbl(rpperjamSBI)
			 subtotbudBI = cdbl(subtotbudBI) + cdbl(rsha("totbudBI"))
			 subtotbudSBI = cdbl(subtotbudSBI) + cdbl(rsha("totbudSBI"))
			 subtotbudTHN = cdbl(subtotbudTHN) + cdbl(rsha("totbudTHN"))
			 subrpperjambudBI = cdbl(subrpperjambudBI) + cdbl(rpperjambudBI)
			 subrpperjambudSBI = cdbl(subrpperjambudSBI) + cdbl(rpperjambudSBI)
			 subrpperjambudTHN = cdbl(subrpperjambudTHN) + cdbl(rpperjambudTHN)
			 rsha.MoveNext
		  wend
	end if
	%>	
   <tr>
       <td style="background-color:#F4F4A0" align="center" colspan="<%= 2 + nSpanTotal %>" ><b>TOTAL</td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrpperjamBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotbudBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrpperjambudBI),2)%></td>
	   
	    <% 'persen BI %>
	   <% IF cdbl(subrpperjamBI) = 0 and cdbl(subrpperjambudBI) > 0 THEN %>
		<td style="background-color:#F4F4A0" align="right"><b>100</td>
	   <% ELSEIF cdbl(subrpperjamBI) > 0 and cdbl(subrpperjambudBI) = 0THEN %>
		<td style="background-color:#F4F4A0" align="right"><b>-100</td>
	   <% ELSEIF cdbl(subrpperjamBI) = 0 and cdbl(subrpperjambudBI) = 0 THEN %>
		 <td style="background-color:#F4F4A0" align="right"><b>0</td>
	   <% ELSEIF cdbl(subrpperjamBI) > 0 and cdbl(subrpperjambudBI) > 0 THEN %>
		 <td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subrpperjambudBI) - cdbl(subrpperjamBI)) / cdbl(subrpperjambudBI) *100,,,-1)%></td>
        <% END IF %>
	   
       <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotSBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrpperjamSBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotbudSBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrpperjambudSBI),2)%></td>
	   
	   <% 'persen SBI %>
	   <% IF cdbl(subrpperjamSBI) = 0 and cdbl(subrpperjambudSBI) > 0 THEN %>
		<td style="background-color:#F4F4A0" align="right"><b>100</td>
	   <% ELSEIF cdbl(subrpperjamSBI) > 0 and cdbl(subrpperjambudSBI) = 0THEN %>
		<td style="background-color:#F4F4A0" align="right"><b>-100</td>
	   <% ELSEIF cdbl(subrpperjamSBI) = 0 and cdbl(subrpperjambudSBI) = 0 THEN %>
		 <td style="background-color:#F4F4A0" align="right"><b>0</td>
	   <% ELSEIF cdbl(subrpperjamSBI) > 0 and cdbl(subrpperjambudSBI) > 0 THEN %>
		 <td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subrpperjambudSBI) - cdbl(subrpperjamSBI)) / cdbl(subrpperjambudSBI) *100,,,-1)%></td>
        <% END IF %> 
					
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotbudTHN),2)%></td>
	   <% IF subrpperjambudTHN = 0 THEN %>
	      <td style="background-color:#F4F4A0"  align="right"><b>0</td>
	   <% ELSE %>
	      <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrpperjambudTHN),2)%></td>
	   <% END IF %> 	  
   </tr>    
</table>

<!--#include file="bottom.Asp"-->