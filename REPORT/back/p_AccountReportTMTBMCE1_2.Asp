<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Account Report Biaya TM/TBM ( Cost Element )</title>
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
body{font-family:'Arial Narrow';font-size:12px;}
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	<div id="header-wrapper">
		<div class="printhide">
		<!--#include file="menutop.asp"-->
		<!--#include file="FormSelectTMTBM.html"-->
		</div>
	
		<div id="header">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered">
			<% Dim ia,thntnm
			   ia = request.Form("vMonth")
			   if ia ="" then
				  ia="0" 
			   end if
			   if   request.Form("vPlantingDate") = "0"  then
					thntnm = "ALL"
			   else		
					thntnm = request.Form("vPlantingDate")
			   end if
					
			%>
			<%  IF request.Form("vStatus") = "TM" THEN %>		
			   <h3>LAPORAN BIAYA TANAMAN MENGHASILKAN (TM) ( Cost Element )</h3>
			<%  ELSE %>		
			   <h3>LAPORAN BIAYA TANAMAN BELUM MENGHASILKAN (TBM) ( Cost Element )</h3>   
			<%  END IF %>	   
			<% IF request.Form("vMonth") <> "" then %>
			<h5>PERIODE     : <%=ucase(MonthName(cint(ia)))%> &nbsp; <%=request.Form("vYear")%></h5>
			<% END IF %>
			<h5>CROP        : <%=request.Form("vCROPTYPE")%> &nbsp; KEPEMILIKAN : <%=request.Form("vOwnership")%></h5>
			<h5>DIVISI	    : <%=request.Form("vDivision")%> &nbsp; TAHUN TANAM : <%=thntnm%></h5>
			<h5>SOURCE      : <%=ucase(request.Form("vJenisReport"))%></h5>
		</div>
	</div>
	
<%


' Budget HA dan luas HA

Dim strha, rsha,luasHA,budHA,ProdTon ,prodkg, strprod,cur_month,strbud_BI,strbud_SBI,x1
	    strbud_BI = "a.Budget_01" 
	    strbud_SBI = "a.Budget_01"
		cur_month = request.Form("vMonth") 
		
		IF Len(cur_month) > 0 then

			strbud_BI = "" 
	        strbud_SBI = ""
			
			x1=1
			
			Do While x1<= cint(cur_month) 'Output the values from 0 to 10
			  'response.write(i & "<br \>")
			  'response.write  " Budget_" & Lpad(x1,"0",2) & " +"
			  strbud_SBI = strbud_SBI & " a.Budget_" & Lpad(x1,"0",2) & " +"
			  x1 = x1 + 1 'increment the value of i for next time loop executes
			Loop

			strbud_BI = "a.Budget_" & Lpad(cur_month,"0",2)
			strbud_SBI = Mid(strbud_SBI,1,Len(strbud_SBI)-1)
            'response.write  strbud_SBI	& "<br \>"		
	        'response.write  strbud_BI	& "<br \>"	
		END IF	
		
luasHA=0
budHA=0

		
strha  = " SELECT DECODE(SUM (ha), NULL,1,0,1,SUM (ha)) HA,DECODE(SUM (bud_ha), NULL,1,0,1,SUM (bud_ha)) bud_ha FROM " & _
		 "	( SELECT fcba, croptype, ownership, division, DX_BLOCK_STATUS as status, TO_CHAR(plantingdate,'YYYY') tahuntanam, SUM (hectarageplanted) ha,0 bud_ha FROM field  " & _
		 "	  GROUP BY fcba, croptype, ownership, division, DX_BLOCK_STATUS, plantingdate " & _
		 "	  UNION " & _
		 "	  select fcba,croptype,ownership,division,status,tahuntanam,0 ha, sum(hectarageplanted)  bud_ha  from  " & _
		 "		( select  distinct t.fcba,t.croptype,t.ownership,t.division,t.status,to_char(t.plantingdate,'YYYY') tahuntanam, hectarageplanted " & _
		 "		  from BUDGET_TRANSACTION_ESTATE t ) " & _
		 "	  GROUP BY fcba, croptype, ownership, division, status, tahuntanam " & _
		 "	) WHERE STATUS = '" & request("vStatus")  & "' AND CROPTYPE = '"& request("vCROPTYPE") &"' and fcba='" & session("fcba") & "' "         

	   
		  if  request.Form("vOwnership") <>"ALL" then
			  strsql2 = "  and ownership ='" & request("vOwnership") & "'" 
		  else
			  strsql2= " "
		  end if
		  
		  if   request.Form("vPlantingDate") <> "0"  then
			  strsql3 = " and tahuntanam ='" & request.Form("vPlantingDate") & "'"
		  else
			  strsql3 = " "
		  end if
		  if  request.Form("vDivision") <>"ALL" then
			  strsql4 = " and Division='" & request.Form("vDivision") & "'" 
		  else
			  strsql4 = " "
		  end if
		  
		  strha = strha + strsql2 +  strsql3 + strsql4	 
		  
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strha, connstr, adOpenStatic
'response.write strha 

if Not rsha.eof then
    luasHA = rsha("ha")
	budHA = rsha("bud_ha")
end if
rsha.close()

Dim prodtonBI,prodkgBI,budProdtonBI,budProdkgBI,prodtonSBI,prodkgSBI,budProdtonSBI,budProdkgSBI,budProdtonthn,budProdkgthn,TbudProdtonthn,TbudProdkgthn, TprodtonBI,TprodtonSBI,TbudProdtonBI,TbudProdtonSBI 

' prodtonBI=1
' prodkgBI=1
' budProdtonBI=1
' budProdkgBI=1
' prodtonSBI=1
' prodkgSBI=1
' budProdtonSBI=1
' budProdkgSBI=1		

if request("vStatus") = "TM" then
				
				  if  request.Form("vOwnership") <>"ALL" then
					  strsql2 = "  and f.ownership  ='" & request("vOwnership") & "'" 
				  else
					  strsql2= " "
				  end if
				  
				  if   request.Form("vPlantingDate") <> "0"  then
					  strsql3 = " and to_char(f.PlantingDate,'YYYY') ='" & request.Form("vPlantingDate") & "'"
				  else
					  strsql3 = " "
				  end if
				  if  request.Form("vDivision") <>"ALL" then
					  strsql4 = " and f.Division='" & request.Form("vDivision") & "'" 
				  else
					  strsql4 = " "
				  end if
				 
		
        strprod = " select fcba,sum(prod_tonBI) prod_tonBI,sum(budProdBI) budProdBI,sum(prod_tonSBI) prod_tonSBI,sum(budProdSBI) budProdSBI,sum(budthn) budthn  from " & _
				  "	( select h.fcba, round(nvl(sum(h.mill_weight),0),0) prod_tonBI,0 budProdBI,0  prod_tonSBI,0 budProdSBI,0 budthn from harvesting h , field f " & _
				  "	  where h.fieldcode = f.fccode and h.fcba = f.fcba and f.STATUS = '" & request("vStatus")  & "' AND f.CROPTYPE =  '"& request("vCROPTYPE") &"' and f.fcba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql4 & "  " & _
				  "	  and to_char(h.receptiondate,'YYYY')='" & request.Form("vYear") & "' and to_char(h.receptiondate,'MM')=lpad('1',2,0) group by h.fcba " & _
				  "	  union " & _
				  "	  select h.fcba,0 prod_tonBI,0 budProdBI,round(nvl(sum(h.mill_weight),0),0)  prod_tonSBI,0 budProdSBI,0 budthn from harvesting h , field f " & _
				  "	  where h.fieldcode = f.fccode and h.fcba = f.fcba and f.STATUS = '" & request("vStatus")  & "' AND f.CROPTYPE = '"& request("vCROPTYPE") &"' and f.fcba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql4 & "  " & _
				  "	  and to_char(h.receptiondate,'YYYY')='" & request.Form("vYear") & "' and to_char(h.receptiondate,'MM')<=lpad('1',2,0) group by h.fcba " & _
				  "	  union " & _
				  "	  select f.fcba,0 prod_tonBI,sum(f.qty_01) budProdBI,0  prod_tonSBI,0 budProdSBI,0 budthn from BUDGET_TRANSACTION_PRODQTY f " & _
				  "	  where f.budget_year ='" & request.Form("vYear") & "' and f.fcba ='" & session("fcba") & "' "  & strsql3 &  strsql4 & "  " & _
				  "	  group by f.fcba " & _
				  "	  union " & _
				  "	  select f.fcba,0 prod_tonBI,0 budProdBI,0  prod_tonSBI,sum(f.qty_01) budProdSBI,0 budthn from BUDGET_TRANSACTION_PRODQTY f " & _
				  "	  where f.budget_year ='" & request.Form("vYear") & "' and f.fcba ='" & session("fcba") & "' "  & strsql3 &  strsql4 & "  " & _
				  "	  group by f.fcba " & _
				  "    union " & _
				  "	  select f.fcba,0 prod_tonBI,0 budProdBI,0  prod_tonSBI,0 budProdSBI,sum(f.qty_01 + f.qty_02 + f.qty_03 + f.qty_04 + f.qty_05 + f.qty_06 + f.qty_07 + f.qty_08 + f.qty_09 + f.qty_10 + f.qty_11 + f.qty_12 ) budthn from BUDGET_TRANSACTION_PRODQTY f " & _
				  "	  where f.budget_year ='" & request.Form("vYear") & "' and f.fcba ='" & session("fcba") & "' "  & strsql3 &  strsql4 & "  " & _
				  "	  group by f.fcba " & _
				  "	) group by fcba	"
				 
				
		Set rsha = Server.CreateObject("ADODB.Recordset")
		rsha.Open strprod, connstr, adOpenStatic
        'response.write strprod
		if Not rsha.eof then
			prodtonBI=cdbl(rsha("prod_tonBI"))/1000
			prodkgBI=cdbl(rsha("prod_tonBI"))
			budProdtonBI=cdbl(rsha("budProdBI"))/1000
			budProdkgBI=cdbl(rsha("budProdBI"))
			prodtonSBI=cdbl(rsha("prod_tonSBI"))/1000
			prodkgSBI=cdbl(rsha("prod_tonSBI"))
			budProdtonSBI=cdbl(rsha("budProdSBI"))/1000
			budProdkgSBI=cdbl(rsha("budProdSBI"))
			budProdtonthn=cdbl(rsha("budthn"))/1000
			budProdkgthn=cdbl(rsha("budthn"))
		end if
		rsha.close()
end if

TprodtonBI=prodtonBI
TprodtonSBI=prodtonSBI 
TbudProdtonBI=budProdtonBI
TbudProdtonSBI=budProdtonSBI
TbudProdtonthn=budProdtonthn

if cdbl(prodtonBI)=0 then
   prodtonBI=1
end if
if cdbl(prodkgBI)=0 then
   prodkgBI=1
end if
if cdbl(budProdtonBI)=0 then
   budProdtonBI=1
end if
if cdbl(budProdkgBI)=0 then 
   budProdkgBI=1
end if
if cdbl(prodtonSBI)=0 then
	prodtonSBI=1
end if
if cdbl(prodkgSBI)=0 then
	prodkgSBI=1
end if
if budProdtonSBI=0 then
	budProdtonSBI=1
end if
if cdbl(budProdkgSBI)=0 then
	budProdkgSBI=1	
end if
if cdbl(budProdkgthn)=0 then
	budProdkgthn=1	
end if
%>	
    
	
<table border="0" width="100%" id="product-table">
	<tr>
		<th colspan="3">&nbsp;</th>
		<th style="color:#FFFFFF" colspan="5">BULAN INI</th>
		<th style="color:#FFFFFF" colspan="5">S/D BULAN INI</th>
		<th style="color:#FFFFFF" colspan="2" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th colspan="3">&nbsp;</th>
		
		<th style="color:#FFFFFF" colspan="2">REALISASI</th>
		
		<th style="color:#FFFFFF" colspan="2">ANGGARAN</th>
		<th style="color:#FFFFFF" >%</th>
		
		<th style="color:#FFFFFF" colspan="2">REALISASI</th>
		<th style="color:#FFFFFF" colspan="2">ANGGARAN</th>
		<th style="color:#FFFFFF" >%</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF" colspan="3" align="right">LUAS Ha</th>

		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(luasHA,0) %></th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(budHA,0) %></th>
		<th style="color:#FFFFFF" align="right" ></th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(luasHA,0)%></th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(budHA,0) %></th>
		<th style="color:#FFFFFF" align="right" ></th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(budHA,0) %></th>

	</tr>
	<tr>
		<th style="color:#FFFFFF" colspan="3" align="right">PRODUKSI - Ton</th>

		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(TprodtonBI,0) %> </th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(TbudProdtonBI,0) %></th>
		
		<% ' Prosentase %>
		<% if cdbl(TbudProdtonBI) = 0 and cdbl(TprodtonBI) > 0 then %>
		    <th style="color:#FFFFFF" align="right" >100</th>
		<% elseif cdbl(TbudProdtonBI) > 0 and cdbl(TprodtonBI) = 0 then%>
		    <th style="color:#FFFFFF" align="right" >-100</th>
		<% elseif cdbl(TbudProdtonBI) = 0 and cdbl(TprodtonBI) = 0 then%>
		    <th style="color:#FFFFFF" align="right" >0</th>
		<% else %>	
		    <th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(TprodtonBI) - cdbl(TbudProdtonBI)) / cdbl(TbudProdtonBI)*100,2) %></th>
		<% End If %>
		
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(TprodtonSBI,0)%></th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(TbudProdtonSBI,0) %></th>
		
		<% ' Prosentase SBI %>
		<% if cdbl(TbudProdtonSBI) = 0 and cdbl(TprodtonSBI) > 0 then %>
		    <th style="color:#FFFFFF" align="right" >100</th>
		<% elseif cdbl(TbudProdtonSBI) > 0 and cdbl(TprodtonSBI) = 0 then%>
		    <th style="color:#FFFFFF" align="right" >-100</th>
		<% elseif cdbl(TbudProdtonSBI) = 0 and cdbl(TprodtonSBI) = 0 then%>
		    <th style="color:#FFFFFF" align="right" >0</th>
		<% else %>	
		    <th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(TprodtonSBI) - cdbl(TbudProdtonSBI)) / cdbl(TbudProdtonSBI)*100,2) %></th>
		<% End If %>
		
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(TbudProdtonthn,0) %></th>
		

	</tr>
	<tr>
		<th style="color:#FFFFFF" colspan="3" align="right">Yield - Ton/Ha�</th>

		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(cdbl(TprodtonBI)/cdbl(luasHA),2) %></th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(cdbl(TbudProdtonBI)/cdbl(budHA),2) %></th>
		<th style="color:#FFFFFF" align="right" ></th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(cdbl(TprodtonSBI)/cdbl(luasHA),2) %></th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(cdbl(TbudProdtonSBI)/cdbl(budHA),2) %></th>
		<th style="color:#FFFFFF" align="right" ></th>
		<th style="color:#FFFFFF" align="right" colspan="2"><%=formatnumber(cdbl(TbudProdtonthn)/cdbl(budHA),2) %></th>
		

	</tr>
	<tr>
		
		<th style="color:#FFFFFF" >KODE</th>
		<% if   request.Form("vJenisReport") = "Group"  then %>
		    <th  style="color:#FFFFFF" >GROUP PEKERJAAN</th>
	    <% else %>
		    <th  style="color:#FFFFFF" >JENIS PEKERJAAN</th>
		<% End If %>
		<th style="color:#FFFFFF" >Cost Element</th>
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Ha</th>
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Ha</th>
		<th style="color:#FFFFFF" >&nbsp;</th>
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Ha</th>
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Ha</th>
		<th style="color:#FFFFFF" >&nbsp;</th>
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Ha</th>
	</tr>
	<%
	 Dim strdtl , strgrp ,cdgrp ,noline , subAmtgrpBI,subAmtgrpSBI , GRAmtgrpBI,GRAmtgrpSBI, subbudgrpBI,subbudgrpSBI,subbudgrpthn , GRbudgrpBI,GRbudgrpSBI ,GRbudgrpthn  , grpDesc
		
			if   request.Form("vPlantingDate") <> "0"  then
				  strsql2 = " and substr(PlantingDate,-4) ='" & request.Form("vPlantingDate") & "'"
			 else
				  strsql2 = " "
			 end if
			 if  request.Form("vDivision") <>"ALL" then
				  strsql3 = " and a.division='" & request.Form("vDivision") & "'" 
			 else
				  strsql3 = " "
			 end if
			 if  request.Form("vYear") <>"ALL" then
				  strsql4 = " and a.period_year='" & request.Form("vYear") & "'" 
			 else
				  strsql4 = " "
			 end if
			 if  request.Form("vMonth") <>"ALL" then
				  strsql5 = " and a.period_month='" & request.Form("vMonth") & "'" 
			 else 
				  strsql5 = " "
			 end if
			if  request.Form("vOwnership") <>"ALL" then
				strsql6 = "  and a.ownership  ='" & request("vOwnership") & "'" 
			else
				strsql6= " "
			end if
	 
	 if   request.Form("vJenisReport") = "Group"  then
					  
		   ' strdtl =  " select trim(code) code,jobheading,trim(job_name) job_name,sourcetype,sourcename,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn from  " & _ 
					 ' "	( select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name ,a.sourcetype,a.sourcename ,sum(a.tm_amt_actual) totalBI,0 BudBI,0 totalSBI,0 BudSBI,0 Budthn from sac_accountreport_tm_tbm a " & _
					 ' "	  where a.period_year='" & request.Form("vYear") & "' and a.period_month=lpad('" & request.Form("vMonth") & "',2,0) and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode,a.sourcetype,a.sourcename " & _
					 ' "	  UNION  " & _
					 ' "	  select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name,a.sourcetype,a.sourcename ,0 totalBI,0 BudBI,sum(a.tm_amt_actual) totalSBI,0 BudSBI,0 Budthn from sac_accountreport_tm_tbm a " & _
					 ' "	  where a.period_year='" & request.Form("vYear") & "' and a.period_month <=lpad('" & request.Form("vMonth") & "',2,0) and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & " " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode ,a.sourcetype,a.sourcename " & _
					 ' "	  UNION " & _
					 ' "	  select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name,a.sourcetype,a.sourcename ,0 totalBI,sum(" & strbud_BI & ")  BudBI  ,0 totalSBI,0 BudSBI,0 Budthn from view_budget_ar_tm_tbm a " & _
					 ' "	  where a.BUDGET_YEAR ='" & request.Form("vYear") & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode,a.sourcetype,a.sourcename " & _
					 ' "	  UNION " & _
					 ' "	  select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name,a.sourcetype,a.sourcename ,0 totalBI,0  BudBI  ,0 totalSBI,sum(" & strbud_SBI & ") BudSBI,0 Budthn from view_budget_ar_tm_tbm a " & _
					 ' "	  where a.BUDGET_YEAR ='" & request.Form("vYear") & "'  and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "   " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode,a.sourcetype,a.sourcename " & _
					 ' "	  UNION " & _
					 ' "	  select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name,a.sourcetype,a.sourcename ,0 totalBI,0  BudBI  ,0 totalSBI,0 BudSBI,sum(a.BUDGET_01+a.BUDGET_02+a.BUDGET_03+a.BUDGET_04+a.BUDGET_05+a.BUDGET_06+a.BUDGET_07+a.BUDGET_08+a.BUDGET_09+a.BUDGET_10+a.BUDGET_11+a.BUDGET_12) Budthn from view_budget_ar_tm_tbm a " & _  
					 ' "	  where a.BUDGET_YEAR ='" & request.Form("vYear") & "'  and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode,a.sourcetype,a.sourcename " & _
					 ' "	) group by code,jobheading,job_name,sourcetype,sourcename order by  code,jobheading,job_name " 	

             strdtl =   " select trim(code) code,jobheading,jobgroup_api.get_jobgroup_name(code,ba) job_name,sourcetype,get_locationtype_name(SOURCETYPE) Sourcename ,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn  from " & _
					    "	( select a.ba,a.accountgroupcode code,a.activity_heading jobheading,a.sourcetype ,sum(a.tm_amt_actual) totalBI,0 BudBI,0 totalSBI,0 BudSBI,0 Budthn from sac_accountreport_tm_tbm a " & _
					    "	  where a.period_year='" & request.Form("vYear") & "' and a.period_month=lpad('" & request.Form("vMonth") & "',2,0)and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
						"	  GROUP BY a.ba,a.activity_heading,a.accountgroupcode ,a.sourcetype " & _
						"	  UNION " & _
						"	  select a.ba,a.accountgroupcode code,a.activity_heading jobheading ,a.sourcetype ,0 totalBI,0 BudBI,sum(a.tm_amt_actual) totalSBI,0 BudSBI,0 Budthn from sac_accountreport_tm_tbm a " & _
						"	  where a.period_year='" & request.Form("vYear") & "' and a.period_month <=lpad('" & request.Form("vMonth") & "',2,0)and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
						"	  GROUP BY a.ba,a.activity_heading,a.accountgroupcode ,a.sourcetype " & _
						"	  UNION " & _
						"	  select a.ba,a.accountgroupcode code,a.activity_heading jobheading,a.sourcetype  ,0 totalBI,sum(" & strbud_BI & ")  BudBI  ,0 totalSBI,0 BudSBI,0 Budthn from view_budget_ar_tm_tbm a " & _
						"	  where a.BUDGET_YEAR ='" & request.Form("vYear") & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
						"	  GROUP BY a.ba, a.activity_heading,a.accountgroupcode,a.sourcetype " & _
						"	  UNION  " & _
						"	  select a.ba,a.accountgroupcode code,a.activity_heading jobheading ,a.sourcetype,0 totalBI,0  BudBI  ,0 totalSBI,sum(" & strbud_SBI & ") BudSBI,0 Budthn from view_budget_ar_tm_tbm a " & _
						"	  where a.BUDGET_YEAR ='" & request.Form("vYear") & "'  and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
						"	  GROUP BY a.ba,a.activity_heading,a.accountgroupcode ,a.sourcetype " & _
						"	  UNION " & _
						"	  select a.ba,a.accountgroupcode code,a.activity_heading jobheading,a.sourcetype  ,0 totalBI,0  BudBI  ,0 totalSBI,0 BudSBI,sum(a.BUDGET_01+a.BUDGET_02+a.BUDGET_03+a.BUDGET_04+a.BUDGET_05+a.BUDGET_06+a.BUDGET_07+a.BUDGET_08+a.BUDGET_09+a.BUDGET_10+a.BUDGET_11+a.BUDGET_12) Budthn from view_budget_ar_tm_tbm a " & _
						"	  where a.BUDGET_YEAR ='" & request.Form("vYear") & "'  and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
						"	  GROUP BY a.ba,a.activity_heading,a.accountgroupcode ,a.sourcetype " & _
						"	) group by ba,code,jobheading,sourcetype order by code,jobheading "


			 
	 else  	
					
		   ' strdtl = " select trim(code) code,jobheading,trim(job_name) job_name,sourcetype,sourcename,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn from " & _
					' "	( select a.accountcode code,a.activity_heading jobheading ,a.Description job_name ,a.sourcetype,a.sourcename ,sum(a.tm_amt_actual) totalBI,0 BudBI,0 totalSBI,0 BudSBI,0  Budthn from sac_accountreport_tm_tbm a " & _
					' "	  where a.period_year='" & request.Form("vYear") & "' and a.period_month=lpad('" & request.Form("vMonth") & "',2,0) and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='KSB' " & _
					' "	  GROUP BY a.activity_heading,a.accountcode,a.Description  ,a.sourcetype,a.sourcename " & _
					' "	  UNION  " & _
					' "	  select a.accountcode code,a.activity_heading jobheading ,a.Description job_name ,a.sourcetype,a.sourcename ,0 totalBI,0 BudBI,sum(a.tm_amt_actual) totalSBI,0 BudSBI,0  Budthn from sac_accountreport_tm_tbm a " & _
					' "	  where a.period_year='" & request.Form("vYear") & "' and a.period_month <=lpad('" & request.Form("vMonth") & "',2,0) and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & " " & _
					' "	  GROUP BY a.activity_heading,a.accountcode,a.Description  ,a.sourcetype,a.sourcename " & _
					' "	  UNION " & _
					' "	  select a.accountcode code,a.activity_heading jobheading ,a.Description job_name ,a.sourcetype,a.sourcename ,0 totalBI,sum(" & strbud_BI & ")  BudBI  ,0 totalSBI,0 BudSBI,0  Budthn from view_budget_ar_tm_tbm a " & _
					' "	  where a.BUDGET_YEAR='" & request.Form("vYear") & "'  and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & " " & _
					' "	  GROUP BY a.activity_heading,a.accountcode,a.Description  ,a.sourcetype,a.sourcename " & _
					' "	  UNION " & _
					' "	  select a.accountcode code,a.activity_heading jobheading ,a.Description job_name ,a.sourcetype,a.sourcename ,0 totalBI,0  BudBI  ,0 totalSBI,sum(" & strbud_SBI & ") BudSBI,0  Budthn from view_budget_ar_tm_tbm a " & _
					' "	  where a.BUDGET_YEAR='" & request.Form("vYear") & "'  and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & " " & _
					' "	  GROUP BY a.activity_heading,a.accountcode,a.Description  ,a.sourcetype,a.sourcename " & _
					' "	  UNION " & _
					' "	  select a.accountcode code,a.activity_heading jobheading ,a.Description job_name ,a.sourcetype,a.sourcename ,0 totalBI,0  BudBI  ,0 totalSBI,0 BudSBI,sum(a.BUDGET_01+a.BUDGET_02+a.BUDGET_03+a.BUDGET_04+a.BUDGET_05+a.BUDGET_06+a.BUDGET_07+a.BUDGET_08+a.BUDGET_09+a.BUDGET_10+a.BUDGET_11+a.BUDGET_12) Budthn from view_budget_ar_tm_tbm a " & _
					' "	  where a.BUDGET_YEAR='" & request.Form("vYear") & "'  and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & " " & _
					' "	  GROUP BY a.activity_heading,a.accountcode,a.Description  ,a.sourcetype,a.sourcename " & _
					' "	 ) group by code,jobheading,job_name,sourcetype,sourcename order by code,jobheading,job_name "
         
     		 strdtl = " select trim(code) code,jobheading,job_api.get_jobcode_name(code,ba) job_name,sourcetype,get_locationtype_name(SOURCETYPE) Sourcename ,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn   from " & _
					  "	( select a.ba,a.accountcode code,a.activity_heading jobheading,a.sourcetype ,sum(a.tm_amt_actual) totalBI,0 BudBI,0 totalSBI,0 BudSBI,0 Budthn  from sac_accountreport_tm_tbm a " & _
					  "	  where a.period_year='" & request.Form("vYear") & "' and a.period_month=lpad('" & request.Form("vMonth") & "',2,0)and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					  "	  GROUP BY a.ba,a.activity_heading,a.accountcode,a.sourcetype " & _
					  "	  UNION " & _
					  "	  select a.ba,a.accountcode code,a.activity_heading jobheading,a.sourcetype,0 totalBI,0 BudBI,sum(a.tm_amt_actual) totalSBI,0 BudSBI,0 Budthn  from sac_accountreport_tm_tbm a " & _
					  "	  where a.period_year='" & request.Form("vYear") & "' and a.period_month <=lpad('" & request.Form("vMonth") & "',2,0)and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					  "	  GROUP BY a.ba,a.activity_heading,a.accountcode,a.sourcetype " & _
					  "	  UNION " & _
					  "	  SELECT  a.ba,a.accountcode code,a.activity_heading jobheading,a.sourcetype,0 totalBI,sum(" & strbud_BI & ")  BudBI  ,0 totalSBI,0 BudSBI,0 Budthn  from view_budget_ar_tm_tbm a " & _
					  "	  where  a.BUDGET_YEAR = '" & request.Form("vYear") & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					  "	  GROUP BY a.ba,a.activity_heading,a.accountcode,a.sourcetype " & _
					  "	  UNION  " & _
					  "	  SELECT  a.ba,a.accountcode code,a.activity_heading jobheading,a.sourcetype,0 totalBI,0  BudBI  ,0 totalSBI,sum(" & strbud_SBI & ") BudSBI,0 Budthn from view_budget_ar_tm_tbm a " & _
					  "	  where  a.BUDGET_YEAR = '" & request.Form("vYear") & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					  "	  GROUP BY a.ba,a.activity_heading,a.accountcode,a.sourcetype " & _
					  "	  UNION " & _
					  "	  SELECT  a.ba,a.accountcode code,a.activity_heading jobheading ,a.sourcetype,0 totalBI,0  BudBI  ,0 totalSBI,0 BudSBI,sum(a.BUDGET_01+a.BUDGET_02+a.BUDGET_03+a.BUDGET_04+a.BUDGET_05+a.BUDGET_06+a.BUDGET_07+a.BUDGET_08+a.BUDGET_09+a.BUDGET_10+a.BUDGET_11+a.BUDGET_12) Budthn from view_budget_ar_tm_tbm a " & _
					  "	  where  a.BUDGET_YEAR = '" & request.Form("vYear") & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					  "	  GROUP BY a.ba,a.activity_heading,a.accountcode,a.sourcetype " & _
					  "	) group by ba,code,jobheading,sourcetype order by code,jobheading "
		   
	 end if				 
			 
			
			 
            
    
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    'response.write strdtl
	   
	%>
	
	<%
	  if Not rsha.eof then
		  while Not rsha.eof
	%> 
	   <tr>  
		
		
		    <% IF  cdgrp = rsha("jobheading")  then
			       noline = cint(noline )+ 1
				   subAmtgrpBI = cdbl(subAmtgrpBI) + cdbl(rsha("totalBI"))
				   subAmtgrpSBI = cdbl(subAmtgrpSBI) + cdbl(rsha("totalSBI"))
				   subbudgrpBI = cdbl(subbudgrpBI) + cdbl(rsha("budBI"))
				   subbudgrpSBI = cdbl(subbudgrpSBI) + cdbl(rsha("budSBI"))
				   subbudgrpthn  = cdbl(subbudgrpthn)  + cdbl(rsha("budthn"))
			%>	   
				   
				   <td align="center"><%=rsha("code")%></td>
				   <td><%=rsha("job_name")%></td>
				   <td><%=rsha("sourcetype")%> - <%=rsha("sourcename")%> </td>
				   <td align="right"><%=formatnumber(rsha("totalBI"),0)%></td>	
				   
				   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
							<td align="right"><%= formatnumber( cdbl(rsha("totalBI")) / cdbl(prodkgBI),0) %></td>
					<% else  %> 
							<td align="right"><%= formatnumber( cdbl(rsha("totalBI")) / cint(luasHA),0) %></td>
					<% end if %>	
				   <td align="right"><%=formatnumber(rsha("budBI"),0)%></td>
				   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
							<td align="right"><%= formatnumber( cdbl(rsha("budBI")) / cdbl(budProdkgBI),0) %></td>
					<% else  %> 
							<td align="right"><%= formatnumber( cdbl(rsha("budBI")) / cint(budHA),0) %></td>
					<% end if %>
					
					<% ' prosentase BI %>
					<% IF cdbl(rsha("budBI")) = 0  and cdbl(rsha("totalBI")) > 0 then %>
				       <td align="right">-100</td>
				    <% Elseif cdbl(rsha("budBI")) > 0  and cdbl(rsha("totalBI")) = 0 then %>
					   <td align="right">100</td>
					<% Elseif cdbl(rsha("budBI")) = 0  and cdbl(rsha("totalBI")) = 0 then %>
					   <td align="right">0</td>  
					<% else %>   
					   <td align="right"><%= formatnumber( (cdbl(rsha("budBI")) - cdbl(rsha("totalBI")))/cdbl(rsha("budBI"))*100,2) %></td>
				    <% End If %>
					
				
				   <td align="right"><%=formatnumber(rsha("totalSBI"),0)%></td>
				   
                    <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
						      <td align="right"><%= formatnumber( cdbl(rsha("totalSBI")) / cdbl(prodkgSBI),0) %></td>
					<% else  %> 
					   		  <td align="right"><%= formatnumber( cdbl(rsha("totalSBI")) / cint(luasHA),0) %></td>
					<% end if %>				   
					
					<td align="right"><%=formatnumber(rsha("budSBI"),0)%></td>
					
					<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
						      <td align="right"><%= formatnumber( cdbl(rsha("budSBI")) / cdbl(budProdkgSBI),0) %></td>
					<% else  %> 
					   		  <td align="right"><%= formatnumber( cdbl(rsha("budSBI")) / cint(budHA),0) %></td>
					<% end if %>
					
					<% ' prosentase SBI %>
					<% IF cdbl(rsha("budSBI")) = 0  and cdbl(rsha("totalSBI")) > 0 then %>
				       <td align="right">-100</td>
				    <% Elseif cdbl(rsha("budSBI")) > 0  and cdbl(rsha("totalSBI")) = 0 then %>
					   <td align="right">100</td>
					<% Elseif cdbl(rsha("budSBI")) = 0  and cdbl(rsha("totalSBI")) = 0 then %>
					   <td align="right">0</td>  
					<% else %>   
					   <td align="right"><%= formatnumber( (cdbl(rsha("budSBI")) - cdbl(rsha("totalSBI")))/cdbl(rsha("budSBI"))*100,2) %></td>
				    <% End If %>
					
					<td align="right"><%=formatnumber(rsha("Budthn"),0)%></td>
					<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
						      <td align="right"><%= formatnumber( cdbl(rsha("Budthn")) / cdbl(budProdkgthn),0) %></td>
					<% else  %> 
					   		  <td align="right"><%= formatnumber( cdbl(rsha("Budthn")) / cint(budHA),0) %></td>
					<% end if %>
				<%   
			   ELSE %>
			      
					   <%  if  cdgrp <> "" then  %>
							 <tr>
								<% GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(subAmtgrpBI) 
								   GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(subAmtgrpSBI)
								   GRbudgrpBI = cdbl(GRbudgrpBI) + cdbl(subbudgrpBI) 
								   GRbudgrpSBI = cdbl(GRbudgrpSBI) + cdbl(subbudgrpSBI)
								   GRbudgrpthn = cdbl(GRbudgrpthn) + cdbl(subbudgrpthn)
								   
								%>
								
								<%  
								   If cdgrp = "PRM" Then
								      grpDesc = "Permulaan "
								   elseif 	 cdgrp = "PML" Then 
								      grpDesc = "Pemeliharaan "
								   elseif 	 cdgrp = "GC" Then 
								      grpDesc = "General Charge "	
  								   elseif	cdgrp = "PNN" Then 								   
								      grpDesc = "Panen "	
                                   end if									  
								%>
								
								<td style="background-color:#F4F4A0" colspan="3"><b> Total&nbsp;<%=grpDesc%> </td>	
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpBI,0)%></td>
								<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( cdbl(subAmtgrpBI) / cdbl(prodkgBI ),0) %></td>
								<% else  %> 
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(subAmtgrpBI) / cint(luasHA),0) %></td>
								<% end if %>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudgrpBI,0)%></td>
								<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( cdbl(subbudgrpBI) / cdbl(budProdkgBI),0) %></td>
								<% else  %> 
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(subbudgrpBI) / cint(budHA),0) %></td>
								<% end if %>
								
								<% ' prosentase BI %>
								<% IF cdbl(subbudgrpBI) = 0  and cdbl(subAmtgrpBI) > 0 then %>
								   <td style="background-color:#F4F4A0" align="right"><b>-100</td>
								<% Elseif cdbl(subbudgrpBI) > 0  and cdbl(subAmtgrpBI) = 0 then %>
								   <td style="background-color:#F4F4A0" align="right"><b>100</td>
								<% Elseif cdbl(subbudgrpBI) = 0  and cdbl(subAmtgrpBI) = 0 then %>
								   <td style="background-color:#F4F4A0" align="right"><b>0</td>  
								<% else %>   
								   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( (cdbl(subbudgrpBI) - cdbl(subAmtgrpBI)) /  cdbl(subbudgrpBI)*100 ,2) %></td>
								<% End If %>
								
								
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpSBI,0)%></td>
								 <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(subAmtgrpSBI) / cdbl(prodkgSBI ),0) %></td>
								<% else  %> 
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(subAmtgrpSBI) / cint(luasHA),0) %></td>
								<% end if %>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudgrpSBI,0)%></td>
								 <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( cdbl(subbudgrpSBI) / cdbl(budProdkgSBI),0) %></td>
								<% else  %> 
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(subbudgrpBI) / cint(budHA),0) %></td>
								<% end if %>
								
								<% ' prosentase SBI %>
								<% IF cdbl(subbudgrpSBI) = 0  and cdbl(subAmtgrpSBI) > 0 then %>
								   <td style="background-color:#F4F4A0" align="right"><b>-100</td>
								<% Elseif cdbl(subbudgrpSBI) > 0  and cdbl(subAmtgrpSBI) = 0 then %>
								   <td style="background-color:#F4F4A0" align="right"><b>100</td>
								<% Elseif cdbl(subbudgrpSBI) = 0  and cdbl(subAmtgrpSBI) = 0 then %>
								   <td style="background-color:#F4F4A0" align="right"><b>0</td>  
								<% else %>   
								   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( (cdbl(subbudgrpSBI) - cdbl(subAmtgrpSBI)) /  cdbl(subbudgrpSBI)*100 ,2) %></td>
								<% End If %>
								
								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudgrpthn,0)%></td>
								 <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( cdbl(subbudgrpthn) / cdbl(budProdkgthn),0) %></td>
								<% else  %> 
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(subbudgrpthn) / cint(budHA),0) %></td>
								<% end if %>
								
							 </tr>
														
						<%  end if %>
				
						<%   noline = 1
							 subAmtgrpBI = rsha("totalBI")
							 subAmtgrpSBI = rsha("totalSBI")
							 subbudgrpBI = rsha("budBI")
				             subbudgrpSBI = rsha("budSBI")
							 subbudgrpthn  = cdbl(rsha("budthn"))
						%>
				   
					   <%  if  cdgrp = "PML" and request("vStatus") = "TM" then  %>
							<tr>
								<td colspan="2">&nbsp;</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
							</tr>		
					   <%  end if %>				   
					   <tr>
						   <td align="center"><%=rsha("code")%></td>
				   
						   <td><%=rsha("job_name")%></td>
						   <td><%=rsha("sourcetype")%> - <%=rsha("sourcename")%> </td>
						   <td align="right"><%=formatnumber(rsha("totalBI"),0)%></td>	
						   
						   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									<td align="right"><%= formatnumber( cdbl(rsha("totalBI")) / cdbl(prodkgBI),0) %></td>
							<% else  %> 
									<td align="right"><%= formatnumber( cdbl(rsha("totalBI")) / cint(luasHA),0) %></td>
							<% end if %>	
						   <td align="right"><%=formatnumber(rsha("budBI"),0)%></td>
						   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									<td align="right"><%= formatnumber( cdbl(rsha("budBI")) / cdbl(budProdkgBI),0) %></td>
							<% else  %> 
									<td align="right"><%= formatnumber( cdbl(rsha("budBI")) / cint(budHA),0) %></td>
							<% end if %>
							
							<% ' prosentase BI %>
							<% IF cdbl(rsha("budBI")) = 0  and cdbl(rsha("totalBI")) > 0 then %>
							   <td align="right">-100</td>
							<% Elseif cdbl(rsha("budBI")) > 0  and cdbl(rsha("totalBI")) = 0 then %>
							   <td align="right">100</td>
							<% Elseif cdbl(rsha("budBI")) = 0  and cdbl(rsha("totalBI")) = 0 then %>
							   <td align="right">0</td>  
							<% else %>   
							   <td align="right"><%= formatnumber( (cdbl(rsha("budBI")) - cdbl(rsha("totalBI")))/cdbl(rsha("budBI"))*100,2) %></td>
							<% End If %>
							
							
							
						   <td align="right"><%=formatnumber(rsha("totalSBI"),0)%></td>
							<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td align="right"><%= formatnumber( cdbl(rsha("totalSBI")) / cdbl(prodkgSBI),0) %></td>
							<% else  %> 
									  <td align="right"><%= formatnumber( cdbl(rsha("totalSBI")) / cint(luasHA),0) %></td>
							<% end if %>				   
							
							<td align="right"><%=formatnumber(rsha("budSBI"),0)%></td>
							<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td align="right"><%= formatnumber( cdbl(rsha("budSBI")) / cdbl(budProdkgSBI),0) %></td>
							<% else  %> 
									  <td align="right"><%= formatnumber( cdbl(rsha("budSBI")) / cint(budHA),0) %></td>
							<% end if %>
							
							<% ' prosentase SBI %>
							<% IF cdbl(rsha("budSBI")) = 0  and cdbl(rsha("totalSBI")) > 0 then %>
							   <td align="right">-100</td>
							<% Elseif cdbl(rsha("budSBI")) > 0  and cdbl(rsha("totalSBI")) = 0 then %>
							   <td align="right">100</td>
							<% Elseif cdbl(rsha("budSBI")) = 0  and cdbl(rsha("totalSBI")) = 0 then %>
							   <td align="right">0</td>  
							<% else %>   
							   <td align="right"><%= formatnumber( (cdbl(rsha("budSBI")) - cdbl(rsha("totalSBI")))/cdbl(rsha("budSBI"))*100,2) %></td>
							<% End If %>
							
							<td align="right"><%=formatnumber(rsha("Budthn"),0)%></td>
							<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td align="right"><%= formatnumber( cdbl(rsha("Budthn")) / cdbl(budProdkgthn),0) %></td>
							<% else  %> 
									  <td align="right"><%= formatnumber( cdbl(rsha("Budthn")) / cint(budHA),0) %></td>
							<% end if %>
							
					   </tr>
					   
					   
				   <%
			   END IF
			%>   
		   
		</tr>
	<%
	      
		  cdgrp = rsha("jobheading")
		rsha.MoveNext
		  wend
	  end if
	%>
	
	<tr>
	     <% GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(subAmtgrpBI) 
		    GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(subAmtgrpSBI) 
			GRbudgrpBI = cdbl(GRbudgrpBI) + cdbl(subbudgrpBI) 
			GRbudgrpSBI = cdbl(GRbudgrpSBI) + cdbl(subbudgrpSBI)
			GRbudgrpthn = cdbl(GRbudgrpthn) + cdbl(subbudgrpthn) 
			
			Dim grandtot_rp_ha_BI,grandtot_rp_ha_budBI,grandtot_rp_kg_BI,grandtot_rp_kg_budBI, grandtot_rp_ha_SBI,grandtot_rp_ha_budSBI ,grandtot_rp_kg_SBI,grandtot_rp_kg_budSBI 
			                       
								   if cdbl(luasHA) = 0 then
								      grandtot_rp_ha_BI = 0
									  grandtot_rp_ha_SBI = 0
								   else
									   grandtot_rp_ha_BI = cdbl(GRAmtgrpBI) / cdbl(luasHA)
									   grandtot_rp_ha_SBI = cdbl(GRAmtgrpSBI) / cint(luasHA)
								   end if
								  
								   
								   if cdbl(budHA) = 0 then 
								       grandtot_rp_ha_budBI = 0
									   grandtot_rp_ha_budSBI = 0
								   else
									   grandtot_rp_ha_budBI = cdbl(GRbudgrpBI) / cint(budHA)
									   grandtot_rp_ha_budSBI = cdbl(GRbudgrpSBI) / cint(budHA)
								   end if
								   
								   If cdbl(prodkgBI) = 0 then
								      grandtot_rp_kg_BI = 0
                                   else									  
								      grandtot_rp_kg_BI = cdbl(GRAmtgrpBI) / cdbl(prodkgBI)
								   end if	  
								   if cdbl(budprodkgBI) = 0 then
								      grandtot_rp_kg_budBI = 0
								   else
								      grandtot_rp_kg_budBI = cdbl(GRbudgrpBI) / cdbl(budprodkgBI)
								   end if
								   
								   if cdbl(prodkgSBI ) = 0 then
								      grandtot_rp_kg_SBI  = 0
								   else
								      grandtot_rp_kg_SBI  = cdbl(GRAmtgrpSBI) / cdbl(prodkgSBI )
								   end if
								   
								   if cdbl(budProdkgSBI)=0 then
								      grandtot_rp_kg_budSBI =  0
								   else
								      grandtot_rp_kg_budSBI =  cdbl(GRbudgrpSBI) / cdbl(budProdkgSBI)
								   end if
		 %>
		<td style="background-color:#F4F4A0" colspan="3"><b> Grand Total&nbsp;</td>	
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,0)%></td>
		<%  if   request("vStatus") = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_kg_BI),0) %></b></td>
		<% else  %> 
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_ha_BI),0) %></b></td>
		<% end if %>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudgrpBI,0)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_kg_budBI),0) %></b></td>
		<% else  %> 
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_ha_budBI),0) %></b></td>
		<% end if %>
		
		
		<% ' prosentase BI %>
		<% IF cdbl(grandtot_rp_kg_budBI) = 0  and cdbl(grandtot_rp_kg_BI) > 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>-100</td>
		<% Elseif cdbl(grandtot_rp_kg_budBI) > 0  and cdbl(grandtot_rp_kg_BI) = 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>100</td>
		<% Elseif cdbl(grandtot_rp_kg_budBI) = 0  and cdbl(grandtot_rp_kg_BI) = 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>0</td>  
		<% else %>   
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( (cdbl(grandtot_rp_kg_budBI) - cdbl(grandtot_rp_kg_BI) )/cdbl(grandtot_rp_kg_budBI)*100,2) %></td>
		<% End If %>
		
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,0)%></td>
		<%  if   request("vStatus") = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_kg_SBI),0) %></b></td>
		<% else  %> 
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_ha_SBI),0) %></b></td>
		<% end if %>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudgrpSBI,0)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_kg_budSBI),0) %></b></td>
		<% else  %> 
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_ha_budSBI),0) %></b></td>
		<% end if %>
		
		
		<% ' prosentase SBI %>
		<% IF cdbl(grandtot_rp_kg_budSBI) = 0  and cdbl(grandtot_rp_kg_SBI) > 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>-100</td>
		<% Elseif cdbl(grandtot_rp_kg_budSBI) > 0  and cdbl(grandtot_rp_kg_SBI) = 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>100</td>
		<% Elseif cdbl(grandtot_rp_kg_budSBI) = 0  and cdbl(grandtot_rp_kg_SBI) = 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>0</td>  
		<% else %>   
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( (cdbl(grandtot_rp_kg_budSBI) - cdbl(grandtot_rp_kg_SBI) )/cdbl(grandtot_rp_kg_budSBI)*100,2) %></td>
		<% End If %>
		
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudgrpthn,0)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(GRbudgrpthn) / cdbl(budProdkgthn),0) %></b></td>
		<% else  %> 
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(GRbudgrpthn) / cint(luasHA),0) %></b></td>
		<% end if %>			
	<tr>		
</table>
<!--#include file="bottom.Asp"-->