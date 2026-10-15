<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Laporan Bibit</title>
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
		<form name="frmBibit" method="post" action="<%=ScriptName%>">
			<tr>				
				<td><font face="verdana" size="1">TAHUN</font></td>
				<td>
<% 
    MM=Request.Form("vMonth")
    yyyy=Request.Form("vYear")

    mon = Month(Session("v_Costbook"))
    If len(MM)=0 then MM=mon

    Yr = Year(Session("v_Costbook"))
    MinYr = Yr - 2
    MaxYr = Yr  + 2
    If len(yyyy)= 0 then yyyy = Yr
%>
				<select name="vYear">
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
					<select name="vMonth">
<%
    For i = 1 to 12
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=Bulan(i)%></option>
<%
    Next
%>
					</select>
			</td>
			<td><font face="verdana" size="1">BATCH</font></td>
			<td>
				<select name="vBatch">
				<option value="ALL" <%if request("vBatch")="ALL" then%> selected <%End if%>>::ALL::</option>
				<%
				Dim rsBatch
				Set rsBatch = Server.CreateObject("ADODB.Recordset")
				rsBatch.CursorLocation = adUseClient
				strsql =  " SELECT NURSERYCODE FCCODE, DESCRIPTION FCNAME FROM NURSERY WHERE FCBA = '" & Session("Fcba") & "' ORDER BY NURSERYCODE ASC"
				rsBatch.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
				if not rsBatch.eof then
				DO while not rsBatch.eof
				%>			
				<option value="<%=rsBatch("FCCODE")%>" <%if request("vBatch")= rsBatch("FCCODE") then%> selected <%End If%>><%=rsBatch("FCCODE")%></option>
				<%rsBatch.movenext
				loop
				End if
				rsBatch.Close
				set rsBatch = Nothing
				%>
				</select>
			</td>
			<td><input type="submit" value="View" /></td>
		</tr>
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

	Dim rsDescBa, strDescBa, vCompanyName, vFcbaName, vManagerName, vKTUName
	strDescBa	=	" SELECT * FROM BUSINESSUNIT WHERE FCCODE = '" & Session("Fcba") & "'"
	Set rsDescBa = Server.CreateObject("ADODB.Recordset")
	rsDescBa.Open strDescBa, connstr, adOpenStatic
	if Not rsDescBa.eof then
		vCompanyName	=	"PT. " & rsDescBa("FCCOMPANYNAME")
		vFcbaName		=	rsDescBa("FCNAME")
		vManagerName	=	rsDescBa("FCMANAGER")
		vKTUName		=	rsDescBa("FCKTU")
	end if
	rsDescBa.Close
	set rsDescBa = Nothing
	
	Dim rsActualBibit, strActualBibit, vActualBibit, vBudgetBibit, rsBudgetBibit, strBudgetBibit
	strActualBibit	=	" SELECT NVL(((SUM(OPENINGQTY)+SUM(QTYRECEIVED))-(SUM(QTYCULLED)+SUM(QTYISSUED))),0) ACTUALBIBIT" & _
						" FROM NURSERYLEDGER WHERE FCBA = '" & Session("Fcba")  & "' AND YEAR = " & yyyy & " and MONTH = " & MM
	strBudgetBibit	=	" SELECT nvl(SUM (budgetbibit),0) budgetbibit" & _
						"   FROM (  SELECT nyear," & _
						"                  nurserycode," & _
						"                  SUM (budgetbibit) budgetbibit," & _
						"                  fcba" & _
						"             FROM (  SELECT BUDGET_YEAR nyear," & _
						"                            BATCHCODE nurserycode," & _
						"                            QTYNURESRY budgetbibit," & _
						"                            FCBA" & _
						"                       FROM BUDGET_TRANSACTION_NURSERY" & _
						"                   GROUP BY BUDGET_YEAR," & _
						"                            BATCHCODE," & _
						"                            QTYNURESRY," & _
						"                            FCBA)" & _
						"         GROUP BY nyear, nurserycode, fcba)" & _
						"  WHERE FCBA = '" & Session("Fcba")  & "' AND nYEAR = " & yyyy

	if  request("vBatch") <> "ALL" then
	  strActualBibit = strActualBibit & " AND NURSERYCODE = '" & request("vBatch")  & "'"
	  strBudgetBibit = strBudgetBibit & " AND NURSERYCODE = '" & request("vBatch")  & "'"
	else
	  strActualBibit = strActualBibit & " "
	  strBudgetBibit = strBudgetBibit & " "
	end if
	'response.write strActualBibit
	Set rsActualBibit = Server.CreateObject("ADODB.Recordset")
	rsActualBibit.Open strActualBibit, connstr, adOpenStatic

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
		<div id="header">
			<h4><%=vCompanyName%></h4>
			<h5><%=vFcbaName%></h5>		
		</div>
		<div id="subheader" class="centered">
			<h3>LAPORAN PEMBIAYAAN BIBITAN</h3>
			<h4>Periode&nbsp;:&nbsp;<%=Bulan(mm)%>&nbsp;<%=yyyy%></h4>
		</div>		
	</div>

<table border="0" width="100%" id="product-table">
	<tr>
		<th style="color:#FFFFFF" colspan="2">&nbsp;</th>
		<th style="color:#FFFFFF" colspan="5">BULAN INI</th>
		<th style="color:#FFFFFF" colspan="5">S/D BULAN INI</th>
		<th style="color:#FFFFFF" colspan="2" rowspan="3">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF" colspan="2">&nbsp;</th>
		<th style="color:#FFFFFF" colspan="2">REALISASI</th>
		<th style="color:#FFFFFF" colspan="2">ANGGARAN</th>
		<th style="color:#FFFFFF" rowspan="3">%</th>
		<th style="color:#FFFFFF" colspan="2">REALISASI</th>
		<th style="color:#FFFFFF" colspan="2">ANGGARAN</th>
		<th style="color:#FFFFFF" rowspan="3">%</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF" colspan="2" align="right">JUMLAH BIBIT</th>
		<th style="color:#FFFFFF" colspan="2" align="right"><%=FormatNumber(vActualBibit,2)%></th>
		<th style="color:#FFFFFF" colspan="2" align="right"><%=FormatNumber(vBudgetBibit,2)%></th>
		<th style="color:#FFFFFF" colspan="2" align="right"><%=FormatNumber(vActualBibit,2)%></th>
		<th style="color:#FFFFFF" colspan="2" align="right"><%=FormatNumber(vBudgetBibit,2)%></th>
	</tr>
	<tr>
		<th style="color:#FFFFFF">CODE</th>
		<th style="color:#FFFFFF">JENIS PEKERJAAN</th>
		<th style="color:#FFFFFF">Total</th>
		<th style="color:#FFFFFF">Rp/Pkk</th>
		<th style="color:#FFFFFF">Total</th>
		<th style="color:#FFFFFF">Rp/Pkk</th>
		<th style="color:#FFFFFF">Total</th>
		<th style="color:#FFFFFF">Rp/Pkk</th>
		<th style="color:#FFFFFF">Total</th>
		<th style="color:#FFFFFF">Rp/Pkk</th>
		<th style="color:#FFFFFF">Total</th>
		<th style="color:#FFFFFF">Rp/Pkk</th>
	</tr>
	<%
Dim rsBibit, strBibit,codegroup,codegroupName,rp_pkk_bi,rp_pkk_sbi,rp_pkk_budbi,rp_pkk_budsbi,sub_rp_pkk_bi,sub_rp_pkk_sbi,sub_rp_pkk_budbi,sub_rp_pkk_budsbi,tot_rp_pkk_bi,tot_rp_pkk_sbi,tot_rp_pkk_budbi,tot_rp_pkk_budsbi
Dim subBI,subSBI,totsubBI,totsubSBI,budsubBI,budsubSBI,budSubthn,budtotsubBI,budtotsubSBI,budtotsubTHN
Dim strsql1,strsql2  

If request.Form("vBatch") <> "ALL"  then
		strsql1 = " and ca.targetcode ='" & request.Form("vBatch") & "'" 
		strsql2 = " and bp.batchcode  ='" & request.Form("vBatch") & "'"
	 else
		strsql1 = ""
		strsql2 = ""
end if
strBibit	=   " select fcba,job_api.get_jobgroupCode(jobcode,fcba) jobgroupcode,jobgroup_api.get_jobgroup_name(job_api.get_jobgroupCode(jobcode,fcba),fcba)   jobgroup_name ,jobcode,job_api.get_jobcode_name(jobcode,fcba) Description,sum(BI) BI,sum(SBI) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn from  " & _
				"	( select ca.fcba,ca.jobcode,sum(ca.amount) BI,0 SBI,0 budBI,0 budSBI,0 budthn FROM COST_ACTIVITY ca " & _
				"	  where ca.jobcode in ( SELECT FCCODE FROM JOB WHERE SUBSTR(FCCODE,1,5) in ('13101','13102','13103'))  and to_char(ca.tdate,'YYYY')='" & request.Form("vYear") & "' and to_char(ca.tdate,'MM')=lpad('" & request.Form("vMonth") & "',2,0) and ca.fcba='" & session("fcba") & "' " & strsql1 &  "  " & _
				"	  group by ca.fcba,ca.jobcode " & _
				"	  union " & _
				"	  select ca.fcba,ca.jobcode,0 BI,sum(ca.amount) SBI,0 budBI,0 budSBI,0 budthn FROM COST_ACTIVITY ca " & _
				"	  where ca.jobcode in ( SELECT FCCODE FROM JOB WHERE SUBSTR(FCCODE,1,5) in ('13101','13102','13103') )  and to_char(ca.tdate,'YYYY')='" & request.Form("vYear") & "' and to_char(ca.tdate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0)and ca.fcba='" & session("fcba") & "' " & strsql1 &  "  " & _
				"	  group by ca.fcba,ca.jobcode " & _
				"	  union " & _
				"	  select bp.fcba,bp.jobcode,0 BI,0 SBI,sum(" & strbud_BI & ") budBI,0 budSBI,0 budTHN from BUDGET_TRANSACTION_NURSERY bp  " & _
				"	  where bp.budget_year='" & request.Form("vYear") & "'  and bp.fcba='" & session("fcba") & "' " & strsql2 &  "  " & _
				"	  group by bp.fcba,bp.jobcode " & _
				"	  union " & _
				"	  select bp.fcba,bp.jobcode,0 BI,0 SBI,0 budBI,sum(" & strbud_SBI & ") budSBI,0 budTHN from BUDGET_TRANSACTION_NURSERY bp  " & _
				"	  where bp.budget_year='" & request.Form("vYear") & "'  and bp.fcba='" & session("fcba") & "' " & strsql2 &  "  " & _
				"	  group by bp.fcba,bp.jobcode " & _
				"	  union " & _
				"	  select bp.fcba,bp.jobcode,0 BI,0 SBI,0 budBI,0 budSBI,sum(bp.budget_01+bp.budget_02+bp.budget_03+bp.budget_04+bp.budget_05+bp.budget_06+bp.budget_07+bp.budget_08+bp.budget_09+bp.budget_10+bp.budget_11+bp.budget_12) budTHN from BUDGET_TRANSACTION_NURSERY bp  " & _
				"	  where bp.budget_year='" & request.Form("vYear") & "'  and bp.fcba='" & session("fcba") & "' " & strsql2 &  "  " & _
				"	  group by bp.fcba,bp.jobcode " & _
				"	)  group by fcba,jobcode  order by jobcode"


Set rsBibit = Server.CreateObject("ADODB.Recordset")
rsBibit.Open strBibit, connstr, adOpenStatic
'response.write strBibit
if Not rsBibit.eof then
	While Not rsBibit.eof
			 ' cari rupiah per PKK
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
						<td><%=rsBibit("jobcode")%></td>
						<td><%=rsBibit("Description")%></td>
						<td align="right"><%=FormatNumber(rsBibit("BI"),2)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_bi),2)%></td>
						<td align="right"><%=FormatNumber(rsBibit("budBI"),2)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_budbi),2)%></td>
						<% IF cdbl(rp_pkk_budbi) <> 0 and cdbl(rp_pkk_bi)=0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(rp_pkk_budbi) = 0 and cdbl(rp_pkk_bi)<> 0 then %>  
						  <td align="right">100</td>
						<% ELSEIF cdbl(rp_pkk_budbi) = 0 and cdbl(rp_pkk_bi) = 0 then %>  
						   <td align="right">0</td>
						<% ELSEIF cdbl(rp_pkk_budbi) <> 0 and cdbl(rp_pkk_bi) <> 0 then %>  
						    <td align="right"><%=FormatNumber((cdbl(rp_pkk_budbi)-cdbl(rp_pkk_bi))/cdbl(rp_pkk_budbi)*100,2)%></td>
						<% END IF %>  
						<td align="right"><%=FormatNumber(rsBibit("SBI"),2)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_sbi),2)%></td>
						<td align="right"><%=FormatNumber(rsBibit("budBI"),2)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_budsbi),2)%></td>
						<% IF cdbl(rp_pkk_budsbi) <> 0 and cdbl(rp_pkk_sbi)=0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) = 0 and cdbl(rp_pkk_sbi)<> 0 then %>  
						  <td align="right">100</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) = 0 and cdbl(rp_pkk_sbi) = 0 then %>  
						   <td align="right">0</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) <> 0 and cdbl(rp_pkk_sbi) <> 0 then %>  
						    <td align="right"><%=FormatNumber((cdbl(rp_pkk_budsbi)-cdbl(rp_pkk_sbi))/cdbl(rp_pkk_budsbi)*100,2)%></td>
						<% END IF %> 
						<td align="right"><%=FormatNumber(rsBibit("budTHN"),2)%></td>
						<td align="right"><%=FormatNumber(0,2)%></td>
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
							<td colspan="2" align="left"><strong>Total <%=codegroupName%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(subBI),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_bi),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(budsubBI),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_budbi),2)%></strong></td>
							<% IF cdbl(sub_rp_pkk_budbi) = 0 and cdbl(sub_rp_pkk_bi) <> 0  then %>
							      <td align="right"><strong>-100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) <> 0 and cdbl(sub_rp_pkk_bi) = 0  then %>
							      <td align="right"><strong>100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) = 0 and cdbl(sub_rp_pkk_bi) = 0  then %>
							      <td align="right"><strong>0</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) <> 0 and cdbl(sub_rp_pkk_bi) <> 0  then %>
							      <td align="right"><strong><%=FormatNumber((cdbl(sub_rp_pkk_budbi)-cdbl(sub_rp_pkk_bi))/cdbl(sub_rp_pkk_budbi)*100,2)%></strong></td>
							<% END IF %>	  
							<td align="right"><strong><%=FormatNumber(cdbl(subSBI),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_sbi),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(budsubSBI),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_budsbi),2)%></strong></td>
							<% IF cdbl(sub_rp_pkk_budsbi) = 0 and cdbl(sub_rp_pkk_sbi) <> 0  then %>
							      <td align="right"><strong>-100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) <> 0 and cdbl(sub_rp_pkk_sbi) = 0  then %>
							      <td align="right"><strong>100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) = 0 and cdbl(sub_rp_pkk_sbi) = 0  then %>
							      <td align="right"><strong>0</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) <> 0 and cdbl(sub_rp_pkk_sbi) <> 0  then %>
							      <td align="right"><strong><%=FormatNumber((cdbl(sub_rp_pkk_budsbi)-cdbl(sub_rp_pkk_sbi))/cdbl(sub_rp_pkk_budsbi)*100,2)%></strong></td>
							<% END IF %>	
							<td align="right"><strong><%=FormatNumber(cdbl(budSubthn),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(0,2)%></strong></td>
						</tr>
					<% END IF %> 
					   
					 <tr>
						<td><%=rsBibit("jobcode")%></td>
						<td><%=rsBibit("Description")%></td>
						<td align="right"><%=FormatNumber(rsBibit("BI"),2)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_bi),2)%></td>
						<td align="right"><%=FormatNumber(rsBibit("budBI"),2)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_budbi),2)%></td>
						<% IF cdbl(rp_pkk_budbi) <> 0 and cdbl(rp_pkk_bi)=0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(rp_pkk_budbi) = 0 and cdbl(rp_pkk_bi)<> 0 then %>  
						  <td align="right">100</td>
						<% ELSEIF cdbl(rp_pkk_budbi) = 0 and cdbl(rp_pkk_bi) = 0 then %>  
						   <td align="right">0</td>
						<% ELSEIF cdbl(rp_pkk_budbi) <> 0 and cdbl(rp_pkk_bi) <> 0 then %>  
						    <td align="right"><%=FormatNumber((cdbl(rp_pkk_budbi)-cdbl(rp_pkk_bi))/cdbl(rp_pkk_budbi)*100,2)%></td>
						<% END IF %>  
						<td align="right"><%=FormatNumber(rsBibit("SBI"),2)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_sbi),2)%></td>
						<td align="right"><%=FormatNumber(rsBibit("budBI"),2)%></td>
						<td align="right"><%=FormatNumber(cdbl(rp_pkk_budsbi),2)%></td>
						<% IF cdbl(rp_pkk_budsbi) <> 0 and cdbl(rp_pkk_sbi)=0 then %>
						  <td align="right">-100</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) = 0 and cdbl(rp_pkk_sbi)<> 0 then %>  
						  <td align="right">100</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) = 0 and cdbl(rp_pkk_sbi) = 0 then %>  
						   <td align="right">0</td>
						<% ELSEIF cdbl(rp_pkk_budsbi) <> 0 and cdbl(rp_pkk_sbi) <> 0 then %>  
						    <td align="right"><%=FormatNumber((cdbl(rp_pkk_budsbi)-cdbl(rp_pkk_sbi))/cdbl(rp_pkk_budsbi)*100,2)%></td>
						<% END IF %> 
						<td align="right"><%=FormatNumber(rsBibit("budTHN"),2)%></td>
						<td align="right"><%=FormatNumber(0,2)%></td>
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
							<td colspan="2" align="left"><strong>Total <%=codegroupName%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(subBI),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_bi),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(budsubBI),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_budbi),2)%></strong></td>
							<% IF cdbl(sub_rp_pkk_budbi) = 0 and cdbl(sub_rp_pkk_bi) <> 0  then %>
							      <td align="right"><strong>-100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) <> 0 and cdbl(sub_rp_pkk_bi) = 0  then %>
							      <td align="right"><strong>100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) = 0 and cdbl(sub_rp_pkk_bi) = 0  then %>
							      <td align="right"><strong>0</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budbi) <> 0 and cdbl(sub_rp_pkk_bi) <> 0  then %>
							      <td align="right"><strong><%=FormatNumber((cdbl(sub_rp_pkk_budbi)-cdbl(sub_rp_pkk_bi))/cdbl(sub_rp_pkk_budbi)*100,2)%></strong></td>
							<% END IF %>	  
							<td align="right"><strong><%=FormatNumber(cdbl(subSBI),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_sbi),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(budsubSBI),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(cdbl(sub_rp_pkk_budsbi),2)%></strong></td>
							<% IF cdbl(sub_rp_pkk_budsbi) = 0 and cdbl(sub_rp_pkk_sbi) <> 0  then %>
							      <td align="right"><strong>-100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) <> 0 and cdbl(sub_rp_pkk_sbi) = 0  then %>
							      <td align="right"><strong>100</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) = 0 and cdbl(sub_rp_pkk_sbi) = 0  then %>
							      <td align="right"><strong>0</strong></td>
							<% ELSEIF cdbl(sub_rp_pkk_budsbi) <> 0 and cdbl(sub_rp_pkk_sbi) <> 0  then %>
							      <td align="right"><strong><%=FormatNumber((cdbl(sub_rp_pkk_budsbi)-cdbl(sub_rp_pkk_sbi))/cdbl(sub_rp_pkk_budsbi)*100,2)%></strong></td>
							<% END IF %>	
							<td align="right"><strong><%=FormatNumber(cdbl(budSubthn),2)%></strong></td>
							<td align="right"><strong><%=FormatNumber(0,2)%></strong></td>
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
			<td colspan="2" align="left"><strong>Total Biaya Pembibitan</strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(totsubBI),2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_pkk_bi),2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(budtotsubBI),2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_pkk_budbi),2)%></strong></td>
			<% IF cdbl(tot_rp_pkk_budbi) = 0 and cdbl(tot_rp_pkk_bi) <> 0  then %>
				<td align="right"><strong>-100</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budbi) <> 0 and cdbl(tot_rp_pkk_bi) = 0  then %>
				<td align="right"><strong>100</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budbi) = 0 and cdbl(tot_rp_pkk_bi) = 0  then %>
				<td align="right"><strong>0</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budbi) <> 0 and cdbl(tot_rp_pkk_bi) <> 0  then %>
				<td align="right"><strong><%=FormatNumber((cdbl(tot_rp_pkk_budbi)-cdbl(tot_rp_pkk_bi))/cdbl(tot_rp_pkk_budbi)*100,2)%></strong></td>
			<% END IF %>
			<td align="right"><strong><%=FormatNumber(cdbl(totsubSBI),2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_pkk_sbi),2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(budtotsubSBI),2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(cdbl(tot_rp_pkk_budsbi),2)%></strong></td>
			<% IF cdbl(tot_rp_pkk_budsbi) = 0 and cdbl(tot_rp_pkk_sbi) <> 0  then %>
				<td align="right"><strong>-100</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budsbi) <> 0 and cdbl(tot_rp_pkk_sbi) = 0  then %>
				<td align="right"><strong>100</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budsbi) = 0 and cdbl(tot_rp_pkk_sbi) = 0  then %>
				<td align="right"><strong>0</strong></td>
			<% ELSEIF cdbl(tot_rp_pkk_budsbi) <> 0 and cdbl(tot_rp_pkk_sbi) <> 0  then %>
				<td align="right"><strong><%=FormatNumber((cdbl(tot_rp_pkk_budsbi)-cdbl(tot_rp_pkk_sbi))/cdbl(tot_rp_pkk_budsbi)*100,2)%></strong></td>
			<% END IF %>
			<td align="right"><strong><%=FormatNumber(cdbl(budtotsubTHN),2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(0,2)%></strong></td>
		</tr>

</table>
<%


%>
</div>
</body>
</html>