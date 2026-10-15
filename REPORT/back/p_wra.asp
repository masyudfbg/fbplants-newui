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
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<!--#include file="FromSelectHDRVRA.html"-->
	
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
		<div id="subheader" class="centered">
			<% Dim ia
			   ia = request.Form("vMonth")
			%>   
			<h3>LAPORAN WORKSHOP RUNNING ACCOUNT (WRA)</h3>
			<% IF request.Form("vMonth") <> "" then %>
			<h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=request.Form("vYear")%></h5>
			<% END IF %>
			
		</div>
	</div>
<%
  Dim cur_month,strbud_BI,strbud_SBI,x1,strbudprod_BI,strbudprod_SBI
	    strbud_BI = "t.Budget_01" 
	    strbud_SBI = "t.Budget_01"
		strbudprod_BI = "t.qty_01"
		strbudprod_SBI = "t.qty_01"
		
		cur_month = request.Form("vMonth") 
		
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
  strhdr  = " select fcba,sum(JK_BI) JK_BI,sum(JK_SBI)  JK_SBI,sum(JK_bud_BI) JK_bud_BI,sum(JK_bud_SBI) JK_bud_SBI,sum(JK_bud_THN) JK_bud_THN from ( " & _
			"	select wa.fcba,sum(wa.unit) JK_BI,0 JK_SBI,0 JK_bud_BI,0 JK_bud_SBI,0 JK_bud_THN from workshopactivity wa  " & _
			"	where to_char(wa.workshopdate,'YYYY')='" & request.Form("vYear") & "' and to_char(wa.workshopdate,'MM')=lpad('" & request.Form("vMonth") & "',2,0) and wa.fcba='" & session("fcba") & "'" & _
			"	group by wa.fcba " & _
			"	union " & _
			"	select wa.fcba,0 JK_BI,sum(wa.unit)  JK_SBI,0 JK_bud_BI,0 JK_bud_SBI,0 JK_bud_THN from workshopactivity wa   " & _
			"	where to_char(wa.workshopdate,'YYYY')='" & request.Form("vYear") & "' and to_char(wa.workshopdate,'MM')<=lpad('" & request.Form("vMonth") & "',2,0) and  wa.fcba='" & session("fcba") & "'" & _
			"	group by wa.fcba " & _
			"	union " & _
			"	select t.fcba,0 JK_BI,0  JK_SBI,sum(" & strbudprod_BI & ") JK_bud_BI,0 JK_bud_SBI,0 JK_bud_THN from BUDGET_TRANSACTION_WSHOUR t " & _
			"	where t.fcba='" & session("fcba") & "'and t.budget_year='" & request.Form("vYear") & "' " & _
			"	group by t.fcba " & _
			"	union " & _
			"	select t.fcba,0 JK_BI,0  JK_SBI,0 JK_bud_BI,sum(" & strbudprod_SBI & ") JK_bud_SBI,0 JK_bud_THN from BUDGET_TRANSACTION_WSHOUR t " & _
			"	where t.fcba='" & session("fcba") & "'and t.budget_year='" & request.Form("vYear") & "' " & _
			"	group by t.fcba " & _
			"	union " & _
			"	select t.fcba,0 JK_BI,0  JK_SBI,0 JK_bud_BI,0 JK_bud_SBI,sum(t.qty_01+t.qty_02+t.qty_03+t.qty_04+t.qty_05+t.qty_06+t.qty_07+t.qty_08+t.qty_09+t.qty_10+t.qty_11+t.qty_12) JK_bud_THN from BUDGET_TRANSACTION_WSHOUR t " & _
			"	where t.fcba='" & session("fcba") & "'and t.budget_year='" & request.Form("vYear") & "' " & _
			"	group by t.fcba " & _
			"	 ) group by fcba "
Set rshdr = Server.CreateObject("ADODB.Recordset") 
'response.write strhdr
rshdr.Open strhdr, connstr, adOpenStatic
	
if Not rshdr.eof then
   jamBI = rshdr("JK_BI")
   jamSBI = rshdr("JK_SBI")
   jambudBI = rshdr("JK_bud_BI")
   jambudSBI = rshdr("JK_bud_BI")
   jambudthn = rshdr("JK_bud_THN")
end if
rshdr.close

%>

<table border="0" width="100%" id="product-table">
	<tr>
		<th  style="color:#FFFFFF" colspan="2" rowspan="2"></th>
		<th  style="color:#FFFFFF" colspan="5" >BULAN INI</th>
		<th  style="color:#FFFFFF" colspan="5" >S/D BULAN INI</th>
		<th  style="color:#FFFFFF" colspan="2"rowspan="2" >ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
	    <th  style="color:#FFFFFF" colspan="2">REALISASI</th>
		<th  style="color:#FFFFFF" colspan="2">ANGGARAN</th>
		<th  style="color:#FFFFFF" >%</th>
		
		<th  style="color:#FFFFFF" colspan="2">REALISASI</th>
		<th  style="color:#FFFFFF" colspan="2">ANGGARAN</th>
		<th  style="color:#FFFFFF" >%</th>
	</tr>

	<tr>
		
		<th style="color:#FFFFFF" colspan="2">JAM KERJA BENGKEL (JAM)</th>
		<th style="color:#FFFFFF" colspan="2"><%=formatnumber(cdbl(jamBI),2)%></th>
		<th style="color:#FFFFFF" colspan="2"><%=formatnumber(cdbl(jambudBI),2)%></th>
		<% IF cdbl(jamBI) = 0 THEN %>
		    <th style="color:#FFFFFF" >100</th>
		<% ELSEIF cdbl(jambudBI) = 0 THEN %>
		    <th style="color:#FFFFFF" >-100</th>
		<% ELSEIF cdbl(jambudBI) = 0 and cdbl(jamBI) = 0 THEN %>
		    <th style="color:#FFFFFF" >0</th>
        <% ELSEIF cdbl(jambudBI) > 0 and cdbl(jamBI) > 0 THEN %>
		    <th style="color:#FFFFFF" ><%=formatnumber((cdbl(jamBI) - cdbl(jambudBI))/cdbl(jambudBI)*100,2)%></th>
        <% END IF %>			
		<th style="color:#FFFFFF" colspan="2"><%=formatnumber(cdbl(jamSBI),2)%></th>
		<th style="color:#FFFFFF" colspan="2"><%=formatnumber(cdbl(jambudSBI),2)%></th>
		<% IF cdbl(jamSBI) = 0 THEN %>
		    <th style="color:#FFFFFF" >100</th>
		<% ELSEIF cdbl(jambudSBI) = 0 THEN %>
		    <th style="color:#FFFFFF" >-100</th>
		<% ELSEIF cdbl(jambudSBI) = 0 and cdbl(jamSBI) = 0 THEN %>
		    <th style="color:#FFFFFF" >0</th>
        <% ELSEIF cdbl(jambudSBI) > 0 and cdbl(jamSBI) > 0 THEN %>
		    <th style="color:#FFFFFF" ><%=formatnumber((cdbl(jamSBI) - cdbl(jambudSBI))/cdbl(jambudSBI)*100,2)%></th>
        <% END IF %>
		
		<th style="color:#FFFFFF" colspan="2"><%=formatnumber(cdbl(jambudthn),2)%></th>
	</tr>
	<tr>
		
		<th style="color:#FFFFFF" >KODE</th>
		<th style="color:#FFFFFF" >NAMA PEKERJAAN</th>
		
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Jam</th>		
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Jam</th>
		<th style="color:#FFFFFF" ></th>
		
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Jam</th>		
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Jam</th>
		<th style="color:#FFFFFF" ></th>
		
		<th style="color:#FFFFFF" >Total</th>
		<th style="color:#FFFFFF" >Rp/Jam</th>
		
		
	</tr>
	<%
	Dim rsha, strdtl ,rpperjamBI,rpperjamSBI,rpperjambudBI,rpperjambudSBI,rpperjambudTHN
	Dim subtotBI, subtotSBI,subrpperjamBI,subrpperjamSBI, subtotbudBI, subtotbudSBI,subrpperjambudBI,subrpperjambudSBI,subtotbudTHN,subrpperjambudTHN
	Dim subrasioBI,subactualBI ,subactualSBI ,subtotactualBI ,subtotactualSBI 
	
			strdtl = " select jobcode,job_api.get_jobcode_name(jobcode,fcba) JobName, sum(totBI) totBI,sum(totSBI) totSBI,sum(totbudBI) totbudBI,sum(totbudSBI) totbudSBI ,sum(totbudTHN) totbudTHN  from ( " & _
					 "	select ca.fcba,ca.jobcode, sum(ca.amount) totBI,0 totSBI,0 totbudBI,0 totbudSBI,0 totbudTHN  from cost_activity ca  " & _
					 "	where ca.targettype='WS' and ca.fcba='" & session("fcba") & "'and to_char(ca.tdate,'YYYY')='" & request.Form("vYear") & "' and to_char(ca.tdate,'MM')=lpad('" & request.Form("vMonth") & "',2,0) " & _
					 "	group by ca.fcba,ca.jobcode " & _
					 "	union " & _
					 "	select ca.fcba,ca.jobcode, 0 totBI,sum(ca.amount) totSBI,0 totbudBI,0 totbudSBI,0 totbudTHN  from cost_activity ca  " & _
					 "	where ca.targettype='WS' and ca.fcba='" & session("fcba") & "'and to_char(ca.tdate,'YYYY')='" & request.Form("vYear") & "' and to_char(ca.tdate,'MM')<=lpad('" & request.Form("vMonth") & "',2,0)  " & _
					 "	group by ca.fcba,ca.jobcode  " & _
					 "	union " & _
					 "	select t.fcba,t.jobcode ,0 totBI,0 totSBI,sum( " & strbud_BI & ") totbudBI,0 totbudSBI,0 totbudTHN from BUDGET_TRANSACTION_WORKSHOP t " & _
					 "	where t.budget_year='" & request.Form("vYear") & "' and  t.fcba='" & session("fcba") & "' " & _
					 "	group by t.fcba,t.jobcode " & _
					 "	union " & _
					 "	select t.fcba,t.jobcode ,0 totBI,0 totSBI,0 totbudBI,sum(" & strbud_SBI & ") totbudSBI,0 totbudTHN from BUDGET_TRANSACTION_WORKSHOP t " & _
					 "	where t.budget_year='" & request.Form("vYear") & "' and  t.fcba='" & session("fcba") & "' " & _
					 "	group by t.fcba,t.jobcode " & _
					 "	union " & _
					 "	select t.fcba,t.jobcode ,0 totBI,0 totSBI,0 totbudBI,0 totbudSBI,sum(t.budget_01+t.budget_02+t.budget_03+t.budget_04+t.budget_05+t.budget_06+t.budget_07+t.budget_08+t.budget_09+t.budget_10+t.budget_11+t.budget_12) totbudTHN from BUDGET_TRANSACTION_WORKSHOP t " & _
					 "	where t.budget_year='" & request.Form("vYear") & "' and  t.fcba='" & session("fcba") & "' " & _
					 "	group by t.fcba,t.jobcode " & _
					 "	) group by fcba,jobcode "	
			
	
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
	'response.write strdtl
    'response.write request.Form("vMonth") &  "  " & request.Form("vYear") & " " & session("fcba") 
	if Not rsha.eof then
		  while Not rsha.eof
		    
			  %>
				<tr>  
					<td align="CENTER"><%=rsha("jobcode")%></td>
					<td align="LEFT"><%=rsha("JobName")%></td>
					<td align="right"><%=formatnumber(cdbl(rsha("totBI")),2)%></td>
					<% if cdbl(jamBI) > 0 then
					      rpperjamBI = cdbl(rsha("totBI"))/ cdbl(jamBI)
					   else 
						  rpperjamBI = "0" 
					   end if %>
					<td align="right"><%=formatnumber(cdbl(rpperjamBI),2)%></td>
					   
					<td align="right"><%=formatnumber(cdbl(rsha("totbudBI")),2)%></td>
					
					<% if cdbl(jambudBI) > 0 then
					      rpperjambudBI = cdbl(rsha("totbudBI"))/ cdbl(jambudBI)
					  else 
						  rpperjambudBI = "0" 
					  end if %>
					<td align="right"><%=formatnumber(cdbl(rpperjambudBI),2)%></td>
					
					<% 'persen BI %>
					<% IF cdbl(rpperjamBI) = 0 and cdbl(rpperjambudBI) > 0 THEN %>
					  <td align="right">100</td>
					<% ELSEIF cdbl(rpperjambudBI) > 0 and cdbl(rpperjambudBI) = 0 THEN %>
					  <td align="right">-100</td>
					<% ELSEIF cdbl(rpperjamBI) = 0 and cdbl(rpperjambudBI) = 0 THEN %>
					  <td align="right">0</td>
					<% ELSEIF cdbl(rpperjamBI) > 0 and cdbl(rpperjambudBI) > 0 THEN %>
					  <td align="right"><%=formatnumber((cdbl(rpperjambudBI) - cdbl(rpperjamBI)) / cdbl(rpperjambudBI)*100,2)%></td>
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
					<% ELSEIF cdbl(rpperjambudSBI) > 0 and cdbl(rpperjambudSBI) = 0THEN %>
					  <td align="right">-100</td>
					<% ELSEIF cdbl(rpperjamSBI) = 0 and cdbl(rpperjambudSBI) = 0 THEN %>
					  <td align="right">0</td>
					<% ELSEIF cdbl(rpperjamSBI) > 0 and cdbl(rpperjambudSBI) > 0 THEN %>
					  <td align="right"><%=formatnumber((cdbl(rpperjambudSBI) - cdbl(rpperjamSBI)) / cdbl(rpperjambudSBI) *100,2)%></td>
                    <% END IF %> 
					
					<td align="right"><%=formatnumber(cdbl(rsha("totbudTHN")),2)%> </td>
					<% if cdbl(jambudTHN) > 0 then
					      rpperjambudTHN = cdbl(rsha("totbudTHN"))/ cdbl(jambudTHN)
					  else 
						  rpperjambudTHN = "0" 
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
       <td style="background-color:#F4F4A0" align="center" colspan="2"><b>TOTAL</td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrpperjamBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotbudBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrpperjambudBI),2)%></td>
	   
	    <% 'persen BI %>
	   <% IF cdbl(subrpperjamBI) = 0 and cdbl(subrpperjambudBI) > 0 THEN %>
		<td style="background-color:#F4F4A0" align="right"><b>100</td>
	   <% ELSEIF cdbl(subrpperjambudBI) > 0 and cdbl(subrpperjambudBI) = 0THEN %>
		<td style="background-color:#F4F4A0" align="right"><b>-100</td>
	   <% ELSEIF cdbl(subrpperjamBI) = 0 and cdbl(subrpperjambudBI) = 0 THEN %>
		 <td style="background-color:#F4F4A0" align="right"><b>0</td>
	   <% ELSEIF cdbl(subrpperjamBI) > 0 and cdbl(subrpperjambudBI) > 0 THEN %>
		 <td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subrpperjambudBI) - cdbl(subrpperjamBI)) / cdbl(subrpperjambudBI) *100,2)%></td>
        <% END IF %>
	   
       <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotSBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrpperjamSBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotbudSBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrpperjambudSBI),2)%></td>
	   
	   <% 'persen SBI %>
	   <% IF cdbl(subrpperjamSBI) = 0 and cdbl(subrpperjambudSBI) > 0 THEN %>
		<td style="background-color:#F4F4A0" align="right"><b>100</td>
	   <% ELSEIF cdbl(subrpperjambudSBI) > 0 and cdbl(subrpperjambudSBI) = 0THEN %>
		<td style="background-color:#F4F4A0" align="right"><b>-100</td>
	   <% ELSEIF cdbl(subrpperjamSBI) = 0 and cdbl(subrpperjambudSBI) = 0 THEN %>
		 <td style="background-color:#F4F4A0" align="right"><b>0</td>
	   <% ELSEIF cdbl(subrpperjamSBI) > 0 and cdbl(subrpperjambudSBI) > 0 THEN %>
		 <td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subrpperjambudSBI) - cdbl(subrpperjamSBI)) / cdbl(subrpperjambudSBI) *100,2)%></td>
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