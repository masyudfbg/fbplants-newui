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
<title>iPla'S Reporting - Analisa Biaya Pabrik</title>
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
	<!--#include file="FromSelectHDRMILL.html"-->
	
	
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
	    <% Dim ia,rekap
		   ia = request.Form("vMonth")
		   if   request.Form("vJenisReport") = "Group"  then
		        rekap = "REKAP BIAYA PER STASIUN"
		   else		
				rekap = "REKAP BIAYA PER STASIUN PER PEKERJAAN"
		   end if
				
		%>   
		<h3>ANALISA BIAYA LANGSUNG PABRIK</h3>
		<% IF request.Form("vMonth") <> "" then %>
		<h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=request.Form("vYear")%></h5>
		<% END IF %>
		<h5><%=rekap%></h5> 
	</div>
	</div>
<%
Dim strtbs, rsha,TBSprocess,ProdCPO ,ProdKER,Palm, strprod
TBSprocess=0
strtbs = " select  sum(t.cageweight) TBS_PROCESS from SHIFTREPORT t " & _
         " where to_char(t.productiondate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.productiondate,'MM') = '" & request.Form("vMonth") & "'" & _
         " group by to_char(t.productiondate,'YYYYMM') " 		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    TBSprocess = rsha("TBS_PROCESS")
end if
rsha.close()

ProdCPO =0

strtbs = " select  sum(t.quantity) Produksi_CPO from PRODUCTSTORAGEREADING t " & _
         " where t.pstorage_storagecode = 'CIP1' and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = '" & request.Form("vMonth") & "'" & _
         " group by to_char(t.readingdate,'YYYYMM') " 		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    ProdCPO = rsha("Produksi_CPO")
end if
rsha.close()

ProdKER =0


strtbs = " select  sum(t.quantity) Produksi_KER from PRODUCTSTORAGEREADING t " & _
         " where t.pstorage_storagecode = 'KIP1'  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = '" & request.Form("vMonth") & "'" & _
		 " group by to_char(t.readingdate,'YYYYMM') " 		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    ProdKER = rsha("Produksi_KER")
end if
rsha.close()
Palm = cdbl(ProdKER) +  cdbl(ProdCPO)
%>	

<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="2">Job Code</th>
		<th rowspan="2">Station</th>
		<th colspan="6">BULAN INI</th>
		<th colspan="6">S/D BULAN INI</th>
		<th colspan="3" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		
		<th colspan="3">REALISASI</th>
		
		<th colspan="3">ANGGARAN</th>
		
		<th colspan="3">REALISASI</th>
		<th colspan="3">ANGGARAN</th>
		
	</tr>
	<tr>
	    <th rowspan="6">&nbsp;</th>
		<th  align="right">TBS PROSES (KG)</th>
		<th colspan="3"><%=formatnumber(TBSprocess,2)%></th>
		<th colspan="3">0</th>
		<th colspan="3"><%=formatnumber(TBSprocess,2)%></th>
		<th colspan="3">0</th>
		<th colspan="3">0</th>
	</tr>
	<tr>
	    
		<th  align="right">PRODUKSI CPO (KG)</th>
		<th colspan="3"><%=formatnumber(ProdCPO,2)%></th>
		<th colspan="3">0</th>
		<th colspan="3"><%=formatnumber(ProdCPO,2)%></th>
		<th colspan="3">0</th>
		<th colspan="3">0</th>
	</tr>
	<tr>
	    
		<th  align="right">PRODUKSI KER (KG)</th>
		<th colspan="3"><%=formatnumber(ProdKER,2)%></th>
		<th colspan="3">0</th>
		<th colspan="3"><%=formatnumber(ProdKER,2)%></th>
		<th colspan="3">0</th>
		<th colspan="3">0</th>
	</tr>
	<tr>
	    
		<th  align="right">PALM (KG)</th>
		<th colspan="3"><%=formatnumber(Palm,2)%></th>
		<th colspan="3">0</th>
		<th colspan="3"><%=formatnumber(Palm,2)%></th>
		<th colspan="3">0</th>
		<th colspan="3">0</th>
	</tr>
	<tr>
		
		<th rowspan="2">&nbsp;</th>
		<th rowspan="2">Total</th>
		<th colspan="2">Rp./Kg.</th>
		
		<th rowspan="2">Total</th>
		<th colspan="2">Rp./Kg.</th>
		<th rowspan="2">Total</th>
		<th colspan="2">Rp./Kg.</th>
		<th rowspan="2">Total</th>
		<th colspan="2">Rp./Kg.</th>
		<th rowspan="2">Total</th>
		<th colspan="2">Rp./Kg.</th>
	</tr>
	<tr>
		
		<th>Palm</th>
		<th>TBS</th>
		
		<th>Palm</th>
		<th>TBS</th>
		
		<th>Palm</th>
		<th>TBS</th>
		
		<th>Palm</th>
		<th>TBS</th>
		
		<th>Palm</th>
		<th>TBS</th>
	</tr>
	
	<%
	Dim strdtl , strgrp ,cdgrp,cdgrpdesc  ,noline , subAmtgrp , GRAmtgrp , grpDesc
	 if   request.Form("vJenisReport") = "Job"  then
			strdtl = " select t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  DESCRIPTION , sum(t.amount) Total from VIEW_ACCOUNTREPORT_MILL t " & _
				     " where t.period_year = '" & request.Form("vYear") & "' and t.period_month = '" & request.Form("vMonth") & "'" 
					
			strgrp =  " group by t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION order by t.stationcode"	
	else  	
           	strdtl = " select t.stationcode jobcode, t.STATIONCODE_DESCRIPTION DESCRIPTION, sum(t.amount) Total from VIEW_ACCOUNTREPORT_MILL t " & _
				     " where t.period_year = '" & request.Form("vYear") & "' and t.period_month = '" & request.Form("vMonth") & "'" 
					
			strgrp =  " group by t.stationcode, t.STATIONCODE_DESCRIPTION  order by t.stationcode"	
			 
	end if		 
			 
				  strsql1 = "  and t.ba='" & session("fcba")  & "'" 
			
		
		
			 strdtl =  strdtl  + strsql1  + strgrp 
	
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    
	  if cdbl(Palm) = 0 then
		      Palm = "1"
		  end if
	  if cdbl(TBSprocess) = 0 then
		      TBSprocess= "1"
	end if
	subAmtgrp ="0"
	if Not rsha.eof then
		  while Not rsha.eof
		  
	
         IF  request.Form("vJenisReport") = "Group" then 
		     GRAmtgrp = cdbl(GRAmtgrp) + cdbl(rsha("Total")) %>
			 
	      <tr>  
				<td><%=rsha("JobCode")%></td>
				<td><%=rsha("DESCRIPTION")%></td>
				<td align="right"><%=formatnumber( cdbl(rsha("Total")),2)%></td>
				<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(palm)),2)%></td>
				<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(TBSprocess)),2)%></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td align="right"><%=formatnumber( cdbl(rsha("Total")),2)%></td>
				<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(palm)),2)%></td>
				<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(TBSprocess)),2)%></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				
		    <tr>		
	   <% ELSE 
	            
				   GRAmtgrp = cdbl(GRAmtgrp) + cdbl(rsha("Total"))
				   
				   IF  cdgrp = rsha("stationcode") then 
				         subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("total"))%>
						
						<tr>  
							<td><%=rsha("JobCode")%></td>
							<td><%=rsha("DESCRIPTION")%></td>
							<td align="right"><%=formatnumber( cdbl(rsha("Total")),2)%></td>
							<td align="right"><%=formatnumber((cdbl(rsha("Total"))/cdbl(palm)),2)%></td>
							<td align="right"><%=formatnumber(( cdbl(rsha("Total"))/cdbl(TBSprocess)),2)%></td>		
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td align="right"><%=formatnumber( cdbl(rsha("Total")),2)%></td>
							<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(palm)),2)%></td>
							<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(TBSprocess)),2)%></td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							
						<tr>		
				  <% Else  %>
				         
				         <%  if  cdgrp <> "" then  %>
								<tr>  
									
									<td colspan="2"><b>Total &nbsp;<%=cdgrpdesc%></td>
									<td align="right"><b><%=formatnumber(subAmtgrp,2)%></td>
									<td align="right"><b><%=formatnumber((cdbl(subAmtgrp)/cdbl(palm)),2)%></td>
									<td align="right"><b><%=formatnumber(( cdbl(subAmtgrp)/cdbl(TBSprocess)),2)%></td>	
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td align="right"><b><%=formatnumber(subAmtgrp,2)%></td>
									<td align="right"><b><%=formatnumber((cdbl(subAmtgrp)/cdbl(palm)),2)%></td>
									<td align="right"><b><%=formatnumber(( cdbl(subAmtgrp)/cdbl(TBSprocess)),2)%></td>	
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>								
								<tr>
                          <% end if %>
                          <tr>  
								<td>&nbsp;</td>
								<td><b><%=rsha("stationcode")%>. <%=rsha("STATIONCODE_DESCRIPTION")%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>	
						  <tr> 
						  <tr>  
							<td><%=rsha("JobCode")%></td>
							<td><%=rsha("DESCRIPTION")%></td>
							<td align="right"><%=formatnumber( cdbl(rsha("Total")),2)%></td>
							<td align="right"><%=formatnumber((cdbl(rsha("Total"))/cdbl(palm)),2)%></td>
							<td align="right"><%=formatnumber(( cdbl(rsha("Total"))/cdbl(TBSprocess)),2)%></td>		
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td align="right"><%=formatnumber( cdbl(rsha("Total")),2)%></td>
							<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(palm)),2)%></td>
							<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(TBSprocess)),2)%></td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							
						<tr>
						  
                          <% 
						     subAmtgrp = cdbl(rsha("Total"))						  
				     End If  
				     cdgrp = rsha("stationcode")
				     cdgrpdesc = rsha("STATIONCODE_DESCRIPTION")  
		 End If  		  
	     
		
		rsha.MoveNext
		wend
	  
		IF  request.Form("vJenisReport") = "Job" then %>
			<tr>  							
				<td colspan="2"><b>Total &nbsp;<%=cdgrpdesc%></td>
				<td align="right"><b><%=formatnumber(subAmtgrp,2)%></td>
				<td align="right"><b><%=formatnumber((cdbl(subAmtgrp)/cdbl(palm)),2)%></td>
				<td align="right"><b><%=formatnumber(( cdbl(subAmtgrp)/cdbl(TBSprocess)),2)%></td>	
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td align="right"><b><%=formatnumber(subAmtgrp,2)%></td>
				<td align="right"><b><%=formatnumber((cdbl(subAmtgrp)/cdbl(palm)),2)%></td>
				<td align="right"><b><%=formatnumber(( cdbl(subAmtgrp)/cdbl(TBSprocess)),2)%></td>	
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>														
			<tr>
	   <%END IF 
	end if  
	
	
	if   request.Form("vJenisReport") = "Job"  then %>
	         <tr>  
				<td>&nbsp;</td>
				<td><b>Rekapitulasi</td>
				<td >&nbsp;</td>
				<td >&nbsp;</td>
				<td >&nbsp;</td
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			<tr>
     <%			
			strdtl = " select t.activity_heading, t.JobCode,t.ACCOUNTCODE_DESCRIPTION DESCRIPTION, sum(t.amount) Total from VIEW_ACCOUNTREPORT_MILL t " & _
				     " where t.period_year = '" & request.Form("vYear") & "' and t.period_month = '" & request.Form("vMonth") & "'" 
					
			strgrp =  " group by  t.activity_heading,t.JobCode,t.ACCOUNTCODE_DESCRIPTION order by t.JobCode"	
	
			
				strsql1 = "  and t.ba='" & session("fcba")  & "'" 
			
			strdtl =  strdtl  + strsql1  + strgrp 
			
			Set rsha = Server.CreateObject("ADODB.Recordset")
			rsha.Open strdtl, connstr, adOpenStatic 
			if Not rsha.eof then
				  while Not rsha.eof 
			 %>
					 <tr>  
						<td><%=rsha("JobCode")%></td>
						<td><%=rsha("DESCRIPTION")%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Total")),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(palm)),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(TBSprocess)),2)%></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td align="right"><%=formatnumber( cdbl(rsha("Total")),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(palm)),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("Total"))/cdbl(TBSprocess)),2)%></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
					
					<tr>	
			 <%
					 rsha.MoveNext
				  wend
			end if
	end if		
	%>
	
	
	<tr>  							
		<td colspan="2"><b> Grand Total</td>
		<td align="right"><b><%=formatnumber(GRAmtgrp,2)%></td>
		<td align="right"><b><%=formatnumber((cdbl(GRAmtgrp)/cdbl(palm)),2)%></td>
		<td align="right"><b><%=formatnumber(( cdbl(GRAmtgrp)/cdbl(TBSprocess)),2)%></td>	
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td align="right"><b><%=formatnumber(GRAmtgrp,2)%></td>
		<td align="right"><b><%=formatnumber((cdbl(GRAmtgrp)/cdbl(palm)),2)%></td>
		<td align="right"><b><%=formatnumber(( cdbl(GRAmtgrp)/cdbl(TBSprocess)),2)%></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp; </td>														
	<tr>   
</table>
<!--#include file="bottom.Asp"-->