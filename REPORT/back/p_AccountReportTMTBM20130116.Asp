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
<title>iPla'S Reporting - Account Report Biaya TM/TBM</title>
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
	<!--#include file="FormSelectTMTBM.html"-->
	<!--#include file="menutop.asp"-->
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
		<h2>LAPORAN BIAYA <%=request.Form("vStatus")%></h3>
		<% IF request.Form("vMonth") <> "" then %>
		<h5>PERIODE     : <%=MonthName(cint(ia))%> &nbsp; <%=request.Form("vYear")%></h5>
		<% END IF %>
		<h5>KEPEMILIKAN : <%=request.Form("vOwnership")%></h5>
		<h5>CROP        : <%=request.Form("vCROPTYPE")%></h5>
		<h5>DEVISION    : <%=request.Form("vDivision")%></h5>
		<h5>TAHUN TANAM : <%=thntnm%></h5>
		<h5>SOURCE      : <%=request.Form("vJenisReport")%></h5>
	</div>
	</div>
	
<%
Dim strha, rsha,luasHA,ProdTon ,prodkg, strprod
luasHA=0
strha = " SELECT DECODE(SUM (ha), NULL,1,0,1,SUM (ha)) HA" & _
		" FROM ( SELECT fcba," & _
		" croptype," & _
		" ownership," & _
		" division," & _
		" DX_BLOCK_STATUS as status," & _
		" TO_CHAR(plantingdate,'YYYY') tahuntanam," & _
		" SUM (hectarageplanted) ha" & _
		" FROM field" & _
		" GROUP BY fcba," & _
		" croptype," & _
		" ownership," & _
		" division," & _
		" DX_BLOCK_STATUS," & _
		" plantingdate)" & _
		" WHERE DX_BLOCK_STATUS = '" & request("vStatus")  & "'" & _
		" AND CROPTYPE = '"& request("vCROPTYPE") &"'" 

        strsql1 = "  and fcba='" & session("fcba") & "'" 
       
		  
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
		  
		  strha = strha +  strsql1 + strsql2 +  strsql3 + strsql4	 
		  
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strha, connstr, adOpenStatic

if Not rsha.eof then
    luasHA = rsha("ha")
end if
rsha.close()



ProdTon=0
prodkg=0
if request("vStatus") = "TM" then
 
		strprod = " select nvl(sum(h.mill_weight),0)  prod_ton from harvesting h , field f" & _
				" where h.fieldcode = f.fccode and h.fcba = f.fcba" & _
				
				" and f.STATUS = '" & request("vStatus")  & "'" & _
				" AND f.CROPTYPE = '"& request("vCROPTYPE") &"'" 
				
				  
				  
					  strsql1 = "  and f.fcba='" & session("fcba") & "'" 
				
				  
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
				  if  request.Form("vYear") <>"ALL" then
						  strsql5 = " and to_char(h.receptiondate,'YYYY')='" & request.Form("vYear") & "'" 
				 else
						  strsql5 = " "
				  end if
				if  request.Form("vMonth") <>"ALL" then
						  strsql6 = " and to_char(h.receptiondate,'MM')='" & request.Form("vMonth") & "'" 
				else 
						  strsql6 = " "
				end if
				  strprod = strprod +  strsql1 + strsql2 + strsql3 + strsql4 +  strsql5 + strsql6	 
		Set rsha = Server.CreateObject("ADODB.Recordset")
		rsha.Open strprod, connstr, adOpenStatic
       
		if Not rsha.eof then
			ProdTon = cdbl(rsha("prod_ton"))/1000
			prodkg = rsha("prod_ton")
		end if
		rsha.close()
end if

%>	
    
	
<table border="0" width="100%" id="product-table">
	<tr>
		<th colspan="4">&nbsp;</th>
		<th colspan="5">BULAN INI</th>
		<th colspan="5">S/D BULAN INI</th>
		<th colspan="2" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th colspan="4">&nbsp;</th>
		
		<th colspan="2">REALISASI</th>
		
		<th colspan="2">ANGGARAN</th>
		<th rowspan="4">%</th>
		
		<th colspan="2">REALISASI</th>
		<th colspan="2">ANGGARAN</th>
		<th rowspan="4">%</th>
	</tr>
	<tr>
		<th colspan="4" align="right">LUAS Ha</th>

		<th align="right" colspan="2"><%=luasHA %></th>
		<th align="right" colspan="2">0</th>
		<th align="right" colspan="2"><%=luasHA %></th>
		<th align="right" colspan="2">0</th>
		<th align="right" colspan="2">0</th>
		

	</tr>
	<tr>
		<th colspan="4" align="right">PRODUKSI - Ton</th>

		<th align="right" colspan="2"><%=ProdTon %> </th>
		<th align="right" colspan="2">0</th>
		<th align="right" colspan="2"><%=ProdTon %></th>
		<th align="right" colspan="2">0</th>
		<th align="right" colspan="2">0</th>
		

	</tr>
	<tr>
		<th colspan="4" align="right">Yield - Ton/Ha�</th>

		<th align="right" colspan="2"><%=formatnumber(cdbl(ProdTon)/cdbl(luasHA),2) %></th>
		<th align="right" colspan="2">0</th>
		<th align="right" colspan="2"><%=formatnumber(cdbl(ProdTon)/cdbl(luasHA),2) %></th>
		<th align="right" colspan="2">0</th>
		<th align="right" colspan="2">0</th>
		

	</tr>
	<tr>
		<th>NO</th>
		<th>CODE</th>
		<th colspan="2">JOBGROUP</th>
		<th>Total</th>
		<th>Rp/Ha</th>
		<th>Total</th>
		<th>Rp/Ha</th>
		<th>&nbsp;</th>
		<th>Total</th>
		<th>Rp/Ha</th>
		<th>Total</th>
		<th>Rp/Ha</th>
		<th>&nbsp;</th>
		<th>Total</th>
		<th>Rp/Ha</th>
	</tr>
	<%
	 Dim strdtl , strgrp ,cdgrp ,noline , subAmtgrpBI,subAmtgrpSBI , GRAmtgrpBI,GRAmtgrpSBI , grpDesc
	 
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
			
					  
			strdtl =  " select code,jobheading,job_name,sum(totalBI) totalBI,sum(totalSBI) totalSBI from " & _
					  "  ( select  a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name ,sum(a.tm_amt_actual) totalBI,0 totalSBI from sac_accountreport_tm_tbm a " & _
                      "    where  a.period_year='" & request.Form("vYear") & "' and  a.period_month='" & request.Form("vMonth") & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"'  and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
                      "    GROUP BY  a.activity_heading,a.accountgroupcode 	" & _		
                      "    UNION " & _
                      "    select  a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name ,0 totalBI,sum(a.tm_amt_actual) totalSBI from sac_accountreport_tm_tbm a " & _
                      "    where  a.period_year='" & request.Form("vYear") & "' and  a.period_month <='" & request.Form("vMonth") & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"'  and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
                      "    GROUP BY  a.activity_heading,a.accountgroupcode " & _
                      "  ) group by code,jobheading,job_name "
	 else  
          
					
					
		   strdtl = " select code,jobheading,job_name,sum(totalBI) totalBI,sum(totalSBI) totalSBI from " & _
					"  ( select  a.accountcode code,a.activity_heading jobheading ,a.Description job_name ,sum(a.tm_amt_actual) totalBI,0 totalSBI from sac_accountreport_tm_tbm a " & _
					"    where a.period_year='" & request.Form("vYear") & "' and  a.period_month='" & request.Form("vMonth") & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"'  and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 & strsql6 & "  " & _				
					"    GROUP BY  a.activity_heading,a.accountcode,a.Description " & _
					"    UNION " & _
					"    select  a.accountcode code,a.activity_heading jobheading ,a.Description job_name ,0 totalBI,sum(a.tm_amt_actual) totalSBI from sac_accountreport_tm_tbm a " & _
					"    where a.period_year='" & request.Form("vYear") & "' and  a.period_month <='" & request.Form("vMonth") & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& request("vCROPTYPE") &"'  and a.ba='" & session("fcba") & "' " & strsql2 & strsql3  & strsql6 & "  " & _				
					"    GROUP BY  a.activity_heading,a.accountcode,a.Description " & _
					"  ) group by code,jobheading,job_name	"		
					
				   
		
          	 
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
			%>	   
				   <td align="right"><%=noline%></td> 
				   <td><%=rsha("code")%></td>
				   <td><%=rsha("jobheading")%></td>
				   <td><%=rsha("job_name")%></td>
				   <td align="right"><%=formatnumber(rsha("totalBI"),0)%></td>		
				   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
							<td align="right"><%= formatnumber( cdbl(rsha("totalBI")) / cdbl(prodkg ),0) %></td>
					<% else  %> 
							<td align="right"><%= formatnumber( cdbl(rsha("totalBI")) / cint(luasHA),0) %></td>
					<% end if %>	
				   <td>&nbsp;</td>
				   <td>&nbsp;</td>
				   <td>&nbsp;</td>
				   <td align="right"><%=formatnumber(rsha("totalSBI"),0)%></td>
                    <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
						      <td align="right"><%= formatnumber( cdbl(rsha("totalSBI")) / cdbl(prodkg ),0) %></td>
					<% else  %> 
					   		  <td align="right"><%= formatnumber( cdbl(rsha("totalSBI")) / cint(luasHA),0) %></td>
					<% end if %>				   
					
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				<%   
			   ELSE %>
			      
					   <%  if  cdgrp <> "" then  %>
							 <tr>
								<% GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(subAmtgrpBI) 
								   GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(subAmtgrpSBI)
								%>
								
								<%  
								   If cdgrp = "PRM" Then
								      grpDesc = "Permulaan (PRM) "
								   elseif 	 cdgrp = "PML" Then 
								      grpDesc = "Pemeliharaan (PML) "
								   elseif 	 cdgrp = "GC" Then 
								      grpDesc = "General Charge (GC) "	
  								   elseif	cdgrp = "PNN" Then 								   
								      grpDesc = "Panen (PNN) "	
                                   end if									  
								%>
								
								<td colspan="4"><b> Total&nbsp;<%=grpDesc%> </td>	
								<td align="right"><b><%=formatnumber(subAmtgrpBI,0)%></td>
								 <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td align="right"><b><%= formatnumber( cdbl(subAmtgrpBI) / cdbl(prodkg ),0) %></td>
								<% else  %> 
									  <td align="right"><b><%= formatnumber( cdbl(subAmtgrpBI) / cint(luasHA),0) %></td>
								<% end if %>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td align="right"><b><%=formatnumber(subAmtgrpSBI,0)%></td>
								 <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
									  <td align="right"><b><%= formatnumber( cdbl(subAmtgrpSBI) / cdbl(prodkg ),0) %></td>
								<% else  %> 
									  <td align="right"><b><%= formatnumber( cdbl(subAmtgrpSBI) / cint(luasHA),0) %></td>
								<% end if %>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							 </tr>
														
						<%  end if %>
				
						<%   noline = 1
							 subAmtgrpBI = rsha("totalBI")
							 subAmtgrpSBI = rsha("totalSBI")
						%>
				   
					   <%  if  cdgrp = "PML" and request("vStatus") = "TM" then  %>
							<tr>
								<td colspan="4">&nbsp;</td>
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
						   <td align="right"><%=noline%>&nbsp;</td> 
						   <td><%=rsha("code")%></td>
						   <td><%=rsha("jobheading")%></td>
						   <td><%=rsha("job_name")%></td>
						   <td align="right"><%=formatnumber(rsha("totalBI"),0)%></td> 
						   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" or  cdgrp = "PML") and request("vStatus") = "TM" then  %>
								  <td align="right"><%= formatnumber( cdbl(rsha("totalBI")) / cdbl(prodkg ),0) %></td> 
						   <% else  %> 
								  <td align="right"><%= formatnumber( cdbl(rsha("totalBI")) / cint(luasHA),0) %></td>
						   <% end if %>	
						   <td>&nbsp;</td>
						   <td>&nbsp;</td>
						   <td>&nbsp;</td>
						   <td align="right"><%=formatnumber(rsha("totalSBI"),0)%></td>
							<%  if  ( cdgrp = "PNN" or cdgrp = "GC" or  cdgrp = "PML" ) and request("vStatus") = "TM" then  %>
								  <td align="right" ><%= formatnumber( cdbl(rsha("totalSBI")) / cdbl(prodkg ),0) %></td>
							<% else  %> 
								  <td align="right"><%= formatnumber( cdbl(rsha("totalSBI")) / cint(luasHA),0) %></td>
							<% end if %>				   
						
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
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
		<td colspan="4"> <b>Total General Charge (GC)</td>	
		<td align="right"><b><%=formatnumber(subAmtgrpBI,0)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
				<td align="right"><b><%= formatnumber( cdbl(subAmtgrpBI) / cdbl(prodkg ),0) %></td>
		<% else  %> 
			   <td align="right"><b><%= formatnumber( cdbl(subAmtgrpBI) / cint(luasHA),0) %></td>
		<% end if %>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td align="right"><b><%=formatnumber(subAmtgrpSBI,0)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
				<td align="right"><b><%= formatnumber( cdbl(subAmtgrpSBI) / cdbl(prodkg ),0) %></td>
		<% else  %> 
			   <td align="right"><b><%= formatnumber( cdbl(subAmtgrpSBI) / cint(luasHA),0) %></td>
		<% end if %>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>				
	<tr>	
	<tr>
	     <% GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(subAmtgrpBI) 
		    GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(subAmtgrpSBI) 
		 %>
		<td colspan="4"><b> Grand Total&nbsp;</td>	
		<td align="right"><b><%=formatnumber(GRAmtgrpBI,0)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
				<td align="right"><b><%= formatnumber( cdbl(GRAmtgrpBI) / cdbl(prodkg ),0) %></b></td>
		<% else  %> 
			   <td align="right"><b><%= formatnumber( cdbl(GRAmtgrpBI) / cint(luasHA),0) %></b></td>
		<% end if %>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td align="right"><b><%=formatnumber(GRAmtgrpSBI,0)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and request("vStatus") = "TM" then  %>
				<td align="right"><b><%= formatnumber( cdbl(GRAmtgrpSBI) / cdbl(prodkg ),0) %></b></td>
		<% else  %> 
			   <td align="right"><b><%= formatnumber( cdbl(GRAmtgrpSBI) / cint(luasHA),0) %></b></td>
		<% end if %>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>				
	<tr>	
</table>
<!--#include file="bottom.Asp"-->