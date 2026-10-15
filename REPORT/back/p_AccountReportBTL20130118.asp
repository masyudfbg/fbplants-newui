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
<title>iPla'S Reporting - Analisa Biaya Tak Langsung</title>
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
	<!--#include file="FromSelectHDRBTLMILL.html"-->
	
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
		<h2>LAPORAN BIAYA TAK LANGSUNG</h3>
		<% IF request.Form("vMonth") <> "" then %>
		<h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=request.Form("vYear")%></h5>
		<% END IF %>
		
	</div>
	</div>
<%
Dim strtbs, rsha,TBSprocess_BI,TBSprocess_SBI,ProdCPO_BI,ProdCPO_SBI,ProdKER_BI,ProdKER_SBI,Palm_BI,Palm_SBI, strprod
TBSprocess_BI=0
TBSprocess_SBI=0

strtbs = "  select fcba, sum(TBS_PROCESS_BI) TBS_PROCESS_BI, sum(TBS_PROCESS_SBI) TBS_PROCESS_SBI  from ( " & _
		 "  select t.fcba, sum(t.cageweight) TBS_PROCESS_BI,0 TBS_PROCESS_SBI from SHIFTREPORT t " & _  
	     "	where to_char(t.productiondate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.productiondate,'MM') = '" & request.Form("vMonth") & "' and t.fcba ='" & session("fcba")  & "'" & _
		 "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  " & _
		 "  union " & _
		 "  select t.fcba, 0 TBS_PROCESS_BI, sum(t.cageweight) TBS_PROCESS_SBI from SHIFTREPORT t  " & _
		 "	where to_char(t.productiondate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.productiondate,'MM') <= '" & request.Form("vMonth") & "' and t.fcba ='" & session("fcba")  & "'" & _
         "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  ) group by fcba " 
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic
'response.write strtbs
if Not rsha.eof then
    TBSprocess_BI = rsha("TBS_PROCESS_BI")
	TBSprocess_SBI = rsha("TBS_PROCESS_SBI")
end if
rsha.close()

ProdCPO_BI =0
ProdCPO_SBI =0

strtbs = " select fcba, sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 " select t.fcba,  sum(t.quantity) Produksi_CPO_BI,0 Produksi_CPO_SBI from PRODUCTSTORAGEREADING t  " & _
		 " where t.pstorage_storagecode = 'CIP1' and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = '" & request.Form("vMonth") & "'  and t.fcba ='" & session("fcba")  & "' " & _
		 " group by t.fcba, to_char(t.readingdate,'YYYYMM') " & _
		 " union " & _
		 " select t.fcba, 0 Produksi_CPO_BI,sum(t.quantity) Produksi_CPO_SBI from PRODUCTSTORAGEREADING t " & _
		 " where t.pstorage_storagecode = 'CIP1' and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') <= '" & request.Form("vMonth") & "'  and t.fcba ='" & session("fcba")  & "' " & _
		 " group by t.fcba, to_char(t.readingdate,'YYYYMM') ) group by fcba " 
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    ProdCPO_BI = rsha("Produksi_CPO_BI")
	ProdCPO_SBI = rsha("Produksi_CPO_SBI")
end if
rsha.close()

ProdKER_BI =0
ProdKER_SBI =0

strtbs = " select fcba, sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI from ( " & _
	     " select t.fcba, sum(t.quantity) Produksi_KER_BI, 0 Produksi_KER_SBI from PRODUCTSTORAGEREADING t  " & _
		 " where t.pstorage_storagecode = 'KIP1'  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = '" & request.Form("vMonth") & "'  and t.fcba ='" & session("fcba")  & "' " & _
		 " group by t.fcba,to_char(t.readingdate,'YYYYMM')  " & _
		 " union " & _
		 " select t.fcba, 0 Produksi_KER_BI, sum(t.quantity) Produksi_KER_SBI from PRODUCTSTORAGEREADING t " & _
		 " where t.pstorage_storagecode = 'KIP1'  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') <= '" & request.Form("vMonth") & "'  and t.fcba ='" & session("fcba")  & "' " & _
		 " group by t.fcba,to_char(t.readingdate,'YYYYMM') ) group by fcba " 
	  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    ProdKER_BI = rsha("Produksi_KER_BI")
	ProdKER_SBI = rsha("Produksi_KER_SBI")
end if
rsha.close()
Palm_BI = cdbl(ProdKER_BI) +  cdbl(ProdCPO_BI)
Palm_SBI = cdbl(ProdKER_SBI) +  cdbl(ProdCPO_SBI)
%>	

<table border="0" width="100%" id="product-table">
	<tr>
		<th  style="color:#FFFFFF" colspan="2"></th>
		
		<th  style="color:#FFFFFF" colspan="7">BULAN INI</th>
		<th  style="color:#FFFFFF" colspan="7">S/D BULAN INI</th>
		<th  style="color:#FFFFFF" colspan="3" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th  style="color:#FFFFFF" colspan="2"></th>
		<th  style="color:#FFFFFF" colspan="3">REALISASI</th>
		
		<th  style="color:#FFFFFF" colspan="3">ANGGARAN</th>
		<th style="color:#FFFFFF" >%</th>
		<th  style="color:#FFFFFF" colspan="3">REALISASI</th>
		<th  style="color:#FFFFFF" colspan="3">ANGGARAN</th>
		<th style="color:#FFFFFF" >%</th>
	</tr>
	<tr>
	   
		<th  style="color:#FFFFFF"  colspan="2" align="right">TBS PROSES (KG)</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(TBSprocess_BI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
		<th style="color:#FFFFFF" align="right" >0</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(TBSprocess_SBI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
		<th style="color:#FFFFFF" align="right" >0</th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
		
		
	</tr>
	<tr>
	    
		<th  style="color:#FFFFFF" colspan="2" align="right">PRODUKSI CPO (KG)</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(ProdCPO_BI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
		<th style="color:#FFFFFF" align="right" >0</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(ProdCPO_SBI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
		<th style="color:#FFFFFF" align="right" >0</th>
		<th style="color:#FFFFFF" align="right" colspan="3">0</th>
	</tr>
	<tr>
	    
		<th  style="color:#FFFFFF" colspan="2" align="right">PRODUKSI KER (KG)</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(ProdKER_BI,2)%></th>
		<th style="color:#FFFFFF" align="right" colspan="3">0</th>
		<th style="color:#FFFFFF" align="right" >0</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(ProdKER_SBI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
		<th style="color:#FFFFFF" align="right" >0</th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
	</tr>
	<tr>
	    
		<th  style="color:#FFFFFF" colspan="2" align="right">PRODUK (KG)</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(Palm_BI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
		<th style="color:#FFFFFF" align="right" >0</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(Palm_SBI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
		<th style="color:#FFFFFF" align="right" >0</th>
		<th  style="color:#FFFFFF" align="right" colspan="3">0</th>
	</tr>
	<tr   >
	    <% if   request.Form("vJenisReport") = "Job"  then %>
		<th style="color:#FFFFFF"rowspan="2" >Kode</th>
		<th style="color:#FFFFFF"rowspan="2" >Nama Pekerjaan</th>
		<% else %>
		<th style="color:#FFFFFF"rowspan="2" >Kode</th>
		<th style="color:#FFFFFF"rowspan="2" >Nama Stasiun</th>
		<% end if %>
		<th style="color:#FFFFFF" rowspan="2">Total</th>
		<th style="color:#FFFFFF" colspan="2">Rp./Kg.</th>
		
		<th style="color:#FFFFFF" rowspan="2">Total</th>
		<th style="color:#FFFFFF" colspan="2">Rp./Kg.</th>
		<th style="color:#FFFFFF" rowspan="2" align="right" ></th>
		<th style="color:#FFFFFF" rowspan="2">Total</th>
		<th style="color:#FFFFFF" colspan="2">Rp./Kg.</th>
		<th style="color:#FFFFFF" rowspan="2">Total</th>
		<th style="color:#FFFFFF" colspan="2">Rp./Kg.</th>
		<th style="color:#FFFFFF" rowspan="2" align="right" ></th>
		<th style="color:#FFFFFF" rowspan="2">Total</th>
		<th style="color:#FFFFFF" colspan="2">Rp./Kg.</th>
	</tr>
	<tr>
		
		<th style="color:#FFFFFF" >Produk</th>
		<th style="color:#FFFFFF" >TBS</th>
		
		<th style="color:#FFFFFF" >Produk</th>
		<th style="color:#FFFFFF" >TBS</th>
		
		<th style="color:#FFFFFF" >Produk</th>
		<th style="color:#FFFFFF" >TBS</th>
		
		<th style="color:#FFFFFF" >Produk</th>
		<th style="color:#FFFFFF" >TBS</th>
		
		<th style="color:#FFFFFF" >Produk</th>
		<th style="color:#FFFFFF" >TBS</th>
	</tr>
	
	<%
	Dim strdtl , strgrp ,cdgrp,cdgrpdesc  ,noline ,  subAmtgrpBI , GRAmtgrpBI, subAmtgrpSBI , GRAmtgrpSBI , grpDesc
	
			strdtl = " select fcba,jobcode,description, sum(BI) BI, sum(SBI) SBI from ( " & _
					 " select to_char(t.tdate,'YYYY') year, to_char(t.tdate,'MM') month,t.fcba, t.jobcode,get_jobcode_name(t.jobcode) Description, sum(t.amount) BI, 0 SBI " & _
					 " from cost_activity t where substr(t.jobcode,1,2)='64' and to_char(t.tdate,'YYYY')= '" & request.Form("vYear") & "' and to_char(t.tdate,'MM')= '" & request.Form("vMonth") & "' and t.fcba ='" & session("fcba") & "' " & _
					 " group by to_char(t.tdate,'YYYY'), to_char(t.tdate,'MM'),t.fcba, t.jobcode " & _
					 " union " & _
					 " select to_char(t.tdate,'YYYY') year, to_char(t.tdate,'MM') month,t.fcba, t.jobcode,get_jobcode_name(t.jobcode) Description,0 BI, sum(t.amount) SBI " & _
					 " from cost_activity t where substr(t.jobcode,1,2)='64' and to_char(t.tdate,'YYYY')= '" & request.Form("vYear") & "' and to_char(t.tdate,'MM') <= '" & request.Form("vMonth") & "' and t.fcba ='" & session("fcba") & "' " & _
					 " group by to_char(t.tdate,'YYYY'), to_char(t.tdate,'MM'),t.fcba, t.jobcode ) " & _
					 " group by fcba,jobcode,description order by fcba,jobcode "
	
		
		
			 strdtl =  strdtl  + strsql1  + strgrp 
	
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    'response.write strdtl
	  if cdbl(Palm_BI) = 0 then
		      Palm_BI = "1"
	  end if
	  if cdbl(Palm_SBI) = 0 then
		      Palm_SBI = "1"
	  end if
	  if cdbl(TBSprocess_BI) = 0 then
		      TBSprocess_BI = "1"
	  end if
	  if cdbl(TBSprocess_SBI) = 0 then
		      TBSprocess_SBI = "1"
	  end if
	GRAmtgrpBI ="0"
	GRAmtgrpSBI ="0"
	if Not rsha.eof then
		  while Not rsha.eof
		    
			 IF rsha("JobCode") <> "6449800"	 THEN  %>
				<tr>  
					<td><%=rsha("JobCode")%></td>
					<td><%=rsha("DESCRIPTION")%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("BI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(palm_BI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(TBSprocess_BI)),2)%></td>
					<td align="right">0</td>
					<td align="right">0</td>
				    <td align="right">0</td>
				    <td align="right">0%</td>
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(palm_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)),2)%></td>
					<td align="right">0</td>
					<td align="right">0</td>
					<td align="right">0</td>
					<td align="right">0%</td>
					<td align="right">0</td>
					<td align="right">0</td>
					<td align="right">0</td>
				<tr>	
            <% ELSE  %>	
               <tr>  							
					<td  style="background-color:#F4F4A0"  colspan="2"><b>Total Biaya Tak Langsung</td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpBI)/cdbl(palm_BI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)),2)%></td>	
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
				    <td  style="background-color:#F4F4A0" align="right">0</td>
				    <td  style="background-color:#F4F4A0" align="right">0%</td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0%</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>														
				<tr>  
                 <% 
				    GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
			        GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI"))
 				 %>
				<tr>  
					<td><%=rsha("JobCode")%></td>
					<td><%=rsha("DESCRIPTION")%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("BI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(palm_BI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(TBSprocess_BI)),2)%></td>
					<td align="right">0</td>
					<td align="right">0</td>
				    <td align="right">0</td>
				    <td align="right">0%</td>
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(palm_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)),2)%></td>
					<td align="right">0</td>
					<td align="right">0</td>
					<td align="right">0</td>
					<td align="right">0%</td>
					<td align="right">0</td>
					<td align="right">0</td>
					<td align="right">0</td>
				</tr>
              					
                <tr>  							
					<td  style="background-color:#F4F4A0" colspan="2"><b>Total Keseluruhan</td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpBI)/cdbl(palm_BI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)),2)%></td>	
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
				    <td  style="background-color:#F4F4A0" align="right">0</td>
				    <td  style="background-color:#F4F4A0" align="right">0%</td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0%</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>
					<td  style="background-color:#F4F4A0" align="right">0</td>														
				<tr>  
          <% END IF %>				
    <%        
	          GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
			  GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI"))  
		      rsha.MoveNext
		  wend
	end if
	%>				
</table>
<!--#include file="bottom.Asp"-->