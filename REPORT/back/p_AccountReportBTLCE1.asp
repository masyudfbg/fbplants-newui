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
Dim cur_month,strbud_BI,strbud_SBI,strbudTBS_BI,strbudTBS_SBI,strbudCPO_BI,strbudCPO_SBI,strbudKER_BI,strbudKER_SBI,x1
	    strbud_BI = "btm.budget_01" 
	    strbud_SBI = "btm.budget_01"
		strbudTBS_BI = "bmp.ffb_01"
		strbudTBS_SBI = "bmp.ffb_01"
		strbudCPO_BI = "bmp.cpo_01"
		strbudCPO_SBI = "bmp.cpo_01"
		strbudKER_BI = "bmp.ker_01"
		strbudKER_SBI = "bmp.ker_01"
		
		cur_month = request.Form("vMonth") 
		
		IF Len(cur_month) > 0 then

			strbud_BI = "" 
	        strbud_SBI = ""
			strbudTBS_BI = ""
		    strbudTBS_SBI = ""
		    strbudCPO_BI = ""
		    strbudCPO_SBI = ""
		    strbudKER_BI = ""
		    strbudKER_SBI = ""
			x1=1
			
			Do While x1<= cint(cur_month) 'Output the values from 0 to 10
			  'response.write(i & "<br \>")
			  'response.write  " Budget_" & Lpad(x1,"0",2) & " +"
			  strbud_SBI = strbud_SBI & " btm.budget_" & Lpad(x1,"0",2) & " +"
			  strbudTBS_SBI = strbudTBS_SBI & "bmp.ffb_" & Lpad(x1,"0",2) & " +"
			  strbudCPO_SBI = strbudCPO_SBI & "bmp.cpo_" & Lpad(x1,"0",2) & " +"
			  strbudKER_SBI = strbudKER_SBI & "bmp.ker_" & Lpad(x1,"0",2) & " +"
			  x1 = x1 + 1 'increment the value of i for next time loop executes
			Loop

			strbud_BI = "btm.budget_" & Lpad(cur_month,"0",2)
			strbudTBS_BI = "bmp.ffb_" & Lpad(cur_month,"0",2)
			strbudCPO_BI = "bmp.cpo_" & Lpad(cur_month,"0",2)
			strbudKER_BI = "bmp.ker_" & Lpad(cur_month,"0",2)
			
			strbud_SBI = Mid(strbud_SBI,1,Len(strbud_SBI)-1)
			strbudTBS_SBI = Mid(strbudTBS_SBI,1,Len(strbudTBS_SBI)-1)
			strbudCPO_SBI = Mid(strbudCPO_SBI,1,Len(strbudCPO_SBI)-1)
			strbudKER_SBI = Mid(strbudKER_SBI,1,Len(strbudKER_SBI)-1)
			
			
            ' response.write  strbudTBS_BI & "<br>"
			' response.write strbudCPO_BI & "<br>"
			' response.write strbudKER_BI & "<br>"
			
			
			' response.write strbudTBS_SBI  & "<br>"
			' response.write strbudCPO_SBI  & "<br>"
			' response.write strbudKER_SBI  & "<br>"
		END IF	



Dim strtbs, rsha,TBSprocess_BI,TBSprocess_SBI,ProdCPO_BI,ProdCPO_SBI,ProdKER_BI,ProdKER_SBI,Palm_BI,Palm_SBI, strprod
TBSprocess_BI=0
TBSprocess_SBI=0

strtbs = "  select fcba, sum(TBS_PROCESS_BI) TBS_PROCESS_BI, sum(TBS_PROCESS_SBI) TBS_PROCESS_SBI  from ( " & _
		 "  select t.fcba, sum(t.cageweight) TBS_PROCESS_BI,0 TBS_PROCESS_SBI from SHIFTREPORT t " & _  
	     "	where to_char(t.productiondate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.productiondate,'MM') = lpad('" & request.Form("vMonth") & "',2,0) and t.fcba ='" & session("fcba")  & "'" & _
		 "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  " & _
		 "  union " & _
		 "  select t.fcba, 0 TBS_PROCESS_BI, sum(t.cageweight) TBS_PROCESS_SBI from SHIFTREPORT t  " & _
		 "	where to_char(t.productiondate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.productiondate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0) and t.fcba ='" & session("fcba")  & "'" & _
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

' strtbs = " select fcba, sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 ' " select t.fcba,  sum(t.quantity) Produksi_CPO_BI,0 Produksi_CPO_SBI from PRODUCTSTORAGEREADING t  " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='CPO')  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba, to_char(t.readingdate,'YYYYMM') " & _
		 ' " union " & _
		 ' " select t.fcba, 0 Produksi_CPO_BI,sum(t.quantity) Produksi_CPO_SBI from PRODUCTSTORAGEREADING t " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='CPO')  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba, to_char(t.readingdate,'YYYYMM') ) group by fcba " 
		 
strtbs = " select fcba, sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_CPO_BI, 0 Produksi_CPO_SBI from MILL_PRODUCTION t  " & _
		 " where t.product_code='CPO'  and to_char(t.tdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.tdate,'MM') = lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union " & _
		 " select t.fcba,0 Produksi_CPO_BI, sum(t.production_today) Produksi_CPO_SBI from MILL_PRODUCTION t " & _
		 " where t.product_code='CPO'  and to_char(t.tdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.tdate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba " 		 
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    ProdCPO_BI = rsha("Produksi_CPO_BI")
	ProdCPO_SBI = rsha("Produksi_CPO_SBI")
end if
rsha.close()

ProdKER_BI =0
ProdKER_SBI =0

' strtbs = " select fcba, sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI from ( " & _
	     ' " select t.fcba, sum(t.quantity) Produksi_KER_BI, 0 Produksi_KER_SBI from PRODUCTSTORAGEREADING t  " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='Kernel')   and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba,to_char(t.readingdate,'YYYYMM')  " & _
		 ' " union " & _
		 ' " select t.fcba, 0 Produksi_KER_BI, sum(t.quantity) Produksi_KER_SBI from PRODUCTSTORAGEREADING t " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='Kernel')  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba,to_char(t.readingdate,'YYYYMM') ) group by fcba " 
		 
strtbs = " select fcba, sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_KER_BI, 0 Produksi_KER_SBI from MILL_PRODUCTION t  " & _
		 " where t.product_code='KER'  and to_char(t.tdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.tdate,'MM') = lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union " & _
		 " select t.fcba,0 Produksi_KER_BI, sum(t.production_today) Produksi_KER_SBI from MILL_PRODUCTION t " & _
		 " where t.product_code='KER'  and to_char(t.tdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.tdate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba " 		
	  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    ProdKER_BI = rsha("Produksi_KER_BI")
	ProdKER_SBI = rsha("Produksi_KER_SBI")
end if
rsha.close()
Palm_BI = cdbl(ProdKER_BI) +  cdbl(ProdCPO_BI)
Palm_SBI = cdbl(ProdKER_SBI) +  cdbl(ProdCPO_SBI)

' ambil data budget CPO, TBS dan ker

Dim TBSbud_BI,TBSbud_SBI,TBSbud_THN,CPObud_BI,CPObud_SBI,CPObud_THN,KERbud_BI,KERbud_SBI,KERbud_THN,PALMbud_BI,PALMbud_SBI,PALMbud_THN
strtbs = "  select fcba , sum(TBSbudBI) TBSbudBI,sum(tbsbudSBI) tbsbudSBI,sum(tbsbudthn) tbsbudthn,sum(cpobudBI) cpobudBI,sum(cpobudSBI) cpobudSBI,sum(cpobudthn) cpobudthn,sum(kerbudBI) kerbudBI,sum(kerbudSBI) kerbudSBI, sum(kerbudthn) kerbudthn  from ( " & _
		 "	 select bmp.fcba , sum(" & strbudTBS_BI & ") TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn  from budget_mill_prod bmp " & _
		 "	 where bmp.period_year='" & request.Form("vYear") & "' and bmp.fcba='" & session("fcba") & "' " & _
		 "	 group by bmp.fcba " & _
		 "	 union " & _
		 "	 select bmp.fcba , 0 TBSbudBI,sum(" & strbudTBS_SBI & ") tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn  from budget_mill_prod bmp " & _
		 "	 where bmp.period_year='" & request.Form("vYear") & "' and bmp.fcba='" & session("fcba") & "' " & _
		 "	 group by bmp.fcba " & _
		 "	 union " & _
		 "	 select bmp.fcba , 0 TBSbudBI,0 tbsbudSBI,sum(bmp.ffb_01+bmp.ffb_02+bmp.ffb_03+bmp.ffb_04+bmp.ffb_05+bmp.ffb_06+bmp.ffb_07+bmp.ffb_08+bmp.ffb_09+bmp.ffb_10+bmp.ffb_11+bmp.ffb_12) tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from budget_mill_prod bmp " & _
		 "	 where bmp.period_year='" & request.Form("vYear") & "' and bmp.fcba='" & session("fcba") & "' " & _
		 "	 group by bmp.fcba " & _
		 "	 union " & _
		 "	 select bmp.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,sum(" & strbudCPO_BI & ") cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from budget_mill_prod bmp " & _
		 "	 where bmp.period_year='" & request.Form("vYear") & "' and bmp.fcba='" & session("fcba") & "' " & _
		 "	 group by bmp.fcba " & _
		 "	 union " & _
		 "	 select bmp.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,sum(" & strbudCPO_SBI & ") cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from budget_mill_prod bmp " & _
		 "	 where bmp.period_year='" & request.Form("vYear") & "' and bmp.fcba='" & session("fcba") & "' " & _
		 "	 group by bmp.fcba " & _
		 "	 union " & _
		 "	 select bmp.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,sum(bmp.cpo_01+bmp.cpo_02+bmp.cpo_03+bmp.cpo_04+bmp.cpo_05+bmp.cpo_06+bmp.cpo_07+bmp.cpo_08+bmp.cpo_09+bmp.cpo_10+bmp.cpo_11+bmp.cpo_12) cpobudthn,0 kerbudBI, 0 kerbudSBI , 0 kerbudthn  from budget_mill_prod bmp " & _
		 "	 where bmp.period_year='" & request.Form("vYear") & "' and bmp.fcba='" & session("fcba") & "' " & _
		 "	 group by bmp.fcba " & _
		 "	  union " & _
		 "	 select bmp.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudthn,0 cpobudSBI,sum(" & strbudKER_BI & ") kerbudBI, 0 kerbudSBI, 0 kerbudthn   from budget_mill_prod bmp " & _
		 "	 where bmp.period_year='" & request.Form("vYear") & "' and bmp.fcba='" & session("fcba") & "' " & _
		 "	 group by bmp.fcba " & _
		 "	  union " & _
		 "	 select bmp.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, sum(" & strbudKER_SBI & ") kerbudSBI, 0 kerbudthn   from budget_mill_prod bmp " & _
		 "	 where bmp.period_year='" & request.Form("vYear") & "' and bmp.fcba='" & session("fcba") & "' " & _
		 "	 group by bmp.fcba " & _
		 "	 union " & _
		 "	 select bmp.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, sum(bmp.ker_01+bmp.ker_02+bmp.ker_03+bmp.ker_04+bmp.ker_05+bmp.ker_06+bmp.ker_07+bmp.ker_08+bmp.ker_09+bmp.ker_10+bmp.ker_11+bmp.ker_12) kerbudthn   from budget_mill_prod bmp " & _
		 "	 where bmp.period_year='" & request.Form("vYear") & "'and bmp.fcba='" & session("fcba") & "' " & _
		 "	 group by bmp.fcba ) group by fcba " 
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic
'response.write strtbs
if Not rsha.eof then
    TBSbud_BI = rsha("TBSbudBI")
	TBSbud_SBI = rsha("TBSbudSBI")
	TBSbud_THN = rsha("TBSbudTHN")
	CPObud_BI = rsha("cpobudBI")
	CPObud_SBI = rsha("cpobudSBI")
	CPObud_THN = rsha("cpobudTHN")
	KERbud_BI = rsha("kerbudBI")
	KERbud_SBI = rsha("kerbudSBI")
	KERbud_THN = rsha("kerbudTHN")
	PALMbud_BI =   cdbl(KERbud_BI) +  cdbl(CPObud_BI)
	PALMbud_SBI =  cdbl(KERbud_SBI) +  cdbl(CPObud_SBI)
	PALMbud_THN =  cdbl(KERbud_THN) +  cdbl(CPObud_THN)
end if
rsha.close()
if cdbl(TBSbud_BI)=0 then
     TBSbud_BI=1
end if
if cdbl(TBSbud_SBI)=0 then
     TBSbud_SBI=1
end if
if cdbl(TBSbud_THN)=0 then
     TBSbud_THN=1
end if
if cdbl(CPObud_BI) = 0 then
   CPObud_BI=1
end if   
if cdbl(CPObud_SBI)=0 then
     CPObud_SBI=1
end if
if cdbl(CPObud_THN)=0 then
     CPObud_THN=1
end if
if cdbl(KERbud_BI)=0 then
     KERbud_BI=1
end if
if cdbl(KERbud_SBI)=0 then
     KERbud_SBI=1
end if
if cdbl(KERbud_THN)=0 then
     KERbud_THN=1
end if
if cdbl(PALMbud_BI)=0 then
     PALMbud_BI=1
end if
if cdbl(PALMbud_SBI)=0 then
     PALMbud_SBI=1
end if
if cdbl(PALMbud_THN)=0 then
     PALMbud_THN=1
end if

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
	    <% 'TBSbud_BI,TBSbud_SBI,TBSbud_THN,CPObud_BI,CPObud_SBI,CPObud_THN,KERbud_BI,KERbud_SBI,KERbud_THN,PALMbud_BI,PALMbud_SBI,PALMbud_THN %>
		<th  style="color:#FFFFFF"  colspan="2" align="right">TBS PROSES (KG)</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(TBSprocess_BI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(TBSbud_BI,2)%></th>
		<th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(TBSprocess_BI)-cdbl(TBSbud_BI))/(cdbl(TBSbud_BI)),2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(TBSprocess_SBI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(TBSbud_SBI,2)%></th>
		<th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(TBSprocess_SBI)-cdbl(TBSbud_SBI))/(cdbl(TBSbud_SBI)),2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(TBSbud_THN,2)%></th>
		
		
	</tr>
	<tr>
	    
		<th  style="color:#FFFFFF" colspan="2" align="right">PRODUKSI CPO (KG)</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(ProdCPO_BI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(CPObud_BI,2)%></th>
		<th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(ProdCPO_BI)-cdbl(CPObud_BI))/(cdbl(CPObud_BI)),2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(ProdCPO_SBI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(CPObud_SBI,2)%></th>
		<th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(ProdCPO_SBI)-cdbl(CPObud_SBI))/(cdbl(CPObud_SBI)),2)%></th>
		<th style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(CPObud_THN,2)%></th>
	</tr>
	<tr>
	    
		<th  style="color:#FFFFFF" colspan="2" align="right">PRODUKSI KER (KG)</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(ProdKER_BI,2)%></th>
		<th style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(KERbud_BI,2)%></th>
		<th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(ProdKER_BI)-cdbl(KERbud_BI))/(cdbl(KERbud_BI)),2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(ProdKER_SBI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(KERbud_SBI,2)%></th>
		<th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(ProdKER_SBI)-cdbl(KERbud_SBI))/(cdbl(KERbud_SBI)),2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(KERbud_THN,2)%></th>
	</tr>
	<tr>
	    
		<th  style="color:#FFFFFF" colspan="2" align="right">PRODUK (KG)</th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(Palm_BI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(PALMbud_BI,2)%></th>
		<th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(Palm_BI)-cdbl(Palmbud_BI))/(cdbl(Palmbud_BI)),2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(Palm_SBI,2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(PALMbud_SBI,2)%></th>
		<th style="color:#FFFFFF" align="right" ><%=formatnumber((cdbl(Palm_SBI)-cdbl(Palmbud_SBI))/(cdbl(Palmbud_SBI)),2)%></th>
		<th  style="color:#FFFFFF" align="right" colspan="3"><%=formatnumber(PALMbud_THN,2)%></th>
	</tr>
	<tr   >
	  
	    <th style="color:#FFFFFF"rowspan="2" >Kode</th>
		<%  if len(request.Form("vJenisReport"))<1 then %>
		      <th style="color:#FFFFFF"rowspan="2" >Nama Stasiun</th>
		<% end if %>
		<% if   request.Form("vJenisReport") = "Job"  then %>
			<th style="color:#FFFFFF"rowspan="2" >Nama Pekerjaan</th>
		<% elseif request.Form("vJenisReport") = "Group"  then %>
			<th style="color:#FFFFFF"rowspan="2" >Nama Stasiun</th>
		<% elseif request.Form("vJenisReport") = "Mechine"  then %>
			<th style="color:#FFFFFF"rowspan="2" >Nama Mesin</th>
		<% elseif request.Form("vJenisReport") = "Combine"  then %>	
			<th style="color:#FFFFFF"rowspan="2" >Nama Pekerjaan</th>
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
	Dim strdtl , strgrp ,cdgrp,cdgrpdesc  ,noline ,  subAmtgrpBI , GRAmtgrpBI, subAmtgrpSBI , GRAmtgrpSBI , grpDesc,GRbudBI,GRbudSBI,GRbudthn
	
			strdtl = " select jobcode,description, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from ( " & _
					 "	select t.jobcode,get_jobcode_name(t.jobcode) Description, sum(t.amount) BI, 0 SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from cost_activity t where substr(t.jobcode,1,2)='64' And t.jobcode <> '6449800' and to_char(t.tdate,'YYYY')=  '" & request.Form("vYear") & "' and to_char(t.tdate,'MM')=  lpad('" & request.Form("vMonth") & "',2,0) and t.fcba ='" & session("fcba") & "' " & _
					 "	group by t.fcba, t.jobcode " & _
					 "	union " & _
					 "	select  t.jobcode,get_jobcode_name(t.jobcode) Description,0 BI, sum(t.amount) SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from cost_activity t where substr(t.jobcode,1,2)='64' And t.jobcode <> '6449800' and to_char(t.tdate,'YYYY')=  '" & request.Form("vYear") & "' and to_char(t.tdate,'MM') <=   lpad('" & request.Form("vMonth") & "',2,0) and t.fcba ='" & session("fcba") & "' " & _
					 "	group by t.fcba, t.jobcode  " & _
					 "	union " & _
					 "	select  t.jobcode,get_jobcode_name(t.jobcode) Description, sum(t.amount)* -1 BI, 0 SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from cost_activity t where t.jobcode = '6610600'  and to_char(t.tdate,'YYYY')=  '" & request.Form("vYear") & "' and to_char(t.tdate,'MM')=  lpad('" & request.Form("vMonth") & "',2,0) and t.fcba ='" & session("fcba") & "' " & _
					 "	group by t.fcba, t.jobcode " & _
					 "	union " & _
					 "	select  t.jobcode,get_jobcode_name(t.jobcode) Description,0 BI, sum(t.amount)* -1  SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from cost_activity t where t.jobcode = '6610600'  and to_char(t.tdate,'YYYY')=  '" & request.Form("vYear") & "' and to_char(t.tdate,'MM') <=   lpad('" & request.Form("vMonth") & "',2,0) and t.fcba ='" & session("fcba") & "' " & _
					 "	group by t.fcba, t.jobcode " & _
					 "	union " & _
					 "	select btm.jobcode ,get_jobcode_name(btm.jobcode) DESCRIPTION,0 BI, 0 sbi ,sum(" & strbud_BI & ")* -1 budBI,0 budSBI,0 budthn  from BUDGET_TRANSACTION_MILL btm " & _
					 "	where btm.budget_year ='" & request.Form("vYear") & "' and btm.fcba ='" & session("fcba") & "' and btm.jobcode='6610600' " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.jobcode ,get_jobcode_name(btm.jobcode) DESCRIPTION,0 BI, 0 sbi ,0 budBI,sum(" & strbud_SBI & ")* -1 budSBI,0 budthn  from BUDGET_TRANSACTION_MILL btm " & _
					 "	where btm.budget_year ='" & request.Form("vYear") & "' and btm.fcba ='" & session("fcba") & "' and btm.jobcode='6610600' " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.jobcode ,get_jobcode_name(btm.jobcode) DESCRIPTION,0 BI, 0 sbi ,0 budBI,0 budSBI,sum(btm.budget_01+btm.budget_02+btm.budget_03+btm.budget_04+btm.budget_05+btm.budget_06+btm.budget_07+btm.budget_08+btm.budget_09+btm.budget_10+btm.budget_11+btm.budget_12)* -1 budthn  from BUDGET_TRANSACTION_MILL btm " & _
					 "	where btm.budget_year ='" & request.Form("vYear") & "' and btm.fcba ='" & session("fcba") & "' and btm.jobcode='6610600' " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.jobcode ,get_jobcode_name(btm.jobcode) DESCRIPTION,0 BI, 0 sbi ,sum(" & strbud_BI & ") budBI,0 budSBI,0 budthn  from BUDGET_TRANSACTION_INDIRCOST btm " & _
					 "	where btm.budget_year ='" & request.Form("vYear") & "' and btm.fcba ='" & session("fcba") & "'and  substr(btm.jobcode,1,2)='64' And btm.jobcode <> '6449800' " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.jobcode ,get_jobcode_name(btm.jobcode) DESCRIPTION,0 BI, 0 sbi ,0 budBI,sum(" & strbud_SBI & ") budSBI,0 budthn  from BUDGET_TRANSACTION_INDIRCOST btm " & _
					 "	where btm.budget_year ='" & request.Form("vYear") & "' and btm.fcba ='" & session("fcba") & "'and  substr(btm.jobcode,1,2)='64' And btm.jobcode <> '6449800' " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.jobcode ,get_jobcode_name(btm.jobcode) DESCRIPTION,0 BI, 0 sbi ,0 budBI,0 budSBI,sum(btm.budget_01+btm.budget_02+btm.budget_03+btm.budget_04+btm.budget_05+btm.budget_06+btm.budget_07+btm.budget_08+btm.budget_09+btm.budget_10+btm.budget_11+btm.budget_12) budthn  from BUDGET_TRANSACTION_INDIRCOST btm " & _
					 "	where btm.budget_year ='" & request.Form("vYear") & "' and btm.fcba ='" & session("fcba") & "'and  substr(btm.jobcode,1,2)='64' And btm.jobcode <> '6449800' " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	 ) " & _
					 "	group by jobcode,description order by jobcode "
	
		
		
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
		    
			 IF rsha("JobCode") <> "6610600"	 THEN  %>
				<tr>  
					<td align="center"><%=rsha("JobCode")%></td>
					<td><%=rsha("DESCRIPTION")%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("BI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(palm_BI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(TBSprocess_BI)),2)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("budBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)),2)%></td>
					<% if (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)) = 0 then %>
						<td align="right"><%=formatnumber(((cdbl(rsha("BI"))/cdbl(palm_BI)) - (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)) ),2)%></td>
					<% else %>
						<td align="right"><%=formatnumber(((cdbl(rsha("BI"))/cdbl(palm_BI)) - (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)) )/(cdbl(rsha("budBI"))/cdbl(PALMbud_BI)),2)%></td>
					<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(palm_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)),2)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("budSBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)),2)%></td>
					<% if (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)) = 0 then %>
						<td align="right"><%=formatnumber(((cdbl(rsha("SBI"))/cdbl(palm_SBI)) - (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)) ),2)%></td>
					<% else %>
						<td align="right"><%=formatnumber(((cdbl(rsha("SBI"))/cdbl(palm_SBI)) - (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)) )/(cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)),2)%></td>
					<% End if %>
					<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)),2)%></td>
				<tr>	
            <% ELSE  %>	
               <tr>  							
					<td  style="background-color:#F4F4A0"  colspan="2"><b>Total Biaya Tak Langsung</td>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpBI)/cdbl(palm_BI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)),2)%></td>	
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudBI,2)%></td>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(PALMbud_BI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(TBSbud_BI)),2)%></td>
					<% If ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) = 0 then %>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRAmtgrpBI)/cdbl(palm_BI)) - ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) ) ,2)%></td>
					<% Else %>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRAmtgrpBI)/cdbl(palm_BI)) - ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) ) / ( cdbl(GRbudBI)/cdbl(PALMbud_BI)),2)%></td>
					<% End If %>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudSBI,2)%></td>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(TBSbud_SBI)),2)%></td>
					<% IF ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) = 0 then %>
						<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)) - ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) ) ,2)%></td>
					<% Else %>
						<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)) - ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) ) / ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
					<% End If %>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudthn,2)%></td>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(PALMbud_THN)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(TBSbud_THN)),2)%></td>														
				<tr>  
                 <% 
					 GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
					 GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI")) 
					 GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))
					 GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
					 GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))
 				 %>
				<tr>  
					<td align="center"><%=rsha("JobCode")%></td>
					<td><%=rsha("DESCRIPTION")%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("BI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(palm_BI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(TBSprocess_BI)),2)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("budBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)),2)%></td>
					<% if (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)) = 0 then %>
						<td align="right"><%=formatnumber(((cdbl(rsha("BI"))/cdbl(palm_BI)) - (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)) ),2)%></td>
					<% else %>
						<td align="right"><%=formatnumber(((cdbl(rsha("BI"))/cdbl(palm_BI)) - (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)) )/(cdbl(rsha("budBI"))/cdbl(PALMbud_BI)),2)%></td>
					<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(palm_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)),2)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("budSBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)),2)%></td>
					<% if (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)) = 0 then %>
						<td align="right"><%=formatnumber(((cdbl(rsha("SBI"))/cdbl(palm_SBI)) - (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)) ),2)%></td>
					<% else %>
						<td align="right"><%=formatnumber(((cdbl(rsha("SBI"))/cdbl(palm_SBI)) - (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)) )/(cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)),2)%></td>
					<% End if %>
					<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)),2)%></td>
				</tr>
              					
                <tr>  							
					<td  style="background-color:#F4F4A0" colspan="2"><b>Total Keseluruhan</td>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpBI)/cdbl(palm_BI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)),2)%></td>	
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudBI,2)%></td>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(PALMbud_BI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(TBSbud_BI)),2)%></td>
					<% If ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) = 0 then %>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRAmtgrpBI)/cdbl(palm_BI)) - ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) ) ,2)%></td>
					<% Else %>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRAmtgrpBI)/cdbl(palm_BI)) - ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) ) / ( cdbl(GRbudBI)/cdbl(PALMbud_BI)),2)%></td>
					<% End If %>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudSBI,2)%></td>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(TBSbud_SBI)),2)%></td>
					<% IF ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) = 0 then %>
						<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)) - ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) ) ,2)%></td>
					<% Else %>
						<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)) - ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) ) / ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
					<% End If %>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudthn,2)%></td>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(PALMbud_THN)),2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(TBSbud_THN)),2)%></td>														
				<tr>  
          <% END IF %>				
    <%        
	                 GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
					 GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI")) 
					 GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))
					 GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
					 GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))
		      rsha.MoveNext
		  wend
	end if
	%>				
</table>
<!--#include file="bottom.Asp"-->