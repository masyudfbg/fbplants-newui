<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, no
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

Dim rsDefaultBa, strDefaultBa, varDefaultBa
strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE'" & _
               "       and fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("fcba")
Else
	varDefaultBa = Session("fcba")
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

v_txtFcba = Request("txtFcba")
If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

Dim v_txtMonth : v_txtMonth = Request("vMonth")
Dim v_txtYear : v_txtYear = Request("vYear")

Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
strcurrentdate = " select currentdate" & _
                 " from periodcontrol" & _                 
                 " where fcba = '" & v_txtFcba & "'"

Set rscurrentdate = Server.CreateObject("ADODB.Recordset")
rscurrentdate.Open strcurrentdate, connstr, adOpenStatic

If Not rscurrentdate.EOF Then
	varDefaultYear	=	Year(rscurrentdate("currentdate"))
	varDefaultMonth	=	Month(rscurrentdate("currentdate"))
Else
	varDefaultYear	=	Year(Session("v_Employee"))
	varDefaultMonth	=	Month(Session("v_Employee"))
End If
rscurrentdate.Close
set rscurrentdate = Nothing

If (v_txtYear = "" Or Len(Trim(v_txtYear)) = 0) Then v_txtYear = varDefaultYear
If (v_txtMonth = "" Or Len(Trim(v_txtMonth)) = 0) Then v_txtMonth = varDefaultMonth

Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName
strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			"     , FCMANAGER" & _
			"     , FCKTU" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varManagerName = rsBaInfo("FCMANAGER")
	varKTUName = rsBaInfo("FCKTU")
Else
	varFcbaName = ""
	varCompanyName = ""
	varManagerName = ""
	varKTUName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing
%>
<html>

<head>
<title>ePlaS Reporting - Rekap Payroll</title>
<link rel="stylesheet" href="/css/font/font-awesome.css"/>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />

<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>

<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
  }, false
);
</script>
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
@media all {
#footer_note {width:65%;margin:auto;margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:24%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
}
@media print
{
#footer_note {margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:24%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.printhide{display:none;}
@page { 
 size:Letter landscape;
}
}

</style>
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div class="printhide">
				<input type="button" value="&#xf02f;" onClick="window.print()" style="font-family: 'FontAwesome';margin-top:-10px;">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
			<form name="frmAR" method="post" action="<%=ScriptName%>">
				<div class="printhide">
					<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
					<table id="form-table" class="centered-margin">
						<tr>
						<th>TAHUN</th>
						<th>BULAN</th>
						<th>&nbsp;</th>
						</tr>
						<tr>
						
								
									<td>
										<font size="2"  color="#0000FF" ></font>
						<% 
							mon = varDefaultMonth
							If len(v_txtMonth) = 0 then v_txtMonth = mon

							Yr = varDefaultYear
							MinYr = Yr - 2
							MaxYr = Yr  + 2
							If len(v_txtYear) = 0 then v_txtYear = Yr
						%>
										<select name="vYear">
											<%
												For i = MinYr to MaxYr
											%>
													<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
											<%
												Next
											%>
										</select>
									</td>
									
									<td>
										<font size="2" color="#0000FF" ></font>
											<select name="vMonth">
												<%
													For i = 1 to 12
												%>
														<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
												<%
													Next
												%>
											</select>
									</td>		

									


									
						<td><input type="submit" value="View" /></td>
						
						</tr>
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="17" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
			<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
				<h4>PT. <%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
		<div id="subheader" class="centered">
			<h2>REKAP GAJI KARYAWAN</h2>
			<h5>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h5>			 
		</div>
	</div>
    </th>
  </tr>
  
     <tr>
		
		<th rowspan="2">Seksi</th>		
		<th rowspan="2">Status</th>		
		<th rowspan="2">Gaji</th>
		<th rowspan="2">Rapel Gaji</th>
		<th rowspan="2">Lembur</th>
		<th rowspan="2">Premi</th>
		<th rowspan="2">Tunjangan</th>
		<th rowspan="2">Borongan</th>		
		<th colspan="7">Potongan</th>		
		<th rowspan="2">Total</th>
		<th rowspan="2">Tanda Tangan</th>
	</tr>
	<tr>
		<th>Jamsostek</th>
		<th>BPJS</th>
		<th>MK/P1</th>
		<th>Panen</th>
		<th>PPh21</th>
		<th>Koperasi</th>
		<th>Lain-Lain</th>
	</tr>
    </thead>
    <tbody>
	<%
	Dim rsha,strdtl , strgrp ,cdgrp1,cdgrp,cdgrpdesc  ,noline 
	Dim subAmtbsal,subAmtrapelgaji, subAmtOT,subAmtpremi,subAmtTunj,subAmtBorong,subAmtJamsos,subAmtbpjskes,subAmtMKP1,subAmtPanen,subAmtPPH21,subAmtPlain,subAmtPKoperasi, subAmtgrp
	Dim grnAmtbsal,grnAmtrapelgaji, grnAmtOT,grnAmtpremi,grnAmtTunj,grnAmtBorong,grnAmtJamsos,grnAmtbpjskes,grnAmtMKP1,grnAmtPanen,grnAmtPPH21,grnAmtPlain,grnAmtPKoperasi, grnAmtgrp
			strdtl = " select t.YearMonth" & _
			         "        , t.fcba" & _
			         "        , t.sectionname" & _
			         "        , t.presentgrade" & _
			         "        , sum(t.base_salary) base_salary" & _
			         "        , sum(t.rapelgaji) rapelgaji" & _
			         "        , sum(t.Overtime) Overtime" & _
			         "        , sum(t.Premi) Premi" & _
			         "        , sum(t.Tunjangan) Tunjangan" & _
			         "        , sum(t.Borongan) Borongan" & _
			         "        , sum(t.P_Jamsostek) P_Jamsostek" & _
					 "        , sum(t.p_bpjskes) p_bpjskes" & _
			         "        , sum(t.P_MK_P1) P_MK_P1" & _
			         "        , sum(t.P_Panen) P_Panen" & _
			         "        , sum(t.P_PPH21) P_PPH21" & _
			         "        , sum(t.p_koperasi) p_koperasi" & _
					 "        , sum(t.P_Lain) P_Lain" & _
			         "        , sum(t.base_salary" & _
			         "              + t.rapelgaji" & _
			         "              + t.Overtime" & _
			         "              + t.Premi" & _
			         "              + t.Tunjangan" & _
			         "              + t.Borongan" & _
			         "              + t.P_Jamsostek" & _
					 "              + p_bpjskes" & _
			         "              + t.P_MK_P1" & _
			         "              + t.P_Panen" & _
			         "              + t.P_PPH21" & _
					 "              + p_koperasi" & _
			         "              + t.P_Lain) total " & _
					 " from VIEW_PAYROLLDETAIL t" & _
					 " where t.presentgrade ='KBT'" & _
					 "       and TO_NUMBER(SUBSTR(T.YEARMONTH,1,4)) = " & v_txtYear & "" & _
					 "       and TO_NUMBER(SUBSTR(T.YEARMONTH,5,2)) = " & v_txtMonth & "" & _
					 "       and t.fcba = '" & v_txtFcba & "' " & _
					 " group by t.YearMonth" & _
					 "          , t.fcba" & _
					 "          , t.sectionname" & _
					 "          , t.presentgrade" & _
					 " order by t.YearMonth" & _
					 "          , t.sectionname" & _
					 "          , t.presentgrade" & _
					 "          , t.fcba" 
			
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    
	  

	if Not rsha.eof then
		while Not rsha.eof  %>
		           
		  <tr>	 
			<td align="center"><%=rsha("sectionname")%></td>
			
			<td align="center"><%=rsha("presentgrade")%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("base_salary")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("rapelgaji")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("Overtime")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("Premi")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("Tunjangan")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("Borongan")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_Jamsostek")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("p_bpjskes")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_MK_P1")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_Panen")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_PPH21")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("p_koperasi")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("P_Lain")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsha("total")),0)%></td>
			<% if cdgrp1 = ""  then 
			 
						  Dim rs1 , sql1, r1
   						  sql1    = "  select count(sectionname) rowsd  from " & _
								   "  (select t.YearMonth,t.fcba,t.sectionname,t.presentgrade " & _
								   "  from VIEW_PAYROLLDETAIL t " & _
								   "  where  TO_NUMBER(SUBSTR(T.YEARMONTH,1,4)) = " & v_txtYear & _
								   " and TO_NUMBER(SUBSTR(T.YEARMONTH,5,2)) = " & v_txtMonth & " and t.fcba ='" & v_txtFcba & "' " & _
								   "  and t.presentgrade = 'KBT'  " & _ 
								   "  group by t.YearMonth,t.fcba,t.sectionname,t.presentgrade " & _
								   "  order by t.YearMonth,t.sectionname,decode(t.presentgrade,'KBT',1,'KHT',2,'KHL',3),t.presentgrade,t.fcba) " 
								  
								   
                        Set rs1 = Server.CreateObject("ADODB.Recordset")
	                    rs1.Open sql1, connstr, adOpenStatic
						if Not rs1.eof then
						   r1 = rs1("rowsd")
						end if
						rs1.close
						
				%>
			<td align="center" rowspan="<%=cdbl(r1)+1%>" ></td>
			<% end if %>
		  </tr>	
		      
    <%        
	        ' cari sub total
			
	          subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary")) 
			  subAmtrapelgaji = cdbl(subAmtrapelgaji) + cdbl(rsha("rapelgaji"))
	          subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
			  subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
			  subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
	          subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
	          subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
			  subAmtbpjskes = cdbl(subAmtbpjskes) + cdbl(rsha("p_bpjskes"))
	          subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
			  subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
	          subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
			  subAmtPKoperasi = cdbl(subAmtPKoperasi) + cdbl(rsha("p_koperasi"))
	          subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
	          subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
			  
			 ' cari grand total 
			  
			  grnAmtbsal = cdbl(grnAmtbsal) + cdbl(rsha("base_salary")) 
			  grnAmtrapelgaji = cdbl(grnAmtrapelgaji) + cdbl(rsha("rapelgaji"))
	          grnAmtOT = cdbl(grnAmtOT) + cdbl(rsha("Overtime"))
			  grnAmtpremi = cdbl(grnAmtpremi) + cdbl(rsha("Premi")) 
			  grnAmtTunj = cdbl(grnAmtTunj) + cdbl(rsha("Tunjangan")) 
	          grnAmtBorong = cdbl(grnAmtBorong) + cdbl(rsha("Borongan")) 
	          grnAmtJamsos = cdbl(grnAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
			  grnAmtbpjskes = cdbl(grnAmtbpjskes) + cdbl(rsha("p_bpjskes"))
	          grnAmtMKP1 = cdbl(grnAmtMKP1) + cdbl(rsha("P_MK_P1")) 
			  grnAmtPanen = cdbl(grnAmtPanen) + cdbl(rsha("P_Panen")) 
	          grnAmtPPH21 = cdbl(grnAmtPPH21) + cdbl(rsha("P_PPH21")) 
			  grnAmtPKoperasi = cdbl(grnAmtPKoperasi) + cdbl(rsha("p_koperasi")) 
	          grnAmtPlain = cdbl(grnAmtPlain) + cdbl(rsha("P_Lain")) 
	          grnAmtgrp = cdbl(grnAmtgrp) + cdbl(rsha("Total"))
			  cdgrp1 = rsha("sectionname") 
		      rsha.MoveNext
		 wend
		%> 
		 <tr>	 
			<td align="center" colspan="2" ><b>Sub Total</td>
			<td align="right"><b><%=formatnumber(subAmtbsal,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtrapelgaji,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtOT,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtpremi,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtTunj,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtBorong,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtJamsos,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtbpjskes,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtMKP1,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPanen,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPPH21,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPKoperasi,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPlain,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtgrp,0)%></td>
		 </tr>	
		 <tr>	 
			<td align="center" colspan="17" ></td>
		 </tr>	
		 
	<%	 
	end if
					 subAmtbsal = cdbl("0") 
					 subAmtrapelgaji = cdbl("0")
					 subAmtOT = cdbl("0") 
					 subAmtpremi = cdbl("0") 
					 subAmtTunj = cdbl("0") 
					 subAmtBorong = cdbl("0") 
					 subAmtJamsos = cdbl("0") 
					 subAmtbpjskes = cdbl("0")
					 subAmtMKP1 = cdbl("0") 
					 subAmtPanen = cdbl("0") 
					 subAmtPPH21 = cdbl("0") 
					 subAmtPKoperasi = cdbl("0")
					 subAmtPlain = cdbl("0") 
					 subAmtgrp = cdbl("0")
	%>	
        
    <%
		
			strdtl = " select t.YearMonth" & _
			         "        , t.fcba" & _
			         "        , t.sectionname" & _
			         "        , t.presentgrade" & _
			         "        , sum(t.base_salary) base_salary" & _
			         "        , sum(t.rapelgaji) rapelgaji" & _
			         "        , sum(t.Overtime)  Overtime" & _
			         "        , sum(t.Premi) Premi" & _
			         "        , sum(t.Tunjangan) Tunjangan" & _
			         "        , sum(t.Borongan) Borongan" & _
			         "        , sum(t.P_Jamsostek) P_Jamsostek" & _
			         "        , sum(t.p_bpjskes) p_bpjskes" & _
			         "        , sum(t.P_MK_P1) P_MK_P1" & _
			         "        , sum(t.P_Panen) P_Panen" & _
			         "        , sum(t.P_PPH21) P_PPH21" & _
			         "        , sum(t.p_koperasi) p_koperasi" & _
					 "        , sum(t.P_Lain) P_Lain" & _
			         "        , sum(t.base_salary" & _
			         "              + t.rapelgaji" & _
			         "              + t.Overtime" & _
			         "              + t.Premi" & _
			         "              + t.Tunjangan" & _
			         "              + t.Borongan" & _
			         "              + t.P_Jamsostek" & _
			         "              + p_bpjskes" & _
			         "              + t.P_MK_P1" & _
			         "              + t.P_Panen" & _
			         "              + t.P_PPH21" & _
					 "              + t.p_koperasi" & _
			         "              + t.P_Lain) total " & _
					 " from VIEW_PAYROLLDETAIL t   " & _
					 " where  t.presentgrade <> 'KBT' and TO_NUMBER(SUBSTR(T.YEARMONTH,1,4)) = " & v_txtYear & _
					 " and TO_NUMBER(SUBSTR(T.YEARMONTH,5,2)) = " & v_txtMonth & " and t.fcba ='" & v_txtFcba & "' " & _
					 " group by t.YearMonth,t.fcba,t.sectionname,t.presentgrade   " & _
					 " order by t.YearMonth,t.sectionname,decode(t.presentgrade,'KBT',1,'KHT',2,'KHL',3),t.presentgrade,t.fcba   " 
			
	
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    
	  

	if Not rsha.eof then
		while Not rsha.eof 
			  IF cdgrp = rsha("sectionname")  then %>
					  <tr>	 
						
						<td align="center"><%=rsha("presentgrade")%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("base_salary")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("rapelgaji")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Overtime")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Premi")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Tunjangan")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Borongan")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Jamsostek")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("p_bpjskes")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_MK_P1")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Panen")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_PPH21")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("p_koperasi")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Lain")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("total")),0)%></td>
						
					  </tr>	
					  <%
								
					' cari sub total
						subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary")) 
						subAmtrapelgaji = cdbl(subAmtrapelgaji) + cdbl(rsha("rapelgaji"))
						subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
						subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
						subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
						subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
						subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
						subAmtbpjskes = cdbl(subAmtbpjskes) + cdbl(rsha("p_bpjskes"))
						subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
						subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
						subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
						subAmtPKoperasi = cdbl(subAmtPKoperasi) + cdbl(rsha("p_koperasi"))
						subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
						subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
							  
					' cari grand total 
							  
						grnAmtbsal = cdbl(grnAmtbsal) + cdbl(rsha("base_salary")) 
						grnAmtrapelgaji = cdbl(grnAmtrapelgaji) + cdbl(rsha("rapelgaji"))
						grnAmtOT = cdbl(grnAmtOT) + cdbl(rsha("Overtime"))
						grnAmtpremi = cdbl(grnAmtpremi) + cdbl(rsha("Premi")) 
						grnAmtTunj = cdbl(grnAmtTunj) + cdbl(rsha("Tunjangan")) 
						grnAmtBorong = cdbl(grnAmtBorong) + cdbl(rsha("Borongan")) 
						grnAmtJamsos = cdbl(grnAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
						grnAmtbpjskes = cdbl(grnAmtbpjskes) + cdbl(rsha("p_bpjskes"))
						grnAmtMKP1 = cdbl(grnAmtMKP1) + cdbl(rsha("P_MK_P1")) 
						grnAmtPanen = cdbl(grnAmtPanen) + cdbl(rsha("P_Panen")) 
						grnAmtPPH21 = cdbl(grnAmtPPH21) + cdbl(rsha("P_PPH21")) 
						grnAmtPKoperasi = cdbl(grnAmtPKoperasi) + cdbl(rsha("p_koperasi"))
						grnAmtPlain = cdbl(grnAmtPlain) + cdbl(rsha("P_Lain")) 
						grnAmtgrp = cdbl(grnAmtgrp) + cdbl(rsha("Total"))
					  
					  
			 ELSE    
    			 if cdgrp <> "" then %>	  
					  <tr>	 
						<td align="center" colspan="2" ><b>Sub Total <%=cdgrp%></td>
						<td align="right"><b><%=formatnumber(subAmtbsal,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtrapelgaji,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtOT,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtpremi,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtTunj,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtBorong,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtJamsos,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtbpjskes,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtMKP1,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtPanen,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtPPH21,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtPKoperasi,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtPlain,0)%></td>
						<td align="right"><b><%=formatnumber(subAmtgrp,0)%></td>
					 </tr>	
					 <tr>	 
						<td align="center" colspan="17" ></td>
					 </tr>
		<%
					 subAmtbsal = cdbl("0") 
					 subAmtrapelgaji = cdbl("0")
					 subAmtOT = cdbl("0") 
					 subAmtpremi = cdbl("0") 
					 subAmtTunj = cdbl("0") 
					 subAmtBorong = cdbl("0") 
					 subAmtJamsos = cdbl("0") 
					 subAmtbpjskes = cdbl("0")
					 subAmtMKP1 = cdbl("0") 
					 subAmtPanen = cdbl("0") 
					 subAmtPPH21 = cdbl("0") 
					 subAmtPKoperasi = cdbl("0")
					 subAmtPlain = cdbl("0") 
					 subAmtgrp = cdbl("0")
                 end if	 %>
				     <%
						  Dim rsrow , sqlrow, rows
   						  sqlrow = "  select YearMonth,fcba,sectionname,count(sectionname) rowsd  from " & _
								   "  (select t.YearMonth,t.fcba,t.sectionname,t.presentgrade " & _
								   "  from VIEW_PAYROLLDETAIL t " & _
								   "  where  TO_NUMBER(SUBSTR(T.YEARMONTH,1,4)) = " & v_txtYear & _
								   " and TO_NUMBER(SUBSTR(T.YEARMONTH,5,2)) = " & v_txtMonth & " and t.fcba ='" & v_txtFcba & "' " & _
								   "  and t.presentgrade <> 'KBT' and t.sectionname = '" & rsha("sectionname") & "'  " & _ 
								   "  group by t.YearMonth,t.fcba,t.sectionname,t.presentgrade " & _
								   "  order by t.YearMonth,t.sectionname,decode(t.presentgrade,'KBT',1,'KHT',2,'KHL',3),t.presentgrade,t.fcba) " & _
								   "  group by YearMonth,fcba,sectionname "
								   
                        Set rsrow = Server.CreateObject("ADODB.Recordset")
	                    rsrow.Open sqlrow, connstr, adOpenStatic
						if Not rsha.eof then
						   rows = rsrow("rowsd")
						end if
					 %>
						
                     <tr>	 
						<td align="center" rowspan="<%=cdbl(rows)%>" ><%=rsha("sectionname")%></td>
						<td align="center"><%=rsha("presentgrade")%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("base_salary")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("rapelgaji")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Overtime")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Premi")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Tunjangan")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("Borongan")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Jamsostek")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("p_bpjskes")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_MK_P1")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Panen")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_PPH21")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("p_koperasi")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("P_Lain")),0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("total")),0)%></td>
						<td align="center" rowspan="<%=cdbl(rows)+1%>" ></td>
					  </tr>	
					  <%
								
					' cari sub total
						subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary")) 
						subAmtrapelgaji = cdbl(subAmtrapelgaji) + cdbl(rsha("rapelgaji"))
						subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
						subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
						subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
						subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
						subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
						subAmtbpjskes = cdbl(subAmtbpjskes) + cdbl(rsha("p_bpjskes"))
						subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
						subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
						subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
						subAmtPKoperasi = cdbl(subAmtPKoperasi) + cdbl(rsha("p_koperasi"))
						subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
						subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
							  
					' cari grand total 
							  
						grnAmtbsal = cdbl(grnAmtbsal) + cdbl(rsha("base_salary")) 
						grnAmtrapelgaji = cdbl(grnAmtrapelgaji) + cdbl(rsha("rapelgaji"))
						grnAmtOT = cdbl(grnAmtOT) + cdbl(rsha("Overtime"))
						grnAmtpremi = cdbl(grnAmtpremi) + cdbl(rsha("Premi")) 
						grnAmtTunj = cdbl(grnAmtTunj) + cdbl(rsha("Tunjangan")) 
						grnAmtBorong = cdbl(grnAmtBorong) + cdbl(rsha("Borongan")) 
						grnAmtJamsos = cdbl(grnAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
						grnAmtbpjskes = cdbl(grnAmtbpjskes) + cdbl(rsha("p_bpjskes"))
						grnAmtMKP1 = cdbl(grnAmtMKP1) + cdbl(rsha("P_MK_P1")) 
						grnAmtPanen = cdbl(grnAmtPanen) + cdbl(rsha("P_Panen")) 
						grnAmtPPH21 = cdbl(grnAmtPPH21) + cdbl(rsha("P_PPH21")) 
						grnAmtPKoperasi = cdbl(grnAmtPKoperasi) + cdbl(rsha("p_koperasi"))
						grnAmtPlain = cdbl(grnAmtPlain) + cdbl(rsha("P_Lain")) 
						grnAmtgrp = cdbl(grnAmtgrp) + cdbl(rsha("Total"))				 
			 END IF      
			 cdgrp = rsha("sectionname")
			 rsha.MoveNext
		wend
	end if
	%>	
	    <tr>	 
			<td align="center" colspan="2" ><b>Sub Total <%=cdgrp%></td>
			<td align="right"><b><%=formatnumber(subAmtbsal,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtrapelgaji,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtOT,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtpremi,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtTunj,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtBorong,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtJamsos,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtbpjskes,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtMKP1,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPanen,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPPH21,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPKoperasi,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtPlain,0)%></td>
			<td align="right"><b><%=formatnumber(subAmtgrp,0)%></td>
		 </tr>	
		 <tr>	 
			<td align="center" colspan="17" ></td>
		 </tr>
        <tr>	 
			<td align="center" colspan="2" ><b>Grand Total</td>
			<td align="right"><b><%=formatnumber(grnAmtbsal,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtrapelgaji,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtOT,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtpremi,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtTunj,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtBorong,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtJamsos,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtbpjskes,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtMKP1,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtPanen,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtPPH21,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtPKoperasi,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtPlain,0)%></td>
			<td align="right"><b><%=formatnumber(grnAmtgrp,0)%></td>
			<td align="right"></td>
		</tr>	
        </tbody>			  
</table>
<div id="footer_note">	
	<div class="block">Prepared by,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
    <div class="block">Checked by,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
    <div class="block">Approved by,
        <br/><br/><br/><br/>
        <input class="isian" type="text" value="<%=varManagerName%>">
        <hr/>
        <input class="isian" type="text" value="MANAGER">
    </div>
	<div class="block">Acknowledged by,
        <br/><br/><br/><br/>
        <input class="isian" type="text" value="" >
        <hr/>
        <input class="isian" type="text" value="">
    </div>
 </div>
 <div class="clear"></div>