<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, no
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

Dim v_txttampilkan : v_txttampilkan = Request("txttampilkan")

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

Dim rsrow , sqlrow, rows


Dim RowsSection : RowsSection = 0
Dim RowsGang : RowsGang = 0

Dim rsRowsGang, sqlRowsGang
Dim rsRowsSection, sqlRowsSection
Dim rsRowsSectionGang, sqlRowsSectionGang, RowSpanSectionGang

Dim tempRowsGang, tempRowsSection, tempRowsSectionGang, temprows

	Dim rsha,strdtl , strgrp ,cdgrp1,cdgrp,cdgrpdesc  ,noline 
	Dim subAmtbsal,subAmtrapelgaji, subAmtOT,subAmtpremi,subAmtTunj,subAmtBorong,subAmtJamsos,subAmtbpjskes,subAmtMKP1,subAmtPanen,subAmtPPH21,subAmtPlain,subAmtPKoperasi, subAmtgrp
	Dim grnAmtbsal,grnAmtrapelgaji, grnAmtOT,grnAmtpremi,grnAmtTunj,grnAmtBorong,grnAmtJamsos,grnAmtbpjskes,grnAmtMKP1,grnAmtPanen,grnAmtPPH21,grnAmtPlain,grnAmtPKoperasi, grnAmtgrp
	Dim st_jumlah_orang, st_hk_month, st_hk, st_mkp1
	Dim gt_jumlah_orang, gt_hk_month, gt_hk, gt_mkp1
	
	Dim st_det_jumlah_orang, st_det_hk_month, st_det_hk, st_det_mkp1
	Dim st_det_Amtbsal,st_det_Amtrapelgaji, st_det_AmtOT,st_det_Amtpremi,st_det_AmtTunj,st_det_AmtBorong,st_det_AmtJamsos,st_det_Amtbpjskes,st_det_AmtMKP1,st_det_AmtPanen,st_det_AmtPPH21,st_det_AmtPlain,st_det_AmtPKoperasi, st_det_Amtgrp
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
						<th>TAMPILKAN</th>
						<th>&nbsp;</th>
						</tr>
						<tr>
						
								
									<td>
										<font size="2"  color="#0000FF" ></font>
						<% 
							mon = varDefaultMonth
							If len(v_txtMonth) = 0 then v_txtMonth = mon

							Yr = varDefaultYear
							MinYr = 2015
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
									<td>
										<select name="txttampilkan">
											<Option Value="" <%If len(trim(v_txttampilkan)) = 0 then%> Selected <%End If%>>Rekap</option>
											<Option Value="detail" <%If v_txttampilkan = "detail" then%> Selected <%End If%>>Detail</option>
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
    <th colspan="21" style="background-color:#FFFFFF; border:none;">  
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
		<%if len(trim(v_txttampilkan)) <> 0 then%>
			<th rowspan="2">Gang</th>
		<%end if%>
		<th rowspan="2">Status</th>	
		<th colspan="4">Jumlah</th>		
		<th rowspan="2">Gaji</th> 
		<th rowspan="2">Lembur</th>
		<th rowspan="2">Premi</th> 
		<th rowspan="2">Borongan</th>
		<th rowspan="2">Total Upah</th>		
		<th colspan="6">Potongan</th>		
		<th rowspan="2">Total</th>
		<%if len(trim(v_txttampilkan)) = 0 then%>
			<th rowspan="2">Tanda Tangan</th>
		<%end if%>
	</tr>
	<tr>
		<th>Orang</th>
		<th>HK Sebulan</th>
		<th>HK Kerja</th>
		<th>MK/P1</th>
		<th>BPJS Tenaga Kerja</th>
		<th>BPJS Kesehatan</th>
		<th>PPh21</th>
		<th>Pinjaman</th>
		<th>Denda</th>	
		<th>Total Potongan</th>	
	</tr>
    </thead>
    <tbody>
<%if len(trim(v_txttampilkan)) = 0 then%> <!-- JIKA REKAP -->
	<%

			strdtl = 	"   SELECT *"  & _
						"     FROM (  SELECT year,"  & _
						"                    month,"  & _
						"                    fcba,"  & _
						"                    employee_api.get_presentgrade2 (fcba, fccode) presentgrade2,"  & _
						"                    SUM (NVL (jumlah_orang, 0)) jumlah_orang,"  & _
						"                    employee_hist_api.get_sectionname (year,"  & _
						"                                                       month,"  & _
						"                                                       fcba,"  & _
						"                                                       fccode)"  & _
						"                       sectionname,"  & _
						"                    SUM (NVL (hk_month, 0)) hk_month,"  & _
						"                    SUM (NVL (hk, 0)) hk,"  & _
						"                    SUM (NVL (mkp1, 0)) mkp1,"  & _
						"                    SUM (NVL (base_salary, 0)) base_salary,"  & _
						"                    SUM (NVL (overtime, 0)) overtime,"  & _
						"                    SUM (NVL (borongan, 0)) borongan,"  & _
						"                    SUM (NVL (p_jamsostek, 0)) p_jamsostek,"  & _
						"                    SUM (NVL (p_bpjskes, 0)) p_bpjskes,"  & _
						"                    SUM (NVL (p_pph21, 0)) p_pph21,"  & _
						"                    SUM (NVL (premi, 0)) premi,"  & _
						"                    SUM (NVL (totalgaji, 0)) totalgaji,"  & _
						"                    SUM (NVL (denda, 0)) denda,"  & _
						"                    SUM (NVL (total, 0)) total,"  & _
						"                    SUM (NVL (p_pinjaman, 0)) pinjaman"  & _
						"               FROM (  SELECT TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)) year,"  & _
						"                              TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)) month,"  & _
						"                              t.fcba,"  & _
						"                              t.fccode,"  & _
						"                              1 jumlah_orang,"  & _
						"                              SUM (get_hk) hk_month,"  & _
						"                              SUM (get_hk_kj) hk,"  & _
						"                              SUM (get_mkp1) mkp1,"  & _
						"                              SUM (t.base_salary) base_salary,"  & _
						"                              SUM (t.rapelgaji) + SUM (t.tunjangan) + SUM (t.premi)"  & _
						"                                 premi,"  & _
						"                              SUM (t.overtime) overtime,"  & _
						"                              SUM (t.borongan) borongan,"  & _
						"                              SUM (t.p_jamsostek) p_jamsostek,"  & _
						"                              SUM (t.p_bpjskes) p_bpjskes,"  & _
						"                              SUM (t.p_pph21) p_pph21,"  & _
						"                              SUM (t.p_pinjaman) p_pinjaman,"  & _
						"                                SUM (t.p_panen)"  & _
						"                              + SUM (t.p_mk_p1)"  & _
						"                              + SUM (t.p_koperasi)"  & _
						"                              + SUM (t.p_lain)"  & _
						"                                 denda,"  & _
						"                                SUM (t.base_salary)"  & _
						"                              + SUM (t.overtime)"  & _
						"                              + SUM (t.borongan)"  & _
						"                              + SUM (t.premi)"  & _
						"                                 totalgaji,"  & _
						"                                SUM (t.base_salary)"  & _
						"                              + SUM (t.overtime)"  & _
						"                              + SUM (t.premi)"  & _
						"                              + SUM (t.rapelgaji)"  & _
						"                              + SUM (t.tunjangan)"  & _
						"                              + SUM (t.borongan)"  & _
						"                              + SUM (t.p_jamsostek)"  & _
						"                              + SUM (t.p_bpjskes)"  & _
						"                              + SUM (t.p_mk_p1)"  & _
						"                              + SUM (t.p_panen)"  & _
						"                              + SUM (t.p_pph21)"  & _
						"                              + SUM (t.p_koperasi)"  & _
						"                              + SUM (t.p_lain)"  & _
						"                                 total"  & _
						"                         FROM view_payrolldetail t"  & _
						"                              INNER JOIN EMPLOYEE emp"  & _
						"                                 ON (emp.fccode = t.fccode AND emp.fcba = t.fcba)"  & _
						"                              INNER JOIN PARAMETER pr"  & _
						"                                 ON (    pr.fcba = emp.fcba"  & _
						"                                     AND pr.fccode = emp.emp_type"  & _
						"                                     AND parcode = 'E0103')"  & _
						"                        WHERE     TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)) = '"& v_txtYear &"'"  & _
						"                              AND TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)) = '"& v_txtMonth &"'"  & _
						"                              AND t.fcba = '"& v_txtFcba &"'"  & _
						"                     GROUP BY TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)),"  & _
						"                              TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)),"  & _
						"                              employee_api.get_presentgrade2 (t.fcba, t.fccode),"  & _
						"                              t.fcba,"  & _
						"                              t.fccode "  & _
						"                     HAVING SUM (t.base_salary) + SUM (t.overtime)+ SUM (t.borongan) + SUM (t.premi) > 0 " & _ 
						"                      )"  & _
						"           GROUP BY year,"  & _
						"                    month,"  & _
						"                    fcba,"  & _
						"                    employee_api.get_presentgrade2 (fcba, fccode),"  & _
						"                    employee_hist_api.get_sectionname (year,"  & _
						"                                                       month,"  & _
						"                                                       fcba,"  & _
						"                                                       fccode))"  & _
						"    WHERE fcba = '"& v_txtFcba &"' AND totalgaji > 0" 
								 
			
				strdtl = strdtl & " order by sectionname  asc, PRESENTGRADE2 desc " 
	' response.write strdtl
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    
	  
    totalpotongan=0
	totalsisaupah=0
	if Not rsha.eof then
		rows = 1
		while Not rsha.eof 
		totalpotongan = cdbl(to_nol(rsha("p_jamsostek"))) + cdbl(to_nol(rsha("p_bpjskes"))) + cdbl(to_nol(rsha("p_pph21"))) + cdbl(to_nol(rsha("pinjaman"))) + cdbl(to_nol(rsha("denda"))) 
		totalsisaupah = cdbl(to_nol(rsha("totalgaji"))) - (-1 * cdbl(to_nol(totalpotongan)) )
%>
			<tr>
<%
			if rows = 1 then
				sqlrow = "  select count(sectionname) rowsd  from ( " & strdtl & " ) where sectionname = '" & rsha("sectionname") & "'"
				'response.write sqlrow
				Set rsrow = Server.CreateObject("ADODB.Recordset")
				rsrow.Open sqlrow, connstr, adOpenStatic
				if not rsrow.EOF then
					temprows = CDbl(rsrow("rowsd"))
				else
					temprows = 0
				end if
			end if
%>

				<%if rows = 1 then%>
					<td align="center" Rowspan="<%=temprows+1 %>"><%=rsha("sectionname")%></td>
				<%end if%>
					<td align="center"><%=rsha("presentgrade2")%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("jumlah_orang")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("hk_month")),1)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("hk")),1)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("mkp1")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("base_salary")),0)%></td> 
					<td align="right"><%=formatnumber( cdbl(rsha("Overtime")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("Premi")),0)%></td> 
					<td align="right"><%=formatnumber( cdbl(rsha("Borongan")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("totalgaji")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("P_JAMSOSTEK")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("P_BPJSKES")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("P_PPH21")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("PINJAMAN")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("DENDA")),0)%></td>
					
					<td align="right"><%=formatnumber( cdbl(totalpotongan),0)%></td>
					<td align="right"><%=formatnumber( cdbl(totalsisaupah),0)%></td>
				
				<%if rows = 1 then%>
					<td align="center" Rowspan="<%=temprows+1 %>"></td>
				<%end if%>
				  </tr>
<%
			if rows = temprows then
				' cari sub total
				' st_jumlah_orang = cdbl(st_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				' st_hk_month = cdbl(st_hk_month) + cdbl(rsha("hk_month")) 
				' st_hk = cdbl(st_hk) + cdbl(rsha("hk"))
				' st_mkp1 = cdbl(st_mkp1) + cdbl(rsha("mkp1")) 
				' subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary"))
				' subAmtrapelgaji = cdbl(subAmtrapelgaji) + cdbl(rsha("rapelgaji"))
				' subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
				' subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
				' subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
				' subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
				' subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
				' subAmtbpjskes = cdbl(subAmtbpjskes) + cdbl(rsha("p_bpjskes"))
				' subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
				' subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
				' subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
				' subAmtPKoperasi = cdbl(subAmtPKoperasi) + cdbl(rsha("p_koperasi"))
				' subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
				' subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
				n_jumlah_orang 	= n_jumlah_orang + cdbl(rsha("jumlah_orang"))
				nhk_month 		= nhk_month 	+ cdbl(rsha("hk_month"))
				nhk		 		= nhk 			+ cdbl(rsha("hk"))
				nmkp1 			= nmkp1		+ cdbl(rsha("mkp1"))
				nBase_Salary 	= nBase_Salary 	+ cdbl(rsha("BASE_SALARY"))
				nOvertime 		= nOvertime 	+ cdbl(rsha("OVERTIME"))
				nBorongan 		= nBorongan 	+ cdbl(rsha("BORONGAN"))
				nP_Jamsostek 	= nP_Jamsostek 	+ cdbl(rsha("P_JAMSOSTEK"))
				nP_bpjskes 		= nP_bpjskes 	+ cdbl(rsha("P_BPJSKES"))
				nTotal 			= nTotal 		+ cdbl(rsha("TOTAL"))
				nDenda 			= nDenda		+ cdbl(rsha("DENDA"))
				nTotalgaji 		= nTotalgaji	+ cdbl(rsha("TOTALGAJI"))
				nPremi 			= nPremi 		+ cdbl(rsha("PREMI"))
				nPinjaman 		= nPinjaman 	+ cdbl(rsha("PINJAMAN"))
				nP_PPH21 		= nP_PPH21 		+ cdbl(rsha("P_PPH21"))
				nTotalPotongan  = nTotalPotongan + cdbl(totalpotongan)
				ntotalsisaupah = ntotalsisaupah + cdbl(totalsisaupah)
				rows = 1
%>
				  <tr bgcolor="#F4F4A0">	 
					<td align="center" ><b>Total</td>
					<td align="right"><b><%=formatnumber(n_jumlah_orang,0)%></td>
					<td align="right"><b><%=formatnumber(nhk_month,1)%></td>
					<td align="right"><b><%=formatnumber(nhk,1)%></td>
					<td align="right"><b><%=formatnumber(nmkp1,0)%></td>
					<td align="right"><b><%=formatnumber(nBase_Salary,0)%></td> 
					<td align="right"><b><%=formatnumber(nOvertime,0)%></td>
					<td align="right"><b><%=formatnumber(nPremi,0)%></td>
					<td align="right"><b><%=formatnumber(nBorongan,0)%></td>
					<td align="right"><b><%=formatnumber(nTotalgaji,0)%></td>
					<td align="right"><b><%=formatnumber(nP_Jamsostek,0)%></td>
					<td align="right"><b><%=formatnumber(nP_bpjskes,0)%></td>
					<td align="right"><b><%=formatnumber(nP_PPH21,0)%></td>
					<td align="right"><b><%=formatnumber(nPinjaman,0)%></td>
					<td align="right"><b><%=formatnumber(nDenda,0)%></td>
					<td align="right"><b><%=formatnumber(nTotalPotongan,0)%></td>
					<td align="right"><b><%=formatnumber(ntotalsisaupah,0)%></td> 
				 </tr>	
				 <tr>
					<td align="center" colspan="19" ></td>
				 </tr>

<%
					 ' st_jumlah_orang = cdbl("0")
					 ' st_hk_month = cdbl("0")
					 ' st_hk = cdbl("0")
					 ' st_mkp1 = cdbl("0")
					 ' subAmtbsal = cdbl("0") 
					 ' subAmtrapelgaji = cdbl("0")
					 ' subAmtOT = cdbl("0") 
					 ' subAmtpremi = cdbl("0") 
					 ' subAmtTunj = cdbl("0") 
					 ' subAmtBorong = cdbl("0") 
					 ' subAmtJamsos = cdbl("0") 
					 ' subAmtbpjskes = cdbl("0")
					 ' subAmtMKP1 = cdbl("0") 
					 ' subAmtPanen = cdbl("0") 
					 ' subAmtPPH21 = cdbl("0") 
					 ' subAmtPKoperasi = cdbl("0")
					 ' subAmtPlain = cdbl("0") 
					 ' subAmtgrp = cdbl("0")
					n_jumlah_orang = cdbl("0")
					nhk_month = cdbl("0")
					nhk  = cdbl("0")
					nBase_Salary = cdbl("0")
					nOvertime = cdbl("0")
					nPremi = cdbl("0")
					nTunjangan = cdbl("0")
					nBorongan = cdbl("0")
					nP_Jamsostek = cdbl("0")
					nmkp1 = cdbl("0")
					nP_Panen = cdbl("0")
					nP_Pph21 = cdbl("0")
					nP_Lain = cdbl("0")
					nTotal = cdbl("0")
					nDenda = cdbl("0")
					nTotalgaji = cdbl("0")
					nPinjaman = cdbl("0")
					nTotalPotongan = cdbl("0")
					ntotalsisaupah = cdbl("0")
					nP_bpjskes = cdbl("0")
			else
				' cari sub total
					' st_jumlah_orang = cdbl(st_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
					' st_hk_month = cdbl(st_hk_month) + cdbl(rsha("hk_month"))
					' st_hk = cdbl(st_hk) + cdbl(rsha("hk"))
					' st_mkp1 = cdbl(st_mkp1) + cdbl(rsha("mkp1")) 
					' subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary"))
					' subAmtrapelgaji = cdbl(subAmtrapelgaji) + cdbl(rsha("rapelgaji"))
					' subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
					' subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
					' subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
					' subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
					' subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
					' subAmtbpjskes = cdbl(subAmtbpjskes) + cdbl(rsha("p_bpjskes"))
					' subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
					' subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
					' subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
					' subAmtPKoperasi = cdbl(subAmtPKoperasi) + cdbl(rsha("p_koperasi"))
					' subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
					' subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
					n_jumlah_orang 	= n_jumlah_orang + cdbl(rsha("jumlah_orang"))
					nhk_month 		= nhk_month 	+ cdbl(rsha("hk_month"))
					nhk		 		= nhk 			+ cdbl(rsha("hk"))
					nmkp1 			= nmkp1		+ cdbl(rsha("mkp1"))
					nBase_Salary 	= nBase_Salary 	+ cdbl(rsha("BASE_SALARY"))
					nOvertime 		= nOvertime 	+ cdbl(rsha("OVERTIME"))
					nBorongan 		= nBorongan 	+ cdbl(rsha("BORONGAN"))
					nP_Jamsostek 	= nP_Jamsostek 	+ cdbl(rsha("P_JAMSOSTEK"))
					nP_bpjskes 		= nP_bpjskes 	+ cdbl(rsha("P_BPJSKES"))
					nTotal 			= nTotal 		+ cdbl(rsha("TOTAL"))
					nDenda 			= nDenda		+ cdbl(rsha("DENDA"))
					nTotalgaji 		= nTotalgaji	+ cdbl(rsha("TOTALGAJI"))
					nPremi 			= nPremi 		+ cdbl(rsha("PREMI"))
					nPinjaman 		= nPinjaman 	+ cdbl(rsha("PINJAMAN"))
					nP_PPH21 		= nP_PPH21 		+ cdbl(rsha("P_PPH21"))
					nTotalPotongan  = nTotalPotongan + cdbl(totalpotongan)
					ntotalsisaupah = ntotalsisaupah + cdbl(totalsisaupah)
					rows = rows + 1
			end if
							  
			' cari grand total 
				' gt_jumlah_orang = cdbl(gt_jumlah_orang) + cdbl(rsha("jumlah_orang"))
				' gt_hk_month = cdbl(gt_hk_month) + cdbl(rsha("hk_month"))
				' gt_hk = cdbl(gt_hk) + cdbl(rsha("hk"))
				' gt_mkp1 = cdbl(gt_mkp1) + cdbl(rsha("mkp1"))
				' grnAmtbsal = cdbl(grnAmtbsal) + cdbl(rsha("base_salary")) 
				' grnAmtrapelgaji = cdbl(grnAmtrapelgaji) + cdbl(rsha("rapelgaji"))
				' grnAmtOT = cdbl(grnAmtOT) + cdbl(rsha("Overtime"))
				' grnAmtpremi = cdbl(grnAmtpremi) + cdbl(rsha("Premi")) 
				' grnAmtTunj = cdbl(grnAmtTunj) + cdbl(rsha("Tunjangan")) 
				' grnAmtBorong = cdbl(grnAmtBorong) + cdbl(rsha("Borongan")) 
				' grnAmtJamsos = cdbl(grnAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
				' grnAmtbpjskes = cdbl(grnAmtbpjskes) + cdbl(rsha("p_bpjskes"))
				' grnAmtMKP1 = cdbl(grnAmtMKP1) + cdbl(rsha("P_MK_P1")) 
				' grnAmtPanen = cdbl(grnAmtPanen) + cdbl(rsha("P_Panen")) 
				' grnAmtPPH21 = cdbl(grnAmtPPH21) + cdbl(rsha("P_PPH21")) 
				' grnAmtPKoperasi = cdbl(grnAmtPKoperasi) + cdbl(rsha("p_koperasi"))
				' grnAmtPlain = cdbl(grnAmtPlain) + cdbl(rsha("P_Lain")) 
				' grnAmtgrp = cdbl(grnAmtgrp) + cdbl(rsha("Total"))
				gt_jumlah_orang = cdbl(gt_jumlah_orang) + cdbl(rsha("jumlah_orang"))
				gt_hk_month = cdbl(gt_hk_month) + cdbl(rsha("hk_month"))
				gt_hk = cdbl(gt_hk) + cdbl(rsha("hk"))
				gt_mkp1 = cdbl(gt_mkp1) + cdbl(rsha("mkp1"))
				gt_BASE_SALARY = cdbl(gt_BASE_SALARY) + cdbl(rsha("BASE_SALARY"))
				gt_OVERTIME = cdbl(gt_OVERTIME) + cdbl(rsha("OVERTIME")) 
				gt_BORONGAN = cdbl(gt_BORONGAN) + cdbl(rsha("BORONGAN"))
				gt_P_JAMSOSTEK = cdbl(gt_P_JAMSOSTEK) + cdbl(rsha("P_JAMSOSTEK"))
				gt_P_BPJSKES = cdbl(gt_P_BPJSKES) + cdbl(rsha("P_BPJSKES")) 
				gt_TOTAL = cdbl(gt_TOTAL) + cdbl(rsha("TOTAL")) 
				gt_DENDA = cdbl(gt_DENDA) + cdbl(rsha("DENDA")) 
				gt_TOTALGAJI = cdbl(gt_TOTALGAJI) + cdbl(rsha("TOTALGAJI")) 
				gt_PREMI = cdbl(gt_PREMI) + cdbl(rsha("PREMI")) 
				gt_PINJAMAN = cdbl(gt_PINJAMAN) + cdbl(rsha("PINJAMAN"))  
				gt_P_PPH21 = cdbl(gt_P_PPH21) + cdbl(rsha("P_PPH21")) 
				gt_totalpotongan = cdbl(gt_totalpotongan) + cdbl(totalpotongan)
				gt_totalsisaupah = cdbl(gt_totalsisaupah) + cdbl(totalsisaupah) 
			
			 rsha.MoveNext
		wend
	end if
	%>	

<%
Dim rsgrandtotal, strgrandtotal, rowgrandtotal
strgrandtotal = " select fcba, presentgrade2," & _
				"					 SUM (NVL (jumlah_orang, 0)) jumlah_orang," & _
                "                    SUM (NVL (hk_month, 0)) hk_month,"  & _
				"                    SUM (NVL (hk, 0)) hk,"  & _
				"                    SUM (NVL (mkp1, 0)) mkp1,"  & _
				"                    SUM (NVL (base_salary, 0)) base_salary,"  & _
				"                    SUM (NVL (overtime, 0)) overtime,"  & _
				"                    SUM (NVL (borongan, 0)) borongan,"  & _
				"                    SUM (NVL (p_jamsostek, 0)) p_jamsostek,"  & _
				"                    SUM (NVL (p_bpjskes, 0)) p_bpjskes,"  & _
				"                    SUM (NVL (p_pph21, 0)) p_pph21,"  & _
				"                    SUM (NVL (premi, 0)) premi,"  & _
				"                    SUM (NVL (totalgaji, 0)) totalgaji,"  & _
				"                    SUM (NVL (denda, 0)) denda,"  & _
				"                    SUM (NVL (total, 0)) total,"  & _
				"                    SUM (NVL (pinjaman, 0)) pinjaman"  & _
				" from ( " & strdtl & " )" & _
				" group by fcba, presentgrade2" & _
				" order by employee_grade_api.get_orderdata(presentgrade2)"
' response.write strgrandtotal
Set rsgrandtotal = Server.CreateObject("ADODB.Recordset")
rsgrandtotal.Open strgrandtotal, connstr, adOpenStatic
	if Not rsgrandtotal.EOF then
		rowgrandtotal = 1
		While Not rsgrandtotal.EOF
		v_totalpotongan = cdbl(to_nol(rsha("p_jamsostek"))) + cdbl(to_nol(rsha("p_bpjskes"))) + cdbl(to_nol(rsha("p_pph21"))) + cdbl(to_nol(rsha("pinjaman"))) + cdbl(to_nol(rsha("denda"))) 
		v_totalsisaupah = cdbl(to_nol(rsha("totalgaji"))) - (-1 * cdbl(to_nol(totalpotongan)) )
%>
			<tr>
				<%if rowgrandtotal = 1 then%>
					<td align="center" rowspan="<%=4+1%>"><b>TOTAL</b></td>
				<%end if%>
				<td align="center"><%=rsgrandtotal("presentgrade")%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("jumlah_orang")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("hk_month")),1)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("hk")),1)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("mkp1")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("base_salary")),0)%></td> 
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("Overtime")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("Premi")),0)%></td> 
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("Borongan")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("totalgaji")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("P_Jamsostek")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("p_bpjskes")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("P_PPH21")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("pinjaman")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("denda")),0)%></td> 
				<td align="right"><%=formatnumber( cdbl(v_totalpotongan,0))%></td>
				<td align="right"><%=formatnumber( cdbl(v_totalsisaupah,0))%></td>
				<%if rowgrandtotal = 1 then%>
					<td align="center" rowspan="<%=4+1%>"></td>
				<%end if%>
			</tr>
			<%if rowgrandtotal = 4 then%>
				<tr bgcolor="#FFFF00">	 
					<td align="center"><b>Total</td>
					<td align="right"><b><%=formatnumber(gt_jumlah_orang,0)%></td>
					<td align="right"><b><%=formatnumber(gt_hk_month,1)%></td>
					<td align="right"><b><%=formatnumber(gt_hk,1)%></td>
					<td align="right"><b><%=formatnumber(gt_mkp1,0)%></td>
					<td align="right"><b><%=formatnumber(gt_BASE_SALARY,0)%></td>
					<td align="right"><b><%=formatnumber(gt_OVERTIME,0)%></td>
					<td align="right"><b><%=formatnumber(gt_PREMI,0)%></td>
					<td align="right"><b><%=formatnumber(gt_BORONGAN,0)%></td>
					<td align="right"><b><%=formatnumber(gt_TOTALGAJI,0)%></td>
					<td align="right"><b><%=formatnumber(gt_P_JAMSOSTEK,0)%></td>
					<td align="right"><b><%=formatnumber(gt_P_BPJSKES,0)%></td>
					<td align="right"><b><%=formatnumber(gt_P_PPH21,0)%></td>
					<td align="right"><b><%=formatnumber(gt_PINJAMAN,0)%></td>
					<td align="right"><b><%=formatnumber(gt_DENDA,0)%></td>
					<td align="right"><b><%=formatnumber(gt_totalpotongan,0)%></td>
					<td align="right"><b><%=formatnumber(gt_totalsisaupah,0)%></td> 
					<!--<td align="right"></td>-->
				</tr>
			<%end if%>
<%
			rowgrandtotal = rowgrandtotal + 1
			rsgrandtotal.MoveNext
		Wend
	end if
%>

		
<%else%> <!-- JIKA DETAIL -->
<%

strdtl=	"   SELECT *"  & _
		"     FROM (  SELECT year,"  & _
		"                    month,"  & _
		"                    fcba,"  & _
		"                    employee_api.get_presentgrade2 (fcba, fccode) presentgrade2,"  & _
		"					 employee_hist_api.get_gangcode(year, month, fcba, fccode) gangcode," & _
		"                    SUM (NVL (jumlah_orang, 0)) jumlah_orang,"  & _
		"                    employee_hist_api.get_sectionname (year,"  & _
		"                                                       month,"  & _
		"                                                       fcba,"  & _
		"                                                       fccode)"  & _
		"                       sectionname,"  & _
		"                    SUM (NVL (hk_month, 0)) hk_month,"  & _
		"                    SUM (NVL (hk, 0)) hk,"  & _
		"                    SUM (NVL (mkp1, 0)) mkp1,"  & _
		"                    SUM (NVL (base_salary, 0)) base_salary,"  & _
		"                    SUM (NVL (overtime, 0)) overtime,"  & _
		"                    SUM (NVL (borongan, 0)) borongan,"  & _
		"                    SUM (NVL (p_jamsostek, 0)) p_jamsostek,"  & _
		"                    SUM (NVL (p_bpjskes, 0)) p_bpjskes,"  & _
		"                    SUM (NVL (p_pph21, 0)) p_pph21,"  & _
		"                    SUM (NVL (premi, 0)) premi,"  & _
		"                    SUM (NVL (totalgaji, 0)) totalgaji,"  & _
		"                    SUM (NVL (denda, 0)) denda,"  & _
		"                    SUM (NVL (total, 0)) total,"  & _
		"                    SUM (NVL (p_pinjaman, 0)) pinjaman"  & _
		"               FROM (  SELECT TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)) year,"  & _
		"                              TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)) month,"  & _
		"                              t.fcba,"  & _
		"                              t.fccode,"  & _
		"                              1 jumlah_orang,"  & _
		"                              SUM (get_hk) hk_month,"  & _
		"                              SUM (get_hk_kj) hk,"  & _
		"                              SUM (get_mkp1) mkp1,"  & _
		"                              SUM (t.base_salary) base_salary,"  & _
		"                              SUM (t.rapelgaji) + SUM (t.tunjangan) + SUM (t.premi)"  & _
		"                                 premi,"  & _
		"                              SUM (t.overtime) overtime,"  & _
		"                              SUM (t.borongan) borongan,"  & _
		"                              SUM (t.p_jamsostek) p_jamsostek,"  & _
		"                              SUM (t.p_bpjskes) p_bpjskes,"  & _
		"                              SUM (t.p_pph21) p_pph21,"  & _
		"                              SUM (t.p_pinjaman) p_pinjaman,"  & _
		"                                SUM (t.p_panen)"  & _
		"                              + SUM (t.p_mk_p1)"  & _
		"                              + SUM (t.p_koperasi)"  & _
		"                              + SUM (t.p_lain)"  & _
		"                                 denda,"  & _
		"                                SUM (t.base_salary)"  & _
		"                              + SUM (t.overtime)"  & _
		"                              + SUM (t.borongan)"  & _
		"                              + SUM (t.premi)"  & _
		"                                 totalgaji,"  & _
		"                                SUM (t.base_salary)"  & _
		"                              + SUM (t.overtime)"  & _
		"                              + SUM (t.premi)"  & _
		"                              + SUM (t.rapelgaji)"  & _
		"                              + SUM (t.tunjangan)"  & _
		"                              + SUM (t.borongan)"  & _
		"                              + SUM (t.p_jamsostek)"  & _
		"                              + SUM (t.p_bpjskes)"  & _
		"                              + SUM (t.p_mk_p1)"  & _
		"                              + SUM (t.p_panen)"  & _
		"                              + SUM (t.p_pph21)"  & _
		"                              + SUM (t.p_koperasi)"  & _
		"                              + SUM (t.p_lain)"  & _
		"                                 total"  & _
		"                         FROM view_payrolldetail t"  & _
		"                              INNER JOIN EMPLOYEE emp"  & _
		"                                 ON (emp.fccode = t.fccode AND emp.fcba = t.fcba)"  & _
		"                              INNER JOIN PARAMETER pr"  & _
		"                                 ON (    pr.fcba = emp.fcba"  & _
		"                                     AND pr.fccode = emp.emp_type"  & _
		"                                     AND parcode = 'E0103')"  & _
		"                        WHERE     TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)) = '"& v_txtYear &"'"  & _
		"                              AND TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)) = '"& v_txtMonth &"'"  & _
		"                              AND t.fcba = '"& v_txtFcba &"'"  & _
		"                     GROUP BY TO_NUMBER (SUBSTR (t.yearmonth, 1, 4)),"  & _
		"                              TO_NUMBER (SUBSTR (t.yearmonth, 5, 2)),"  & _
		"                              employee_api.get_presentgrade2 (t.fcba, t.fccode),"  & _
		"                              t.fcba,"  & _
		"                              t.fccode "  & _
		"                     HAVING SUM (t.base_salary)+ SUM (t.overtime)+ SUM (t.borongan) + SUM (t.premi) "  & _
		"                     )"  & _
		"           GROUP BY year,"  & _
		"                    month,"  & _
		"                    fcba,"  & _
		"                    employee_api.get_presentgrade2 (fcba, fccode), "  & _
		"					 employee_hist_api.get_gangcode(year, month, fcba, fccode)," & _
		"                    employee_hist_api.get_sectionname (year,"  & _
		"                                                       month,"  & _
		"                                                       fcba,"  & _
		"                                                       fccode))"  & _
		"    WHERE fcba = '"& v_txtFcba &"' AND totalgaji > 0" 
		
		strdtl = strdtl & " order by sectionname , gangcode asc, PRESENTGRADE2 desc " 
	' response.write strdtl
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
	
	if not rsha.EOF then
		RowsSection = 1
		RowsGang = 1
		while not rsha.EOF
		totalpotongan = cdbl(to_nol(rsha("p_jamsostek"))) + cdbl(to_nol(rsha("p_bpjskes"))) + cdbl(to_nol(rsha("p_pph21"))) + cdbl(to_nol(rsha("pinjaman"))) + cdbl(to_nol(rsha("denda"))) 
		totalsisaupah = cdbl(to_nol(rsha("totalgaji"))) - (-1 * cdbl(to_nol(totalpotongan)) )
%>
			<tr>
<%
	if RowsSection = 1 then 
		sqlRowsSection = " Select Count (sectionname) RowsSection from ( " & strdtl & " ) where sectionname = '" & rsha("sectionname") & "'"
		Set rsRowsSection = Server.CreateObject("ADODB.Recordset")
		rsRowsSection.Open sqlRowsSection ,connstr, adOpenStatic
		
		if not rsRowsSection.EOF then
			tempRowsSection = CDbl(rsRowsSection("RowsSection"))
		else
			tempRowsSection = 0
		end if
		
		sqlRowsSectionGang = " select sum(RowsSectionGang) RowsSectionGang" & _
							 " from ( Select gangcode, Count (gangcode) + 1 RowsSectionGang" & _
							 "        from ( " & strdtl & " )" & _
							 "        where sectionname = '" & rsha("sectionname") & "'" & _
							 "        group by gangcode" & _
							 " )"
		Set rsRowsSectionGang = Server.CreateObject("ADODB.Recordset")
		rsRowsSectionGang.Open sqlRowsSectionGang ,connstr, adOpenStatic
		
		if not rsRowsSectionGang.EOF then
			tempRowsSectionGang = CDbl(rsRowsSectionGang("RowsSectionGang"))
		else
			tempRowsSectionGang = 0
		end if
	end if
	
	if RowsGang = 1 then
		sqlRowsGang = " Select Count (gangcode) RowsGang from ( " & strdtl & " ) where sectionname = '" & rsha("sectionname") & "' and gangcode = '" & rsha("gangcode") & "'"
		'response.write "-->" & sqlRowsGang & "<br/>"
		Set rsRowsGang = Server.CreateObject("ADODB.Recordset")
		rsRowsGang.Open sqlRowsGang ,connstr, adOpenStatic
		if not rsRowsGang.EOF then
			tempRowsGang = CDbl(rsRowsGang("RowsGang"))
		else
			tempRowsGang = 0
		end if
	end if
	

%>
				
				<%if RowsSection = 1 then%>
					<% RowSpanSectionGang = 1 + tempRowsSectionGang%>						
						<td align="center" Rowspan="<%=RowSpanSectionGang%>"><%=rsha("sectionname")%></td>
				<%end if%>
				
				<%if RowsGang = 1 then%>
					<td align="center" Rowspan="<%=tempRowsGang+1 %>"><%=rsha("gangcode")%></td>
				<%end if%>
				<td align="center"><%=rsha("presentgrade2")%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("jumlah_orang")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("hk_month")),1)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("hk")),1)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("mkp1")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("base_salary")),0)%></td> 
					<td align="right"><%=formatnumber( cdbl(rsha("Overtime")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("Premi")),0)%></td> 
					<td align="right"><%=formatnumber( cdbl(rsha("Borongan")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("totalgaji")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("P_JAMSOSTEK")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("P_BPJSKES")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("P_PPH21")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("PINJAMAN")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("DENDA")),0)%></td>
			
					<td align="right"><%=formatnumber( cdbl(totalpotongan),0)%></td>
					<td align="right"><%=formatnumber( cdbl(totalsisaupah),0)%></td>
			</tr>
<%
			if RowsGang = tempRowsGang then
				' st_det_jumlah_orang = cdbl(st_det_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				' st_det_hk_month = cdbl(st_det_hk_month) + cdbl(rsha("hk_month")) 
				' st_det_hk = cdbl(st_det_hk) + cdbl(rsha("hk"))
				' st_det_mkp1 = cdbl(st_det_mkp1) + cdbl(rsha("mkp1"))
				' st_det_Amtbsal = cdbl(st_det_Amtbsal) + cdbl(rsha("base_salary"))
				' st_det_Amtrapelgaji = cdbl(st_det_Amtrapelgaji) + cdbl(rsha("rapelgaji"))
				' st_det_AmtOT = cdbl(st_det_AmtOT) + cdbl(rsha("Overtime"))
				' st_det_Amtpremi = cdbl(st_det_Amtpremi) + cdbl(rsha("Premi")) 
				' st_det_AmtTunj = cdbl(st_det_AmtTunj) + cdbl(rsha("Tunjangan")) 
				' st_det_AmtBorong = cdbl(st_det_AmtBorong) + cdbl(rsha("Borongan")) 
				' st_det_AmtJamsos = cdbl(st_det_AmtJamsos) + cdbl(rsha("P_Jamsostek")) 
				' st_det_Amtbpjskes = cdbl(st_det_Amtbpjskes) + cdbl(rsha("p_bpjskes"))
				' st_det_AmtMKP1 = cdbl(st_det_AmtMKP1) + cdbl(rsha("P_MK_P1")) 
				' st_det_AmtPanen = cdbl(st_det_AmtPanen) + cdbl(rsha("P_Panen")) 
				' st_det_AmtPPH21 = cdbl(st_det_AmtPPH21) + cdbl(rsha("P_PPH21")) 
				' st_det_AmtPKoperasi = cdbl(st_det_AmtPKoperasi) + cdbl(rsha("p_koperasi"))
				' st_det_AmtPlain = cdbl(st_det_AmtPlain) + cdbl(rsha("P_Lain")) 
				' st_det_Amtgrp = cdbl(st_det_Amtgrp) + cdbl(rsha("Total"))
				n_jumlah_orang 	= n_jumlah_orang + cdbl(rsha("jumlah_orang"))
				nhk_month 		= nhk_month 	+ cdbl(rsha("hk_month"))
				nhk		 		= nhk 			+ cdbl(rsha("hk"))
				nmkp1 			= nmkp1		+ cdbl(rsha("mkp1"))
				nBase_Salary 	= nBase_Salary 	+ cdbl(rsha("BASE_SALARY"))
				nOvertime 		= nOvertime 	+ cdbl(rsha("OVERTIME"))
				nBorongan 		= nBorongan 	+ cdbl(rsha("BORONGAN"))
				nP_Jamsostek 	= nP_Jamsostek 	+ cdbl(rsha("P_JAMSOSTEK"))
				nP_bpjskes 		= nP_bpjskes 	+ cdbl(rsha("P_BPJSKES"))
				nTotal 			= nTotal 		+ cdbl(rsha("TOTAL"))
				nDenda 			= nDenda		+ cdbl(rsha("DENDA"))
				nTotalgaji 		= nTotalgaji	+ cdbl(rsha("TOTALGAJI"))
				nPremi 			= nPremi 		+ cdbl(rsha("PREMI"))
				nPinjaman 		= nPinjaman 	+ cdbl(rsha("PINJAMAN"))
				nP_PPH21 		= nP_PPH21 		+ cdbl(rsha("P_PPH21"))
				nTotalPotongan  = nTotalPotongan + cdbl(totalpotongan)
				ntotalsisaupah = ntotalsisaupah + cdbl(totalsisaupah)
				RowsGang = 1
%>
				<tr bgcolor="#F4F4A0">	 
					<td align="center" ><b>Total</td>
					<td align="right"><b><%=formatnumber(n_jumlah_orang,0)%></td>
					<td align="right"><b><%=formatnumber(nhk_month,1)%></td>
					<td align="right"><b><%=formatnumber(nhk,1)%></td>
					<td align="right"><b><%=formatnumber(nmkp1,0)%></td>
					<td align="right"><b><%=formatnumber(nBase_Salary,0)%></td> 
					<td align="right"><b><%=formatnumber(nOvertime,0)%></td>
					<td align="right"><b><%=formatnumber(nPremi,0)%></td>
					<td align="right"><b><%=formatnumber(nBorongan,0)%></td>
					<td align="right"><b><%=formatnumber(nTotalgaji,0)%></td>
					<td align="right"><b><%=formatnumber(nP_Jamsostek,0)%></td>
					<td align="right"><b><%=formatnumber(nP_bpjskes,0)%></td>
					<td align="right"><b><%=formatnumber(nP_PPH21,0)%></td>
					<td align="right"><b><%=formatnumber(nPinjaman,0)%></td>
					<td align="right"><b><%=formatnumber(nDenda,0)%></td>
					<td align="right"><b><%=formatnumber(nTotalPotongan,0)%></td>
					<td align="right"><b><%=formatnumber(ntotalsisaupah,0)%></td> 
				 </tr>	

<%
				' st_det_jumlah_orang = cdbl("0")
				' st_det_hk_month = cdbl("0")
				' st_det_hk = cdbl("0")
				' st_det_mkp1 = cdbl("0")
				' st_det_Amtbsal = cdbl("0") 
				' st_det_Amtrapelgaji = cdbl("0")
				' st_det_AmtOT = cdbl("0") 
				' st_det_Amtpremi = cdbl("0") 
				' st_det_AmtTunj = cdbl("0") 
				' st_det_AmtBorong = cdbl("0") 
				' st_det_AmtJamsos = cdbl("0") 
				' st_det_Amtbpjskes = cdbl("0")
				' st_det_AmtMKP1 = cdbl("0") 
				' st_det_AmtPanen = cdbl("0") 
				' st_det_AmtPPH21 = cdbl("0") 
				' st_det_AmtPKoperasi = cdbl("0")
				' st_det_AmtPlain = cdbl("0") 
				' st_det_Amtgrp = cdbl("0")
				n_jumlah_orang = cdbl("0")
				nhk_month = cdbl("0")
				nhk  = cdbl("0")
				nBase_Salary = cdbl("0")
				nOvertime = cdbl("0")
				nPremi = cdbl("0")
				nTunjangan = cdbl("0")
				nBorongan = cdbl("0")
				nP_Jamsostek = cdbl("0")
				nmkp1 = cdbl("0")
				nP_Panen = cdbl("0")
				nP_Pph21 = cdbl("0")
				nP_Lain = cdbl("0")
				nTotal = cdbl("0")
				nDenda = cdbl("0")
				nTotalgaji = cdbl("0")
				nPinjaman = cdbl("0")
				nTotalPotongan = cdbl("0")
				ntotalsisaupah = cdbl("0")
				nP_bpjskes = cdbl("0")
			else
				' st_det_jumlah_orang = cdbl(st_det_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				' st_det_hk_month = cdbl(st_det_hk_month) + cdbl(rsha("hk_month")) 
				' st_det_hk = cdbl(st_det_hk) + cdbl(rsha("hk"))
				' st_det_mkp1 = cdbl(st_det_mkp1) + cdbl(rsha("mkp1"))
				' st_det_Amtbsal = cdbl(st_det_Amtbsal) + cdbl(rsha("base_salary"))
				' st_det_Amtrapelgaji = cdbl(st_det_Amtrapelgaji) + cdbl(rsha("rapelgaji"))
				' st_det_AmtOT = cdbl(st_det_AmtOT) + cdbl(rsha("Overtime"))
				' st_det_Amtpremi = cdbl(st_det_Amtpremi) + cdbl(rsha("Premi")) 
				' st_det_AmtTunj = cdbl(st_det_AmtTunj) + cdbl(rsha("Tunjangan")) 
				' st_det_AmtBorong = cdbl(st_det_AmtBorong) + cdbl(rsha("Borongan")) 
				' st_det_AmtJamsos = cdbl(st_det_AmtJamsos) + cdbl(rsha("P_Jamsostek")) 
				' st_det_Amtbpjskes = cdbl(st_det_Amtbpjskes) + cdbl(rsha("p_bpjskes"))
				' st_det_AmtMKP1 = cdbl(st_det_AmtMKP1) + cdbl(rsha("P_MK_P1")) 
				' st_det_AmtPanen = cdbl(st_det_AmtPanen) + cdbl(rsha("P_Panen")) 
				' st_det_AmtPPH21 = cdbl(st_det_AmtPPH21) + cdbl(rsha("P_PPH21")) 
				' st_det_AmtPKoperasi = cdbl(st_det_AmtPKoperasi) + cdbl(rsha("p_koperasi"))
				' st_det_AmtPlain = cdbl(st_det_AmtPlain) + cdbl(rsha("P_Lain")) 
				' st_det_Amtgrp = cdbl(st_det_Amtgrp) + cdbl(rsha("Total"))
				n_jumlah_orang 	= n_jumlah_orang + cdbl(rsha("jumlah_orang"))
				nhk_month 		= nhk_month 	+ cdbl(rsha("hk_month"))
				nhk		 		= nhk 			+ cdbl(rsha("hk"))
				nmkp1 			= nmkp1		+ cdbl(rsha("mkp1"))
				nBase_Salary 	= nBase_Salary 	+ cdbl(rsha("BASE_SALARY"))
				nOvertime 		= nOvertime 	+ cdbl(rsha("OVERTIME"))
				nBorongan 		= nBorongan 	+ cdbl(rsha("BORONGAN"))
				nP_Jamsostek 	= nP_Jamsostek 	+ cdbl(rsha("P_JAMSOSTEK"))
				nP_bpjskes 		= nP_bpjskes 	+ cdbl(rsha("P_BPJSKES"))
				nTotal 			= nTotal 		+ cdbl(rsha("TOTAL"))
				nDenda 			= nDenda		+ cdbl(rsha("DENDA"))
				nTotalgaji 		= nTotalgaji	+ cdbl(rsha("TOTALGAJI"))
				nPremi 			= nPremi 		+ cdbl(rsha("PREMI"))
				nPinjaman 		= nPinjaman 	+ cdbl(rsha("PINJAMAN"))
				nP_PPH21 		= nP_PPH21 		+ cdbl(rsha("P_PPH21"))
				nTotalPotongan  = nTotalPotongan + cdbl(totalpotongan)
				ntotalsisaupah = ntotalsisaupah + cdbl(totalsisaupah)
				RowsGang = RowsGang + 1
			end if
%>
<%
			if RowsSection = tempRowsSection then			
				' st_jumlah_orang = cdbl(st_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				' st_hk_month = cdbl(st_hk_month) + cdbl(rsha("hk_month")) 
				' st_hk = cdbl(st_hk) + cdbl(rsha("hk"))
				' st_mkp1 = cdbl(st_mkp1) + cdbl(rsha("mkp1")) 
				' subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary"))
				' subAmtrapelgaji = cdbl(subAmtrapelgaji) + cdbl(rsha("rapelgaji"))
				' subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
				' subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
				' subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
				' subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
				' subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
				' subAmtbpjskes = cdbl(subAmtbpjskes) + cdbl(rsha("p_bpjskes"))
				' subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
				' subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
				' subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
				' subAmtPKoperasi = cdbl(subAmtPKoperasi) + cdbl(rsha("p_koperasi"))
				' subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
				' subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
				v_jumlah_orang 	= v_jumlah_orang + cdbl(rsha("jumlah_orang"))
				vhk_month 		= vhk_month 	+ cdbl(rsha("hk_month"))
				vhk		 		= vhk 			+ cdbl(rsha("hk"))
				vmkp1 			= vmkp1		+ cdbl(rsha("mkp1"))
				vBase_Salary 	= vBase_Salary 	+ cdbl(rsha("BASE_SALARY"))
				vOvertime 		= vOvertime 	+ cdbl(rsha("OVERTIME"))
				vBorongan 		= vBorongan 	+ cdbl(rsha("BORONGAN"))
				vP_Jamsostek 	= vP_Jamsostek 	+ cdbl(rsha("P_JAMSOSTEK"))
				vP_bpjskes 		= vP_bpjskes 	+ cdbl(rsha("P_BPJSKES"))
				vTotal 			= vTotal 		+ cdbl(rsha("TOTAL"))
				vDenda 			= vDenda		+ cdbl(rsha("DENDA"))
				vTotalgaji 		= vTotalgaji	+ cdbl(rsha("TOTALGAJI"))
				vPremi 			= vPremi 		+ cdbl(rsha("PREMI"))
				vPinjaman 		= vPinjaman 	+ cdbl(rsha("PINJAMAN"))
				vP_PPH21 		= vP_PPH21 		+ cdbl(rsha("P_PPH21"))
				vTotalPotongan  = vTotalPotongan + cdbl(totalpotongan)
				vtotalsisaupah = vtotalsisaupah + cdbl(totalsisaupah)
				RowsSection = 1
%>
				<tr bgcolor="#F4F4A0">	 
					<td align="center" colspan="2"><b>Total</b></td>
					<td align="right"><b><%=formatnumber(v_jumlah_orang,0)%></td>
					<td align="right"><b><%=formatnumber(vhk_month,1)%></td>
					<td align="right"><b><%=formatnumber(vhk,1)%></td>
					<td align="right"><b><%=formatnumber(vmkp1,0)%></td>
					<td align="right"><b><%=formatnumber(vBase_Salary,0)%></td> 
					<td align="right"><b><%=formatnumber(vOvertime,0)%></td>
					<td align="right"><b><%=formatnumber(vPremi,0)%></td>
					<td align="right"><b><%=formatnumber(vBorongan,0)%></td>
					<td align="right"><b><%=formatnumber(vTotalgaji,0)%></td>
					<td align="right"><b><%=formatnumber(vP_Jamsostek,0)%></td>
					<td align="right"><b><%=formatnumber(vP_bpjskes,0)%></td>
					<td align="right"><b><%=formatnumber(vP_PPH21,0)%></td>
					<td align="right"><b><%=formatnumber(vPinjaman,0)%></td>
					<td align="right"><b><%=formatnumber(vDenda,0)%></td>
					<td align="right"><b><%=formatnumber(vTotalPotongan,0)%></td>
					<td align="right"><b><%=formatnumber(vtotalsisaupah,0)%></td> 
				 </tr>	

<%
				' st_jumlah_orang = cdbl("0")
				' st_hk_month = cdbl("0")
				' st_hk = cdbl("0")
				' st_mkp1 = cdbl("0")
				' subAmtbsal = cdbl("0") 
				' subAmtrapelgaji = cdbl("0")
				' subAmtOT = cdbl("0") 
				' subAmtpremi = cdbl("0") 
				' subAmtTunj = cdbl("0") 
				' subAmtBorong = cdbl("0") 
				' subAmtJamsos = cdbl("0") 
				' subAmtbpjskes = cdbl("0")
				' subAmtMKP1 = cdbl("0") 
				' subAmtPanen = cdbl("0") 
				' subAmtPPH21 = cdbl("0") 
				' subAmtPKoperasi = cdbl("0")
				' subAmtPlain = cdbl("0") 
				' subAmtgrp = cdbl("0")
				v_jumlah_orang = cdbl("0")
				vhk_month = cdbl("0")
				vhk  = cdbl("0")
				vBase_Salary = cdbl("0")
				vOvertime = cdbl("0")
				vPremi = cdbl("0")
				vTunjangan = cdbl("0")
				vBorongan = cdbl("0")
				vP_Jamsostek = cdbl("0")
				vmkp1 = cdbl("0")
				vP_Panen = cdbl("0")
				vP_Pph21 = cdbl("0")
				vP_Lain = cdbl("0")
				vTotal = cdbl("0")
				vDenda = cdbl("0")
				vTotalgaji = cdbl("0")
				vPinjaman = cdbl("0")
				vTotalPotongan = cdbl("0")
				vtotalsisaupah = cdbl("0")
				vP_bpjskes = cdbl("0")
			else
				' st_jumlah_orang = cdbl(st_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				' st_hk_month = cdbl(st_hk_month) + cdbl(rsha("hk_month"))
				' st_hk = cdbl(st_hk) + cdbl(rsha("hk"))
				' st_mkp1 = cdbl(st_mkp1) + cdbl(rsha("mkp1")) 
				' subAmtbsal = cdbl(subAmtbsal) + cdbl(rsha("base_salary"))
				' subAmtrapelgaji = cdbl(subAmtrapelgaji) + cdbl(rsha("rapelgaji"))
				' subAmtOT = cdbl(subAmtOT) + cdbl(rsha("Overtime"))
				' subAmtpremi = cdbl(subAmtpremi) + cdbl(rsha("Premi")) 
				' subAmtTunj = cdbl(subAmtTunj) + cdbl(rsha("Tunjangan")) 
				' subAmtBorong = cdbl(subAmtBorong) + cdbl(rsha("Borongan")) 
				' subAmtJamsos = cdbl(subAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
				' subAmtbpjskes = cdbl(subAmtbpjskes) + cdbl(rsha("p_bpjskes"))
				' subAmtMKP1 = cdbl(subAmtMKP1) + cdbl(rsha("P_MK_P1")) 
				' subAmtPanen = cdbl(subAmtPanen) + cdbl(rsha("P_Panen")) 
				' subAmtPPH21 = cdbl(subAmtPPH21) + cdbl(rsha("P_PPH21")) 
				' subAmtPKoperasi = cdbl(subAmtPKoperasi) + cdbl(rsha("p_koperasi"))
				' subAmtPlain = cdbl(subAmtPlain) + cdbl(rsha("P_Lain")) 
				' subAmtgrp = cdbl(subAmtgrp) + cdbl(rsha("Total"))
				v_jumlah_orang 	= v_jumlah_orang + cdbl(rsha("jumlah_orang"))
				vhk_month 		= vhk_month 	+ cdbl(rsha("hk_month"))
				vhk		 		= vhk 			+ cdbl(rsha("hk"))
				vmkp1 			= vmkp1		+ cdbl(rsha("mkp1"))
				vBase_Salary 	= vBase_Salary 	+ cdbl(rsha("BASE_SALARY"))
				vOvertime 		= vOvertime 	+ cdbl(rsha("OVERTIME"))
				vBorongan 		= vBorongan 	+ cdbl(rsha("BORONGAN"))
				vP_Jamsostek 	= vP_Jamsostek 	+ cdbl(rsha("P_JAMSOSTEK"))
				vP_bpjskes 		= vP_bpjskes 	+ cdbl(rsha("P_BPJSKES"))
				vTotal 			= vTotal 		+ cdbl(rsha("TOTAL"))
				vDenda 			= vDenda		+ cdbl(rsha("DENDA"))
				vTotalgaji 		= vTotalgaji	+ cdbl(rsha("TOTALGAJI"))
				vPremi 			= vPremi 		+ cdbl(rsha("PREMI"))
				vPinjaman 		= vPinjaman 	+ cdbl(rsha("PINJAMAN"))
				vP_PPH21 		= vP_PPH21 		+ cdbl(rsha("P_PPH21"))
				vTotalPotongan  = vTotalPotongan + cdbl(totalpotongan)
				vtotalsisaupah = vtotalsisaupah + cdbl(totalsisaupah)
				RowsSection = RowsSection + 1
			end if
			
			'Grand Total
			' gt_jumlah_orang = cdbl(gt_jumlah_orang) + cdbl(rsha("jumlah_orang"))
			' gt_hk_month = cdbl(gt_hk_month) + cdbl(rsha("hk_month")) 
			' gt_hk = cdbl(gt_hk) + cdbl(rsha("hk"))
			' gt_mkp1 = cdbl(gt_mkp1) + cdbl(rsha("mkp1")) 
			' grnAmtbsal = cdbl(grnAmtbsal) + cdbl(rsha("base_salary")) 
			' grnAmtrapelgaji = cdbl(grnAmtrapelgaji) + cdbl(rsha("rapelgaji"))
			' grnAmtOT = cdbl(grnAmtOT) + cdbl(rsha("Overtime"))
			' grnAmtpremi = cdbl(grnAmtpremi) + cdbl(rsha("Premi")) 
			' grnAmtTunj = cdbl(grnAmtTunj) + cdbl(rsha("Tunjangan")) 
			' grnAmtBorong = cdbl(grnAmtBorong) + cdbl(rsha("Borongan")) 
			' grnAmtJamsos = cdbl(grnAmtJamsos) + cdbl(rsha("P_Jamsostek")) 
			' grnAmtbpjskes = cdbl(grnAmtbpjskes) + cdbl(rsha("p_bpjskes"))
			' grnAmtMKP1 = cdbl(grnAmtMKP1) + cdbl(rsha("P_MK_P1")) 
			' grnAmtPanen = cdbl(grnAmtPanen) + cdbl(rsha("P_Panen")) 
			' grnAmtPPH21 = cdbl(grnAmtPPH21) + cdbl(rsha("P_PPH21")) 
			' grnAmtPKoperasi = cdbl(grnAmtPKoperasi) + cdbl(rsha("p_koperasi"))
			' grnAmtPlain = cdbl(grnAmtPlain) + cdbl(rsha("P_Lain")) 
			' grnAmtgrp = cdbl(grnAmtgrp) + cdbl(rsha("Total"))
			gt_jumlah_orang = cdbl(gt_jumlah_orang) + cdbl(rsha("jumlah_orang"))
			gt_hk_month = cdbl(gt_hk_month) + cdbl(rsha("hk_month"))
			gt_hk = cdbl(gt_hk) + cdbl(rsha("hk"))
			gt_mkp1 = cdbl(gt_mkp1) + cdbl(rsha("mkp1"))
			gt_BASE_SALARY = cdbl(gt_BASE_SALARY) + cdbl(rsha("BASE_SALARY"))
			gt_OVERTIME = cdbl(gt_OVERTIME) + cdbl(rsha("OVERTIME")) 
			gt_BORONGAN = cdbl(gt_BORONGAN) + cdbl(rsha("BORONGAN"))
			gt_P_JAMSOSTEK = cdbl(gt_P_JAMSOSTEK) + cdbl(rsha("P_JAMSOSTEK"))
			gt_P_BPJSKES = cdbl(gt_P_BPJSKES) + cdbl(rsha("P_BPJSKES")) 
			gt_TOTAL = cdbl(gt_TOTAL) + cdbl(rsha("TOTAL")) 
			gt_DENDA = cdbl(gt_DENDA) + cdbl(rsha("DENDA")) 
			gt_TOTALGAJI = cdbl(gt_TOTALGAJI) + cdbl(rsha("TOTALGAJI")) 
			gt_PREMI = cdbl(gt_PREMI) + cdbl(rsha("PREMI")) 
			gt_PINJAMAN = cdbl(gt_PINJAMAN) + cdbl(rsha("PINJAMAN"))  
			gt_P_PPH21 = cdbl(gt_P_PPH21) + cdbl(rsha("P_PPH21")) 
			gt_totalpotongan = cdbl(gt_totalpotongan) + cdbl(totalpotongan)
			gt_totalsisaupah = cdbl(gt_totalsisaupah) + cdbl(totalsisaupah)
			rsha.movenext
		wend
%>
				<tr>
					<td align="center" colspan="20"></td>
				</tr>
				<tr bgcolor="#FFFF00">
					<td align="center" colspan="3"><b>G.Total</b></td>
					<td align="right"><b><%=formatnumber(gt_jumlah_orang,0)%></td>
					<td align="right"><b><%=formatnumber(gt_hk_month,1)%></td>
					<td align="right"><b><%=formatnumber(gt_hk,1)%></td>
					<td align="right"><b><%=formatnumber(gt_mkp1,0)%></td>
					<td align="right"><b><%=formatnumber(gt_BASE_SALARY,0)%></td>
					<td align="right"><b><%=formatnumber(gt_OVERTIME,0)%></td>
					<td align="right"><b><%=formatnumber(gt_PREMI,0)%></td>
					<td align="right"><b><%=formatnumber(gt_BORONGAN,0)%></td>
					<td align="right"><b><%=formatnumber(gt_TOTALGAJI,0)%></td>
					<td align="right"><b><%=formatnumber(gt_P_JAMSOSTEK,0)%></td>
					<td align="right"><b><%=formatnumber(gt_P_BPJSKES,0)%></td>
					<td align="right"><b><%=formatnumber(gt_P_PPH21,0)%></td>
					<td align="right"><b><%=formatnumber(gt_PINJAMAN,0)%></td>
					<td align="right"><b><%=formatnumber(gt_DENDA,0)%></td>
					<td align="right"><b><%=formatnumber(gt_totalpotongan,0)%></td>
					<td align="right"><b><%=formatnumber(gt_totalsisaupah,0)%></td> 
<%
	end if
%>
<%end if%>
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