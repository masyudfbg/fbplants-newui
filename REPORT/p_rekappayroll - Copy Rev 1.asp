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
		<th rowspan="2">Rapel Gaji</th>
		<th rowspan="2">Lembur</th>
		<th rowspan="2">Premi</th>
		<th rowspan="2">Tunjangan</th>
		<th rowspan="2">Borongan</th>		
		<th colspan="7">Potongan</th>		
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
<%if len(trim(v_txttampilkan)) = 0 then%> <!-- JIKA REKAP -->
	<%

			strdtl = " select aa.fcba" & _
                     "        , aa.sectioncode sectionname" & _
					 "        , aa.ordersec" & _
					 "        , aa.gradecode presentgrade" & _
					 "        , nvl(bb.jumlah_orang, 0) jumlah_orang" & _
					 "        , nvl(bb.hk_month, 0) hk_month" & _
					 "        , nvl(bb.hk, 0) hk" & _
					 "        , nvl(bb.mkp1, 0) mkp1" & _
					 "        , nvl(bb.base_salary, 0) base_salary" & _
                     "        , nvl(bb.rapelgaji, 0) rapelgaji" & _
                     "        , nvl(bb.overtime, 0) overtime" & _
                     "        , nvl(bb.premi, 0) premi" & _
                     "        , nvl(bb.tunjangan, 0) tunjangan" & _
                     "        , nvl(bb.borongan, 0) borongan" & _
                     "        , nvl(bb.p_jamsostek, 0) p_jamsostek" & _
                     "        , nvl(bb.p_bpjskes, 0) p_bpjskes" & _
                     "        , nvl(bb.p_mk_p1, 0) p_mk_p1" & _
                     "        , nvl(bb.p_panen, 0) p_panen" & _
                     "        , nvl(bb.p_pph21, 0) p_pph21" & _
                     "        , nvl(bb.p_koperasi, 0) p_koperasi" & _
                     "        , nvl(bb.p_lain, 0) p_lain" & _
                     "        , nvl(bb.total, 0) total" & _
                     " from ( select bu.fccode fcba" & _
                     "               , sc.fccode sectioncode" & _
					 "               , sc.ordersec" & _
                     "               , gr.fccode gradecode" & _
                     "        from employee_grade gr" & _
                     "             , businessunit bu" & _
                     "             , (select fcba, fccode, sectionname_api.get_orderdata(fccode) ordersec" & _
					 "                from (select distinct fcba, sectionname fccode" & _
                     "                      from view_payrolldetail t" & _
                     "                      where to_number(substr(t.yearmonth,1,4)) = " & v_txtYear & "" & _
                     "                            and to_number(substr(t.yearmonth,5,2)) = " & v_txtMonth & "" & _
                     "                            and t.fcba = '" & v_txtFcba & "')) sc" & _
                     "        where bu.fccode = sc.fcba" & _
                     "              and bu.fccode = '" & v_txtFcba & "'" & _
                     " ) aa left join ( select bb.fcba" & _
                     "                         , bb.yearmonth" & _
                     "                         , employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode) sectioncode" & _
                     "                         , sectionname_api.get_fcname(employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode)) sectionname" & _
                     "                         , employee_hist_api.get_presentgrade(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode) presentgrade" & _
                     "                         , sum(bb.jumlah_orang) jumlah_orang" & _
					 "                         , sum(hk_month) hk_month" & _
					 "                         , sum(hk) hk" & _
					 "                         , sum(mkp1) mkp1" & _
					 "                         , sum(base_salary) base_salary" & _
                     "                         , sum(rapelgaji) rapelgaji" & _
                     "                         , sum(overtime) overtime" & _
                     "                         , sum(premi) premi" & _
                     "                         , sum(tunjangan) tunjangan" & _
                     "                         , sum(borongan) borongan" & _
                     "                         , sum(p_jamsostek) p_jamsostek" & _
                     "                         , sum(p_bpjskes) p_bpjskes" & _
                     "                         , sum(p_mk_p1) p_mk_p1" & _
                     "                         , sum(p_panen) p_panen" & _
                     "                         , sum(p_pph21) p_pph21" & _
                     "                         , sum(p_koperasi) p_koperasi" & _
                     "                         , sum(p_lain) p_lain" & _
                     "                         , sum(total) total" & _
                     "                  from ( select t.yearmonth" & _
                     "                                , t.fcba" & _
					 "                                , t.fccode" & _
					 "                                , 1 jumlah_orang" & _
					 "                                , sum(t.get_hk) hk_month" & _
					 "                                , sum(t.get_hk_kj) hk" & _
					 "                                , sum(t.get_mk + get_p1) mkp1" & _
                     "                                , sum(t.base_salary) base_salary" & _
                     "                                , sum(t.rapelgaji) rapelgaji" & _
                     "                                , sum(t.overtime) overtime" & _
                     "                                , sum(t.premi) premi" & _
                     "                                , sum(t.tunjangan) tunjangan" & _
                     "                                , sum(t.borongan) borongan" & _
                     "                                , sum(t.p_jamsostek) p_jamsostek" & _
                     "                                , sum(t.p_bpjskes) p_bpjskes" & _
                     "                                , sum(t.p_mk_p1) p_mk_p1" & _
                     "                                , sum(t.p_panen) p_panen" & _
                     "                                , sum(t.p_pph21) p_pph21" & _
                     "                                , sum(t.p_koperasi) p_koperasi" & _
                     "                                , sum(t.p_lain) p_lain" & _
                     "                                , sum(t.base_salary" & _
                     "                                      + t.rapelgaji" & _
                     "                                      + t.overtime" & _
                     "                                      + t.premi" & _
                     "                                      + t.tunjangan" & _
                     "                                      + t.borongan" & _
                     "                                      + t.p_jamsostek" & _
                     "                                      + t.p_bpjskes" & _
                     "                                      + t.p_mk_p1" & _
                     "                                      + t.p_panen" & _
                     "                                      + t.p_pph21" & _
                     "                                      + t.p_koperasi" & _
                     "                                      + t.p_lain) total " & _
                     "                         from view_payrolldetail t" & _
                     "                         where to_number(substr(t.yearmonth,1,4)) = " & v_txtYear & "" & _
                     "                               and to_number(substr(t.yearmonth,5,2)) = " & v_txtMonth & "" & _
                     "                               and t.fcba = '" & v_txtFcba & "'" & _
                     "                         group by t.yearmonth" & _
                     "                                  , t.fcba" & _
					 "                                  , t.fccode" & _
                     "                  ) bb" & _
                     "                  group by bb.fcba" & _
                     "                           , bb.yearmonth" & _
                     "                           , employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode)" & _
                     "                           , sectionname_api.get_fcname(employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode))" & _
                     "                           , employee_hist_api.get_presentgrade(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode)" & _
                     " ) bb on (aa.fcba = bb.fcba and aa.gradecode = bb.presentgrade and aa.sectioncode = bb.sectioncode)" & _
                     " order by aa.fcba" & _
                     "          , aa.ordersec" & _
                     "          , employee_grade_api.get_orderdata(aa.gradecode)"
	'response.write strdtl
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    
	  

	if Not rsha.eof then
		rows = 1
		while Not rsha.eof 
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
					<td align="center"><%=rsha("presentgrade")%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("jumlah_orang")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("hk_month")),1)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("hk")),1)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("mkp1")),0)%></td>
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
				<%if rows = 1 then%>
					<td align="center" Rowspan="<%=temprows+1 %>"></td>
				<%end if%>
				  </tr>
<%
			if rows = temprows then
				' cari sub total
				st_jumlah_orang = cdbl(st_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				st_hk_month = cdbl(st_hk_month) + cdbl(rsha("hk_month")) 
				st_hk = cdbl(st_hk) + cdbl(rsha("hk"))
				st_mkp1 = cdbl(st_mkp1) + cdbl(rsha("mkp1")) 
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
				rows = 1
%>
				  <tr bgcolor="#F4F4A0">	 
					<td align="center" ><b>Total</td>
					<td align="right"><b><%=formatnumber(st_jumlah_orang,0)%></td>
					<td align="right"><b><%=formatnumber(st_hk_month,1)%></td>
					<td align="right"><b><%=formatnumber(st_hk,1)%></td>
					<td align="right"><b><%=formatnumber(st_mkp1,0)%></td>
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
					<td align="center" colspan="21" ></td>
				 </tr>

<%
					 st_jumlah_orang = cdbl("0")
					 st_hk_month = cdbl("0")
					 st_hk = cdbl("0")
					 st_mkp1 = cdbl("0")
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
			else
				' cari sub total
					st_jumlah_orang = cdbl(st_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
					st_hk_month = cdbl(st_hk_month) + cdbl(rsha("hk_month"))
					st_hk = cdbl(st_hk) + cdbl(rsha("hk"))
					st_mkp1 = cdbl(st_mkp1) + cdbl(rsha("mkp1")) 
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
					rows = rows + 1
			end if
							  
			' cari grand total 
				gt_jumlah_orang = cdbl(gt_jumlah_orang) + cdbl(rsha("jumlah_orang"))
				gt_hk_month = cdbl(gt_hk_month) + cdbl(rsha("hk_month"))
				gt_hk = cdbl(gt_hk) + cdbl(rsha("hk"))
				gt_mkp1 = cdbl(gt_mkp1) + cdbl(rsha("mkp1"))
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
					  
			
			 rsha.MoveNext
		wend
	end if
	%>	

<%
Dim rsgrandtotal, strgrandtotal, rowgrandtotal
strgrandtotal = " select fcba, presentgrade" & _
                "        , sum(base_salary) base_salary" & _
				"        , sum(hk_month) hk_month" & _
				"        , sum(hk) hk" & _
				"        , sum(mkp1) mkp1" & _
				"        , sum(jumlah_orang) jumlah_orang" & _
                "        , sum(rapelgaji) rapelgaji" & _
                "        , sum(overtime) overtime" & _
                "        , sum(premi) premi" & _
                "        , sum(tunjangan) tunjangan" & _
                "        , sum(borongan) borongan" & _
                "        , sum(p_jamsostek) p_jamsostek" & _
                "        , sum(p_bpjskes) p_bpjskes" & _
                "        , sum(p_mk_p1) p_mk_p1" & _
                "        , sum(p_panen) p_panen" & _
                "        , sum(p_pph21) p_pph21" & _
                "        , sum(p_koperasi) p_koperasi" & _
                "        , sum(p_lain) p_lain" & _
                "        , sum(total) total" & _
				" from ( " & strdtl & " )" & _
				" group by fcba, presentgrade" & _
				" order by employee_grade_api.get_orderdata(presentgrade)"
'response.write strgrandtotal
Set rsgrandtotal = Server.CreateObject("ADODB.Recordset")
rsgrandtotal.Open strgrandtotal, connstr, adOpenStatic
	if Not rsgrandtotal.EOF then
		rowgrandtotal = 1
		While Not rsgrandtotal.EOF
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
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("rapelgaji")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("Overtime")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("Premi")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("Tunjangan")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("Borongan")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("P_Jamsostek")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("p_bpjskes")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("P_MK_P1")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("P_Panen")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("P_PPH21")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("p_koperasi")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("P_Lain")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsgrandtotal("total")),0)%></td>
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

strdtl = " select bb.fcba" & _
         "        , bb.yearmonth" & _
         "        , employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode) sectioncode" & _
         "        , sectionname_api.get_fcname(employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode)) sectionname" & _
		 "        , sectionname_api.get_orderdata(employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode)) ordersec" & _
         "        , employee_hist_api.get_gangcode(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode) gangcode" & _
		 "        , employee_hist_api.get_presentgrade(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode) presentgrade" & _
         "        , sum(bb.jumlah_orang) jumlah_orang" & _
         "        , sum(hk_month) hk_month" & _
		 "        , sum(hk) hk" & _
         "        , sum(mkp1) mkp1" & _
         "        , sum(base_salary) base_salary" & _
         "        , sum(rapelgaji) rapelgaji" & _
         "        , sum(overtime) overtime" & _
         "        , sum(premi) premi" & _
         "        , sum(tunjangan) tunjangan" & _
         "        , sum(borongan) borongan" & _
         "        , sum(p_jamsostek) p_jamsostek" & _
         "        , sum(p_bpjskes) p_bpjskes" & _
         "        , sum(p_mk_p1) p_mk_p1" & _
         "        , sum(p_panen) p_panen" & _
         "        , sum(p_pph21) p_pph21" & _
         "        , sum(p_koperasi) p_koperasi" & _
         "        , sum(p_lain) p_lain" & _
         "        , sum(total) total" & _
         " from ( select t.yearmonth" & _
         "               , t.fcba" & _
         "               , t.fccode" & _
         "               , 1 jumlah_orang" & _
         "               , sum(t.get_hk) hk_month" & _
		 "               , sum(t.get_hk_kj) hk" & _
         "               , sum(t.get_mk + get_p1) mkp1" & _
         "               , sum(t.base_salary) base_salary" & _
         "               , sum(t.rapelgaji) rapelgaji" & _
         "               , sum(t.overtime) overtime" & _
         "               , sum(t.premi) premi" & _
         "               , sum(t.tunjangan) tunjangan" & _
         "               , sum(t.borongan) borongan" & _
         "               , sum(t.p_jamsostek) p_jamsostek" & _
         "               , sum(t.p_bpjskes) p_bpjskes" & _
         "               , sum(t.p_mk_p1) p_mk_p1" & _
         "               , sum(t.p_panen) p_panen" & _
         "               , sum(t.p_pph21) p_pph21" & _
         "               , sum(t.p_koperasi) p_koperasi" & _
         "               , sum(t.p_lain) p_lain" & _
         "               , sum(t.base_salary" & _
         "                 + t.rapelgaji" & _
         "                 + t.overtime" & _
         "                 + t.premi" & _
         "                 + t.tunjangan" & _
         "                 + t.borongan" & _
         "                 + t.p_jamsostek" & _
         "                 + t.p_bpjskes" & _
         "                 + t.p_mk_p1" & _
         "                 + t.p_panen" & _
         "                 + t.p_pph21" & _
         "                 + t.p_koperasi" & _
         "                 + t.p_lain) total " & _
         "        from view_payrolldetail t" & _
         "        where to_number(substr(t.yearmonth,1,4)) = " & v_txtYear & "" & _
         "              and to_number(substr(t.yearmonth,5,2)) = " & v_txtMonth & "" & _
         "              and t.fcba = '" & v_txtFcba & "'" & _
         "        group by t.yearmonth" & _
         "                 , t.fcba" & _
         "                 , t.fccode" & _
		 "                  ) bb" & _
		 " group by bb.fcba" & _
		 "          , bb.yearmonth" & _
         "          , employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode)" & _
         "          , sectionname_api.get_fcname(employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode))" & _
		 "          , sectionname_api.get_orderdata(employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode))" & _
         "          , employee_hist_api.get_gangcode(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode)" & _
		 "          , employee_hist_api.get_presentgrade(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode)" & _
		 " order by bb.fcba" & _
		 "         , bb.yearmonth" & _
		 "         , sectionname_api.get_orderdata(employee_hist_api.get_sectionname(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode))" & _
		 "         , employee_hist_api.get_gangcode(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode)" & _
		 "         , employee_grade_api.get_orderdata(employee_hist_api.get_presentgrade(to_number(substr(bb.yearmonth, 1, 4)), to_number(substr(bb.yearmonth, 5, 2)), bb.fcba, bb.fccode))"
                   
	'response.write strdtl
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
	
	if not rsha.EOF then
		RowsSection = 1
		RowsGang = 1
		while not rsha.EOF
%>
			<tr>
<%
	if RowsSection = 1 then 
		sqlRowsSection = " Select Count (sectioncode) RowsSection from ( " & strdtl & " ) where sectioncode = '" & rsha("sectioncode") & "'"
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
							 "        where sectioncode = '" & rsha("sectioncode") & "'" & _
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
		sqlRowsGang = " Select Count (gangcode) RowsGang from ( " & strdtl & " ) where sectioncode = '" & rsha("sectioncode") & "' and gangcode = '" & rsha("gangcode") & "'"
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
						<td align="center" Rowspan="<%=RowSpanSectionGang%>"><%=rsha("sectioncode")%></td>
				<%end if%>
				
				<%if RowsGang = 1 then%>
					<td align="center" Rowspan="<%=tempRowsGang+1 %>"><%=rsha("gangcode")%></td>
				<%end if%>
				<td align="center"><%=rsha("presentgrade")%></td>
				<td align="right"><%=formatnumber( cdbl(rsha("jumlah_orang")),0)%></td>
				<td align="right"><%=formatnumber( cdbl(rsha("hk_month")),1)%></td>
				<td align="right"><%=formatnumber( cdbl(rsha("hk")),1)%></td>
				<td align="right"><%=formatnumber( cdbl(rsha("mkp1")),0)%></td>
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
			if RowsGang = tempRowsGang then
				st_det_jumlah_orang = cdbl(st_det_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				st_det_hk_month = cdbl(st_det_hk_month) + cdbl(rsha("hk_month")) 
				st_det_hk = cdbl(st_det_hk) + cdbl(rsha("hk"))
				st_det_mkp1 = cdbl(st_det_mkp1) + cdbl(rsha("mkp1"))
				st_det_Amtbsal = cdbl(st_det_Amtbsal) + cdbl(rsha("base_salary"))
				st_det_Amtrapelgaji = cdbl(st_det_Amtrapelgaji) + cdbl(rsha("rapelgaji"))
				st_det_AmtOT = cdbl(st_det_AmtOT) + cdbl(rsha("Overtime"))
				st_det_Amtpremi = cdbl(st_det_Amtpremi) + cdbl(rsha("Premi")) 
				st_det_AmtTunj = cdbl(st_det_AmtTunj) + cdbl(rsha("Tunjangan")) 
				st_det_AmtBorong = cdbl(st_det_AmtBorong) + cdbl(rsha("Borongan")) 
				st_det_AmtJamsos = cdbl(st_det_AmtJamsos) + cdbl(rsha("P_Jamsostek")) 
				st_det_Amtbpjskes = cdbl(st_det_Amtbpjskes) + cdbl(rsha("p_bpjskes"))
				st_det_AmtMKP1 = cdbl(st_det_AmtMKP1) + cdbl(rsha("P_MK_P1")) 
				st_det_AmtPanen = cdbl(st_det_AmtPanen) + cdbl(rsha("P_Panen")) 
				st_det_AmtPPH21 = cdbl(st_det_AmtPPH21) + cdbl(rsha("P_PPH21")) 
				st_det_AmtPKoperasi = cdbl(st_det_AmtPKoperasi) + cdbl(rsha("p_koperasi"))
				st_det_AmtPlain = cdbl(st_det_AmtPlain) + cdbl(rsha("P_Lain")) 
				st_det_Amtgrp = cdbl(st_det_Amtgrp) + cdbl(rsha("Total"))
				RowsGang = 1
%>
				<tr bgcolor="#F4F4A0">
					<td align="center"><b>Total</b></td>
					<td align="right"><b><%=formatnumber(st_det_jumlah_orang,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_hk_month,1)%></td>
					<td align="right"><b><%=formatnumber(st_det_hk,1)%></td>
					<td align="right"><b><%=formatnumber(st_det_mkp1,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_Amtbsal,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_Amtrapelgaji,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_AmtOT,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_Amtpremi,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_AmtTunj,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_AmtBorong,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_AmtJamsos,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_Amtbpjskes,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_AmtMKP1,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_AmtPanen,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_AmtPPH21,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_AmtPKoperasi,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_AmtPlain,0)%></td>
					<td align="right"><b><%=formatnumber(st_det_Amtgrp,0)%></td>
				</tr>

<%
				st_det_jumlah_orang = cdbl("0")
				st_det_hk_month = cdbl("0")
				st_det_hk = cdbl("0")
				st_det_mkp1 = cdbl("0")
				st_det_Amtbsal = cdbl("0") 
				st_det_Amtrapelgaji = cdbl("0")
				st_det_AmtOT = cdbl("0") 
				st_det_Amtpremi = cdbl("0") 
				st_det_AmtTunj = cdbl("0") 
				st_det_AmtBorong = cdbl("0") 
				st_det_AmtJamsos = cdbl("0") 
				st_det_Amtbpjskes = cdbl("0")
				st_det_AmtMKP1 = cdbl("0") 
				st_det_AmtPanen = cdbl("0") 
				st_det_AmtPPH21 = cdbl("0") 
				st_det_AmtPKoperasi = cdbl("0")
				st_det_AmtPlain = cdbl("0") 
				st_det_Amtgrp = cdbl("0")
			else
				st_det_jumlah_orang = cdbl(st_det_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				st_det_hk_month = cdbl(st_det_hk_month) + cdbl(rsha("hk_month")) 
				st_det_hk = cdbl(st_det_hk) + cdbl(rsha("hk"))
				st_det_mkp1 = cdbl(st_det_mkp1) + cdbl(rsha("mkp1"))
				st_det_Amtbsal = cdbl(st_det_Amtbsal) + cdbl(rsha("base_salary"))
				st_det_Amtrapelgaji = cdbl(st_det_Amtrapelgaji) + cdbl(rsha("rapelgaji"))
				st_det_AmtOT = cdbl(st_det_AmtOT) + cdbl(rsha("Overtime"))
				st_det_Amtpremi = cdbl(st_det_Amtpremi) + cdbl(rsha("Premi")) 
				st_det_AmtTunj = cdbl(st_det_AmtTunj) + cdbl(rsha("Tunjangan")) 
				st_det_AmtBorong = cdbl(st_det_AmtBorong) + cdbl(rsha("Borongan")) 
				st_det_AmtJamsos = cdbl(st_det_AmtJamsos) + cdbl(rsha("P_Jamsostek")) 
				st_det_Amtbpjskes = cdbl(st_det_Amtbpjskes) + cdbl(rsha("p_bpjskes"))
				st_det_AmtMKP1 = cdbl(st_det_AmtMKP1) + cdbl(rsha("P_MK_P1")) 
				st_det_AmtPanen = cdbl(st_det_AmtPanen) + cdbl(rsha("P_Panen")) 
				st_det_AmtPPH21 = cdbl(st_det_AmtPPH21) + cdbl(rsha("P_PPH21")) 
				st_det_AmtPKoperasi = cdbl(st_det_AmtPKoperasi) + cdbl(rsha("p_koperasi"))
				st_det_AmtPlain = cdbl(st_det_AmtPlain) + cdbl(rsha("P_Lain")) 
				st_det_Amtgrp = cdbl(st_det_Amtgrp) + cdbl(rsha("Total"))
				RowsGang = RowsGang + 1
			end if
%>
<%
			if RowsSection = tempRowsSection then			
				st_jumlah_orang = cdbl(st_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				st_hk_month = cdbl(st_hk_month) + cdbl(rsha("hk_month")) 
				st_hk = cdbl(st_hk) + cdbl(rsha("hk"))
				st_mkp1 = cdbl(st_mkp1) + cdbl(rsha("mkp1")) 
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
				RowsSection = 1
%>
				<tr bgcolor="#FFFF00">
					<td align="center" colspan="2"><b>Total</b></td>
					<td align="right"><b><%=formatnumber(st_jumlah_orang,0)%></td>
					<td align="right"><b><%=formatnumber(st_hk_month,1)%></td>
					<td align="right"><b><%=formatnumber(st_hk,1)%></td>
					<td align="right"><b><%=formatnumber(st_mkp1,0)%></td>
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

<%
				st_jumlah_orang = cdbl("0")
				st_hk_month = cdbl("0")
				st_hk = cdbl("0")
				st_mkp1 = cdbl("0")
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
			else
				st_jumlah_orang = cdbl(st_jumlah_orang) + cdbl(rsha("jumlah_orang")) 
				st_hk_month = cdbl(st_hk_month) + cdbl(rsha("hk_month"))
				st_hk = cdbl(st_hk) + cdbl(rsha("hk"))
				st_mkp1 = cdbl(st_mkp1) + cdbl(rsha("mkp1")) 
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
				RowsSection = RowsSection + 1
			end if
			
			'Grand Total
			gt_jumlah_orang = cdbl(gt_jumlah_orang) + cdbl(rsha("jumlah_orang"))
			gt_hk_month = cdbl(gt_hk_month) + cdbl(rsha("hk_month")) 
			gt_hk = cdbl(gt_hk) + cdbl(rsha("hk"))
			gt_mkp1 = cdbl(gt_mkp1) + cdbl(rsha("mkp1")) 
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
			
			rsha.movenext
		wend
%>
				<tr>
					<td align="center" colspan="21"></td>
				</tr>
				<tr bgcolor="#FFFF00;">
					<td align="center" colspan="3"><b>Grand Total</b></td>
					<td align="right"><b><%=formatnumber(gt_jumlah_orang,0)%></td>
					<td align="right"><b><%=formatnumber(gt_hk_month,1)%></td>
					<td align="right"><b><%=formatnumber(gt_hk,1)%></td>
					<td align="right"><b><%=formatnumber(gt_mkp1,0)%></td>
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
				</tr>
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