<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>

<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<% Dim Judul : Judul = "slip gaji karyawan"%>
<html>
<head>
	<title>ePlaS Reporting - <%=PCase(Judul)%></title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<link rel="stylesheet" href="/css/font/font-awesome.css"/>
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
	<script language="JavaScript" src="/js/jquery.printPage.js"></script> 
	<script>
					function toggle(id){
						var e = document.getElementById(id);
						if(e.style.display == 'inline-block'){
							e.style.display = 'none';
						}else if(e.style.display == 'none'){
							e.style.display = 'inline-block';
						}
					}
	</script>
	<script>
	 function printDiv(divName) {
	 var printContents = document.getElementById(divName).innerHTML;
     var originalContents = document.body.innerHTML;

     document.body.innerHTML = printContents;

     window.print();

     document.body.innerHTML = originalContents;
}
</script>
<style type="text/css">
/* ---------- Halaman slip gaji (tampilan baru) ---------- */
#wrapper { background: #f3f6f5; min-height: 100vh; box-sizing: border-box; }
#header-wrapper { display: block; }
.navi_to_print {
	display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
	margin: 6px auto 14px;
	width: 980px; max-width: 100%;
}
a.button_print {
	display: inline-flex; align-items: center; gap: 8px;
	padding: 8px 16px;
	font-size: 13px; font-weight: 600;
	color: #fff; background: #1B7863;
	border: 0; border-radius: 8px;
	cursor: pointer; text-decoration: none;
}
a.button_print:hover { background: #12584a; color: #fff; text-decoration: none; }
#navi_pagebreak { font-size: 13px; color: #5d6d69; }
#navi_pagebreak input[type="text"] { min-height: 30px; text-align: center; }
.letter {
	box-sizing: border-box;
	width: 980px; max-width: 100%;
	margin: 0 auto 30px;
	padding: 24px;
	background: #fff;
	border: 1px solid #e1e7e5;
	border-radius: 14px;
	box-shadow: 0 8px 24px rgba(23, 38, 42, .06);
}
@media print {
	#header-wrapper, .navi_to_print, .printhide, a.button_print { display: none !important; }
	#wrapper { background: #fff; padding: 0; min-height: 0; }
	.letter { width: 100%; margin: 0; padding: 0; border: 0; border-radius: 0; box-shadow: none; }
	.page-break { display: block; page-break-before: always; }
}
</style>

</head>
<body>

<%
dim x
Dim v_txtMonth, v_txtYear
Dim ScriptName, v_txtGangCode, v_txtStatus, v_txtemployeecode

v_txtMonth		= Request("txtMonth")
v_txtYear		= Request("txtYear")

If Len(Trim(v_txtMonth)) = 0 Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

If Len(Trim(v_txtYear)) = 0 Then
	v_txtYear = Year(Session("v_Costbook"))
End If	


v_txtGangCode	  = Request("txtGangCode")
v_txtStatus		  = Request("txtStatus")
v_txtemployeecode = Request("txtemployeecode")
ScriptName		  = Request.ServerVariables("SCRIPT_NAME")
Dim oRS, oSQL
oSQL	=	" select" & _
			"     fcba" & _
			"     ,YEAR||MONTH yearmonth" & _
			"     ,FCCODE employeecode" & _
			"     ,FCNAME employeename" & _
			"     , gangcode" & _
			"     ,SECTIONNAME division" & _
			"     ,PRESENTGRADE2 presentgrade" & _
			"     ,NVL(GET_HK,0)mandays" & _
			"     ,'' annualleave" & _
			"     ,'' sick" & _
			"     ,'' permission" & _
			"     ,'' absent" & _
			"     ,'' menstruation" & _
			"     ,'' maternity" & _
			"     ,'' overtime" & _
			"     ,NVL(BASE_SALARY,0) basicsalary" & _
			"     ,NVL('',0) premipanen" & _
			"     ,NVL('',0) BONUSPREMIPANEN" & _
			"     ,NVL('',0) premibrondolan" & _
			"     ,NVL(OVERTIME,0) rupiahlembur" & _
			"     ,NVL(BORONGAN,0) rupiahborongan" & _
			"     ,NVL(PREMI,0) allowances" & _
			"     ,NVL(TOTALGAJI,0) grosssalary" & _
			"     ,NVL(P_JAMSOSTEK,0) jamsostek" & _
			"     ,NVL(P_BPJSKES,0) potonganbpjskes" & _
			"     ,NVL(P_PPH21,0)pph" & _
			"     ,NVL('',0) potonganmkp1" & _
			"     ,NVL(DENDA,0) potonganpanen" & _
			"     ,NVL(PINJAMAN,0) potonganpinjamanuang" & _
			"     ,NVL('',0) potonganpinjamanmaterial" & _
			"     ,NVL('',0) potonganalatkerja" & _
			"     ,NVL('',0) potonganpengobatan" & _
			"     ,NVL('',0) potongankoperasi" & _
			"     ,NVL('',0) deductions" & _
			"     ,NVL(TOTALPOTONGAN,0) totaldeductions" & _
			"     ,NVL(TOTAL,0) takehomepay" & _
			"     , BUSINESSUNIT_API.get_FCTYPE(fcba) fctype" & _
			" from VIEW_BKMPAYROLL" & _
			" where fcba = '" & Session("FCBA") & "'" & _
			"       and to_number(month) = " & v_txtMonth & "" & _
			"       and to_number(year) = " & v_txtYear & ""
If  Len(Trim(v_txtGangCode)) = 0 Then
	oSQL	= oSQL & " and gangcode = 'TEST'"
Else
	If  Len(Trim(v_txtGangCode)) > 0 Then
		If v_txtGangCode <> "ALL" Then
			oSQL	= oSQL & " and gangcode = '" & v_txtGangCode & "'"
		End If
	End If
End If
If  Len(Trim(v_txtStatus)) > 0 Then
	oSQL	= oSQL & " and PRESENTGRADE2 = '" & v_txtStatus & "'"
End If
If  Len(Trim(v_txtemployeecode)) > 0 Then
	oSQL	= oSQL & " and fccode = '" & v_txtemployeecode & "'"
End If
			
'response.write oSQL
Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open oSQL, connstr, adOpenStatic



%>
	<div id="wrapper">	
		<div id="header-wrapper">
			<div class="printhide">
				<input type="button" value="&#xf02f;" onClick="window.print()" style="font-family: 'FontAwesome';margin-top:-10px;">
			<div class="export-btn">
        
			 <!--#include file="../print_excel.asp"-->
        </div>
				<!--#include file="menutop.asp"-->
			</div>
			<form name="frmAR" method="post" action="<%=ScriptName%>">
			<div class="printhide">
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">TAHUN</font>
						</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 4
							If len(v_txtYear)= 0 then v_txtYear = Yr
						%>
						<td width="4">
							<select name="txtYear">
							<%For i = MinYr to MaxYr%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">BULAN</font>
						</td>
						<td width="2">
							<select name="txtMonth">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">GANGCODE</font>
						</td>
						<td>
							<select name="txtGangCode">
								<option value = "ALL">--ALL--</option>
								<%
								Dim rsGangCode, strGangCode
								strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE ACTIVATION = 'Y' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
								
								Set rsGangCode = Server.CreateObject("ADODB.Recordset")
								rsGangCode.Open strGangCode, connstr, adOpenStatic 
								
								if not rsGangCode.eof then
								Do While Not rsGangCode.eof
								if rsGangCode("FCCODE") = v_txtGangCode then%>
										<option value="<%=rsGangCode("FCCODE")%>" Selected><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option> 
									<%Else%>
										<option value="<%=rsGangCode("FCCODE")%>"><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option>  
									<%End if
									rsGangCode.movenext
								loop
								end if
								rsGangCode.Close
								set rsGangCode = Nothing
								%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">STATUS</font>
						</td>
						<td>
							<select name="txtStatus">
								<option value = "" <%If Len(Trim(v_txtStatus)) = 0 Then %> Selected <%End If%> >--ALL--</option>
								<option value = "KHL" <%If v_txtStatus = "KHL" Then %> Selected <%End If%> >KHL</option>
								<option value = "KHT" <%If v_txtStatus = "KHT" Then %> Selected <%End If%> >KHT</option>
								<option value = "KBT" <%If v_txtStatus = "KBT" Then %> Selected <%End If%> >KBT</option>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">NIK KARYAWAN</font>
						</td>
						<td align="center">
							<input type="text" maxlength="20" name="txtemployeecode" size="10" value="<%=v_txtemployeecode%>" title="Input Stock code">
						</td>
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>
<div class="navi_to_print">
<a id="open_navi" class="button_print" onclick="javascript:toggle('navi_pagebreak','open_navi')" >
<i class="fa fa-print"></i> Print</a>
<div id="navi_pagebreak" style="display:none; ">&nbsp;&nbsp;
1&nbsp;page&nbsp;:&nbsp;<input maxlength="1" type="text" style="width:30px; text-align:center; "  value="3" disabled />&nbsp;Payslip &nbsp;&nbsp;<input type="button" value="Print" onclick="printDiv('product-table')" />
</div>
</div>
<div class="letter">
<div id="product-table">
<style>
/* Gaya slip ditaruh di dalam #product-table agar ikut tersalin saat printDiv('product-table') */
.slip {
	box-sizing: border-box;
	margin: 0 auto 18px;
	max-width: 900px;
	border: 1px solid #cfdcd8;
	border-radius: 12px;
	overflow: hidden;
	font-family: 'Segoe UI', Roboto, Arial, sans-serif;
	font-size: 12px;
	color: #17262a;
	background: #fff;
	page-break-inside: avoid;
	break-inside: avoid;
}
.slip * { box-sizing: border-box; }
.slip-head {
	display: flex; align-items: center; justify-content: space-between; gap: 12px;
	padding: 10px 16px;
	color: #fff;
	background: #1B7863;
	-webkit-print-color-adjust: exact; print-color-adjust: exact;
}
.slip-ba { font-size: 11px; font-weight: 600; letter-spacing: .4px; opacity: .9; }
.slip-title { font-size: 16px; font-weight: 700; letter-spacing: 1px; }
.slip-period {
	padding: 4px 12px;
	font-size: 12px; font-weight: 700; letter-spacing: .5px;
	color: #17262a; background: #F7BD28;
	border-radius: 999px;
	-webkit-print-color-adjust: exact; print-color-adjust: exact;
}
.slip-emp {
	display: grid; grid-template-columns: 2fr 1fr 1.4fr; gap: 8px 16px;
	padding: 10px 16px;
	background: #f6f9f8;
	border-bottom: 1px solid #e1e7e5;
	-webkit-print-color-adjust: exact; print-color-adjust: exact;
}
.slip-emp .lbl { display: block; font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: .4px; color: #5d6d69; }
.slip-emp b { font-size: 13px; font-weight: 700; }
.slip-cols { display: grid; grid-template-columns: 1fr 1fr; gap: 0; }
.slip-cols > div + div { border-left: 1px solid #e1e7e5; }
#product-table table.slip-tbl { width: 100%; border-collapse: collapse; margin: 0; }
#product-table table.slip-tbl th {
	padding: 7px 16px;
	font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .4px;
	text-align: left; color: #12584a;
	background: #fff; border: 0; border-bottom: 1px solid #e1e7e5;
}
#product-table table.slip-tbl th.num, #product-table table.slip-tbl td.num { text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }
#product-table table.slip-tbl td { padding: 4px 16px; border: 0; font-size: 12px; }
#product-table table.slip-tbl tbody tr:nth-child(even) td { background: #fbfcfc; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
#product-table table.slip-tbl tfoot td {
	padding: 7px 16px;
	font-weight: 700;
	border-top: 1px solid #cfdcd8;
	background: #f6f9f8;
	-webkit-print-color-adjust: exact; print-color-adjust: exact;
}
.slip-net {
	display: flex; align-items: center; justify-content: space-between; gap: 12px;
	padding: 10px 16px;
	background: #e8f3f0;
	border-top: 1px solid #b9d9d0; border-bottom: 1px solid #b9d9d0;
	-webkit-print-color-adjust: exact; print-color-adjust: exact;
}
.slip-net span { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .5px; color: #12584a; }
.slip-net b { font-size: 18px; font-weight: 800; color: #12584a; font-variant-numeric: tabular-nums; }
.slip-att { display: grid; grid-template-columns: repeat(8, 1fr); border-bottom: 1px solid #e1e7e5; }
.slip-att div { padding: 6px 8px; text-align: center; border-right: 1px solid #eef2f1; }
.slip-att div:last-child { border-right: 0; }
.slip-att .lbl { display: block; font-size: 9px; font-weight: 600; text-transform: uppercase; letter-spacing: .3px; color: #5d6d69; }
.slip-att b { font-size: 12px; }
.slip-sign { display: grid; grid-template-columns: 1fr 1fr; }
.slip-sign div { padding: 8px 16px 10px; font-size: 11px; color: #5d6d69; }
.slip-sign div + div { border-left: 1px solid #e1e7e5; }
.slip-sign .line { display: block; margin-top: 34px; padding-top: 4px; border-top: 1px solid #17262a; color: #17262a; font-weight: 600; min-height: 18px; }
.slip-empty {
	padding: 40px 16px;
	text-align: center;
	color: #5d6d69;
	font-family: 'Segoe UI', Roboto, Arial, sans-serif;
	font-size: 14px;
	border: 1px dashed #cfdcd8;
	border-radius: 12px;
}
.slip-empty b { display: block; margin-bottom: 4px; font-size: 16px; color: #17262a; }
.page-break hr { display: block; border: 0; border-top: 1px dashed #c0392b; margin: 6px 0 18px; }
@media print {
	/* Ringkas agar 3 slip muat dalam 1 halaman A4 */
	.slip { margin-bottom: 6px; font-size: 9px; line-height: 1.25; border-radius: 6px; }
	.slip-head { padding: 4px 10px; }
	.slip-ba { font-size: 8px; }
	.slip-title { font-size: 11px; }
	.slip-period { padding: 1px 8px; font-size: 9px; }
	.slip-emp { padding: 3px 10px; gap: 2px 10px; }
	.slip-emp .lbl { font-size: 7px; }
	.slip-emp b { font-size: 10px; }
	#product-table table.slip-tbl th { padding: 2px 10px; font-size: 8px; }
	#product-table table.slip-tbl td { padding: 0 10px; font-size: 9px; line-height: 1.3; }
	#product-table table.slip-tbl tfoot td { padding: 2px 10px; }
	.slip-net { padding: 3px 10px; }
	.slip-net span { font-size: 9px; }
	.slip-net b { font-size: 12px; }
	.slip-att div { padding: 1px 4px; }
	.slip-att .lbl { font-size: 7px; }
	.slip-att b { font-size: 9px; }
	.slip-sign div { padding: 3px 10px 4px; font-size: 8px; }
	.slip-sign .line { margin-top: 14px; padding-top: 2px; min-height: 12px; }
	.page-break hr { display: none; }
}
</style>

<!--[START] Jika data Tidak Ada !!!-->
<%If oRS.BOF and oRS.EOF Then%>
	<div class="slip-empty">
		<b>Belum ada slip gaji untuk ditampilkan</b>
		Pilih Tahun, Bulan, Gangcode / Status / NIK karyawan di atas, lalu klik <b style="display:inline;font-size:inherit;">View</b>.
	</div>
<%End If%>
<!--[END] Jika data Tidak Ada !!!-->

<!--[START] Jika data Ada-->
<%
dim n
n=1
If Not oRS.EOF then
	Do While Not oRS.EOF

%>
<div class="slip">
	<div class="slip-head">
		<div>
			<div class="slip-ba"><%=Session("BA")%></div>
			<div class="slip-title">SLIP GAJI</div>
		</div>
		<div class="slip-period"><%=Ucase(Bulan(v_txtMonth))%>&nbsp;<%=v_txtYear%></div>
	</div>
	<div class="slip-emp">
		<div><span class="lbl">NIK / Nama Karyawan</span><b><%=oRS("employeecode")%> / <%=oRS("employeename")%></b></div>
		<div><span class="lbl">Status</span><b><%=oRS("presentgrade")%></b></div>
		<div><span class="lbl">Divisi</span><b><%=oRS("division")%></b></div>
	</div>
	<div class="slip-cols">
		<div>
			<table class="slip-tbl">
				<thead><tr><th>Pendapatan</th><th class="num">Rp</th></tr></thead>
				<tbody>
					<tr><td>Gaji Pokok</td><td class="num"><%=FormatNumber(oRS("basicsalary"), 0)%></td></tr>
					<tr><td>Premi Panen</td><td class="num"><%=FormatNumber(oRS("premipanen"), 0)%></td></tr>
					<tr><td>Bonus Premi Panen</td><td class="num"><%=FormatNumber(oRS("bonuspremipanen"), 0)%></td></tr>
					<tr><td>Premi Brondolan</td><td class="num"><%=FormatNumber(oRS("premibrondolan"), 0)%></td></tr>
					<tr><td>Borongan</td><td class="num"><%=FormatNumber(oRS("rupiahborongan"), 0)%></td></tr>
					<tr><td>Lembur</td><td class="num"><%=FormatNumber(oRS("rupiahlembur"), 0)%></td></tr>
					<tr><td>Lain - Lain</td><td class="num"><%=FormatNumber(oRS("allowances"), 0)%></td></tr>
					<tr class="fill"><td>&nbsp;</td><td></td></tr>
					<tr class="fill"><td>&nbsp;</td><td></td></tr>
					<tr class="fill"><td>&nbsp;</td><td></td></tr>
					<tr class="fill"><td>&nbsp;</td><td></td></tr>
				</tbody>
				<tfoot><tr><td>Gaji Bruto</td><td class="num"><%=FormatNumber(oRS("grosssalary"), 0)%></td></tr></tfoot>
			</table>
		</div>
		<div>
			<table class="slip-tbl">
				<thead><tr><th>Potongan</th><th class="num">Rp</th></tr></thead>
				<tbody>
					<tr><td>BPJS Ketenagakerjaan</td><td class="num"><%=FormatNumber(oRS("jamsostek"), 0)%></td></tr>
					<tr><td>BPJS Kesehatan</td><td class="num"><%=FormatNumber(oRS("potonganbpjskes"), 0)%></td></tr>
					<tr><td>PPh</td><td class="num"><%=FormatNumber(oRS("pph"), 0)%></td></tr>
					<tr><td>Potongan MK/P1</td><td class="num"><%=FormatNumber(oRS("potonganmkp1"), 0)%></td></tr>
					<tr><td>Potongan Panen</td><td class="num"><%=FormatNumber(oRS("potonganpanen"), 0)%></td></tr>
					<tr><td>Potongan Pinjaman Uang</td><td class="num"><%=FormatNumber(oRS("potonganpinjamanuang"), 0)%></td></tr>
					<tr><td>Potongan Pinjaman Material</td><td class="num"><%=FormatNumber(oRS("potonganpinjamanmaterial"), 0)%></td></tr>
					<tr><td>Potongan Alat Kerja</td><td class="num"><%=FormatNumber(oRS("potonganalatkerja"), 0)%></td></tr>
					<tr><td>Potongan Pengobatan DLT</td><td class="num"><%=FormatNumber(oRS("potonganpengobatan"), 0)%></td></tr>
					<tr><td>Potongan Koperasi</td><td class="num"><%=FormatNumber(oRS("potongankoperasi"), 0)%></td></tr>
					<tr><td>Potongan Lain - Lain</td><td class="num"><%=FormatNumber(oRS("deductions"), 0)%></td></tr>
				</tbody>
				<tfoot><tr><td>Total Potongan</td><td class="num"><%=FormatNumber(oRS("totaldeductions"), 0)%></td></tr></tfoot>
			</table>
		</div>
	</div>
	<div class="slip-net">
		<span>Gaji Netto (Take Home Pay)</span>
		<b>Rp <%=FormatNumber(oRS("takehomepay"), 0)%></b>
	</div>
	<div class="slip-att">
		<div><span class="lbl">HK</span><b><%=oRS("mandays")%></b></div>
		<div><span class="lbl">Cuti</span><b><%=oRS("annualleave")%></b></div>
		<div><span class="lbl">Sakit</span><b><%=oRS("sick")%></b></div>
		<div><span class="lbl">Izin</span><b><%=oRS("permission")%></b></div>
		<div><span class="lbl">MK/P1</span><b><%=oRS("absent")%></b></div>
		<div><span class="lbl">Cuti Haid</span><b><%=oRS("menstruation")%></b></div>
		<div><span class="lbl">Cuti Melahirkan</span><b><%=oRS("maternity")%></b></div>
		<div><span class="lbl">Lembur</span><b><%=oRS("overtime")%></b></div>
	</div>
	<div class="slip-sign">
		<div>Diberikan oleh<span class="line">&nbsp;</span></div>
		<div>Diterima oleh<span class="line"><%=oRS("employeename")%></span></div>
	</div>
</div>

<%
x=3
if  n mod x = 0 then%>
	<div class="page-break"><hr/></div>
<%else

end if
 n = n + 1
		oRS.MoveNext
	Loop
End If

oRS.Close
set oRS = Nothing
%>


<!--[End] Jika data Ada-->
		</div>
	</div>
</div>

</body>
</html>