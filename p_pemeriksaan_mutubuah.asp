<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<title>ePlaS Reporting - Pemeriksaan Mutu Buah</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

' v_txtCUSTOMERCODE= trim(request("periodyear"))
division =request("division")
transdate_From =request("transdate_From")
transdate_To =request("transdate_To")
pemeriksa=request("pemeriksa")
collectionpoint=request("collectionpoint")
fieldcode=request("fieldcode")
empcode = Request("empcode")
txtlineno = request("txtlineno")
supervision =request("supervision")
gangcode =request("gangcode")
janjangperiksa_p=request("janjangperiksa_p")
buahnormal_p=request("buahnormal_p")
buahmentah_p=request("buahmentah_p")
buahbusuk_p = Request("buahbusuk_p")
overripe_p = request("overripe_p")
buahabnormal_p=request("buahabnormal_p")
tangkalpanjang_p=request("tangkalpanjang_p")
alasbrondolan_p = Request("alasbrondolan_p")
janjangperiksa_kb=request("janjangperiksa_kb")
buahnormal_kb=request("buahnormal_kb")
buahmentah_kb=request("buahmentah_kb")
buahbusuk_kb = Request("buahbusuk_kb")
overripe_kb = request("overripe_kb")
buahabnormal_kb=request("buahabnormal_kb")
tangkalpanjang_kb=request("tangkalpanjang_kb")
alasbrondolan_kb = Request("alasbrondolan_kb")

Dim rsVH, intRecordsPerPage, cColor
txtSQL = " select pm.transdate, pm.fieldcode, pm.collectionpoint, pm.pemeriksa, pm.janjangperiksa_p, pm.buahnormal_p, " & _
		 " pm.buahmentah_p, pm.buahbusuk_p, pm.overripe_p, pm.buahabnormal_p, pm.tangkalpanjang_p, pm.alasbrondolan_p, pm.janjangperiksa_kb,  " & _
		 " pm.buahnormal_kb, pm.buahmentah_kb, pm.buahbusuk_kb, pm.overripe_kb, pm.buahabnormal_kb, pm.tangkalpanjang_kb, pm.alasbrondolan_kb," & _
		 " emp.fccode as empcode, emp.fcname as empname, tt.fccode, tt.fcname as spvname, pr.fccode, pr.fcname as namapemeriksa" & _
		 " from pemeriksaan_mutubuah pm" & _
		 " inner join employee emp on (pm.empcode = emp.fccode and pm.fcba  = emp.fcba) " & _
		 " inner join gang gg on (emp.gangcode = gg.fccode and pm.fcba = gg.fcba)" & _
		 " inner join pemeriksa pr on (pr.fccode = pm.pemeriksa and pr.fcba = pm.fcba)" & _
		 " left join " & _
		 " (select emp.fccode, emp.fcname, gg.fccode gangcode, emp.fcba " & _
		 " from employee emp " & _
		 " inner join gang gg on (emp.gangcode = gg.fccode and emp.fcba  = gg.fcba)" & _
		 " )tt on (pm.supervision = tt.fccode and pm.fcba = tt.fcba )" & _
		 " where pm.fcba = '" & Session("fcba") & "'" & _
		 " and pm.division = '"& division & "'" & _
		 " and transdate >=to_Date('" & transdate_From & "','MM/DD/YYYY') and transdate <=to_Date('" & transdate_To & "','MM/DD/YYYY') order by pm.transdate, emp.fccode, pm.fieldcode, namapemeriksa" 
Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open txtSQL, connstr

if  len(trim(transdate_From)) > 0 then
    txtSQL = txtSQL & " and transdate >=to_Date('" & transdate_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(transdate_From,1)
end if
if  len(trim(transdate_To)) > 0 then
    txtSQL =txtSQL &  " and transdate <=to_Date('" & transdate_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(transdate_To,1)
end if

txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

' response.Write(txtsql)
' response.Flush()
%>
 <body>
		<div id="wrapper">
		<!--#include file="print_excel.asp"-->
 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="22" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PEMERIKSAAN MUTU BUAH</h2>
				</div>
			</div>
		 </th>
			</tr>
						<!--<div class="printhide">-->
				<tr>
					<td colspan="22" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th rowspan="3">Date</th>
					<th rowspan="3">Blok</th>
					<th rowspan="3">No.TPH</th>
					<th rowspan="3">Kode Pemanen</th>
					<th rowspan="3">Nama Pemanen</th>
					<th rowspan="3">Nama Mandor</th>
					<th rowspan="3">Nama Pemeriksa</th>
					<th colspan="20">Diperiksa Oleh</th>
				</tr>
				
				<tr>
					<th colspan="8">Pemeriksa</th>
					<th colspan="8">Kerani Buah</th>
				</tr>
				<tr>
					<th>Janjang Periksa </th>
					<th>Buah Normal (N)</th>
					<th>Buah Mentah (A)</th>
					<th>Buah Busuk (E)</th>
					<th>Over Ripe (O)</th>
					<th>Buah Abnormal (BA)</th>
					<th>Tangkal Panjang (TP)</th>
					<th>Alas Brondolan (AB)</th>
					
					<th>Janjang Periksa </th>
					<th>Buah Normal (N)</th>
					<th>Buah Mentah (A)</th>
					<th>Buah Busuk (E)</th>
					<th>Over Ripe (O)</th>
					<th>Buah Abnormal (BA)</th>
					<th>Tangkal Panjang (TP)</th>
					<th>Alas Brondolan (AB)</th>					
				</tr>
                </thead>
               <tbody>
	<%
	if Not rscat.eof then
	While Not rscat.EOF
	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fieldcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("collectionpoint")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("empcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("empname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("spvname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("namapemeriksa")%></td>
			
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("janjangperiksa_p")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("buahnormal_p")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("buahmentah_p")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("buahbusuk_p")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("overripe_p")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("buahabnormal_p")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("tangkalpanjang_p")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("alasbrondolan_p")%></td>

			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("janjangperiksa_kb")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("buahnormal_kb")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("buahmentah_kb")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("buahbusuk_kb")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("overripe_kb")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("buahabnormal_kb")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("tangkalpanjang_kb")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("alasbrondolan_kb")%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='22'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->