<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="include/db_param.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>Reporting - Detail Listing Journal</title>
</head>
<%
' =====================================================================
' Export Excel listing jurnal (dibuka dari menu laporan, langsung mengunduh .xlsx).
' Percepatan: data diambil sekaligus dengan GetRows (dulu dibaca per kolom lewat nama,
' 26 kolom x puluhan ribu baris), markup HTML dipangkas, angka dikirim mentah agar
' terbaca sebagai angka di Excel. Query berparameter (dulu nilai URL disambung ke SQL).
' =====================================================================
Dim v_bulan, v_tahun, v_company, v_fcba, v_sourcecode, v_groupaccount
Dim jrnfrom, jrnto, tglprd, posting
v_bulan        = Trim(Request("bulan"))
v_tahun        = Trim(Request("tahun"))
v_company      = Trim(Request("company"))
v_fcba         = Trim(Request("fcba"))
v_sourcecode   = Trim(Request("sourcecode"))
v_groupaccount = Trim(Request("groupaccount"))
jrnfrom = Trim(Request("jrnfrom"))
jrnto   = Trim(Request("jrnto"))
tglprd  = Trim(Request("tglprd"))
posting = Request("posting")
Server.ScriptTimeout = 600
%>

<script src="/js/fbp-xlsx.js?v=20261015a"></script>
<script src="/distexcel/tableToExcel.js"></script>
<script type="text/javascript">

	function getExport()
	{
		var bulan 	= '<%= Server.HTMLEncode(v_bulan) %>';
		var tahun 	= '<%= Server.HTMLEncode(v_tahun) %>';
		var pcompanygl = '<%= Server.HTMLEncode(v_company) %>';
		var pfcbagl    = '<%= Server.HTMLEncode(v_fcba) %>';
		var nm = 'Journal_' + pcompanygl + '_' + pfcbagl + '_'+tahun+bulan+".xlsx";

		var table = document.getElementById("product-table");
		var done = function () { var l = document.getElementById('loadingjudul1'); if (l) l.innerHTML = 'File Excel sudah diunduh. Halaman ini boleh ditutup.'; };
		// cara cepat (fbp-xlsx.js); browser lama memakai tableToExcel.js seperti dulu
		var slow = function () {
			TableToExcel.convert(table, { name: nm, sheet: { name: "Journal" } });
			done();
		};
		if (window.FbpXlsx) FbpXlsx.fromTable(table, nm, "Journal").then(done, slow); else slow();
	}
</script>
<%
Dim q, prm(), np, jrRs, jrData, jrN, jrI, jrRow
ReDim prm(10) : np = 0
Sub AddP(v)
	prm(np) = v : np = np + 1
End Sub

' urutan kolom SELECT = urutan kolom tabel (indeks GetRows)
q = " SELECT b.accountingperiod, b.accountingyear, b.fcba, b.docno, to_char(b.docdate,'MM/DD/YYYY') docdate, b.sourcecode, a.source, a.bkmno, v.fcname vehiclename, a.description," & _
    " a.itemcode, a.itemname, a.jobcode, c.fcname jobname, a.locationtype, a.locationcode," & _
    " c.jobgroupcode, d.fcname jobgroupname, c.job_forbudgetcode, c.job_forbudgetname," & _
    " a.amountdebit, a.amountkredit, a.hk, a.unit," & _
    " CASE WHEN b.sourcecode = 'MAM-SIS' THEN e.UNITOFMEASUREMENT WHEN b.sourcecode = 'FAC-VRA' THEN vg.units ELSE c.uom_unit END AS uom," & _
    " a.output" & _
    " FROM bkmjournaldtl a INNER JOIN bkmjournalhdr b ON (a.docno=b.docno AND a.fcba=b.fcba) " & _
    " INNER JOIN jobmaster c ON (a.jobcode=c.fccode) " & _
    " LEFT JOIN jobgroupmaster d ON (c.jobgroupcode=d.fccode) " & _
    " LEFT JOIN bankstock e ON (a.itemcode = e.fccode) " & _
    " LEFT JOIN vehicle v on (v.fccode=a.bkmno and v.fcba=a.fcba)" & _
    " LEFT JOIN vehiclegroup vg on (vg.VEHICLEGROUPCODE =v.VEHICLEGROUPCODE and vg.fcba=v.fcba) " & _
    " WHERE "
If v_fcba = "-" Then
	q = q & " a.fcba IN (SELECT fccode FROM businessunit WHERE fccompanycode = ?) " : AddP v_company
Else
	q = q & " a.fcba = ? " : AddP v_fcba
End If
If tglprd = "Tgl" Then
	q = q & " AND b.DOCDATE between to_date(?,'mm/dd/yyyy') and to_date(?,'mm/dd/yyyy') " : AddP jrnfrom : AddP jrnto
Else
	q = q & " AND Accountingyear = ? AND Accountingperiod = ? " : AddP v_tahun : AddP v_bulan
End If
q = q & " AND SOURCECODE Like '%' || ? || '%'" : AddP v_sourcecode
If v_groupaccount <> "-" Then
	q = q & " AND C.jobgroupcode like '%' || ? || '%' " : AddP v_groupaccount
End If
q = q & " ORDER BY b.fcba, b.sourcecode, b.docdate, b.docno "
ReDim Preserve prm(np - 1)

Set jrRs = DbQuery(q, prm)
jrN = -1
If Not jrRs.EOF Then
	jrData = jrRs.GetRows()
	jrN = UBound(jrData, 2)
End If
jrRs.Close
Set jrRs = Nothing

' sel teks / tanggal / angka (angka mentah, titik desimal, tanpa pemisah ribuan)
Function JrT(v)
	JrT = "<td>" & v & "</td>"
End Function
Function JrN2(v)
	If IsNull(v) Then JrN2 = "<td data-t=""n""></td>" Else JrN2 = "<td data-t=""n"">" & Replace(CStr(v), ",", ".") & "</td>"
End Function
%>

<body>
    <span id="loadingjudul1">Proses export data (<%=jrN + 1%> baris)...</span>
    <div id="tableview" style="display:none;">
	<table border="0" width="100%" id="product-table">
		<thead>
			<tr>
				<th align="left" colspan="26" style="background-color:#FFFFFF; color:#000000; border:none; text-align:center;">JOURNAL LISTING <%=Server.HTMLEncode(v_company)%> Unit : <%=Server.HTMLEncode(v_fcba)%> Periode : <%=Server.HTMLEncode(v_tahun)%> / <%=Server.HTMLEncode(v_bulan)%></th>
			</tr>
			<tr>
				<th>Acc.Period</th><th>Acc.Year</th><th>FCBA</th><th>No.Journal</th><th>Tanggal</th><th>SourceCode</th><th>Source</th><th>No.Reference</th><th>Vehicle Name</th><th>Deskripsi</th>
				<th>Kode Barang</th><th>Nama Barang</th><th>Kode Akun</th><th>Nama Akun</th><th>Target Type</th><th>Target Code</th><th>Group Akun</th><th>Group Nama</th><th>Budget Akun</th><th>Budget Name</th>
				<th>Debit</th><th>Kredit</th><th>Hk</th><th>Unit</th><th>UOM</th><th>Output(Ha)</th>
			</tr>
		</thead>
<tbody>
<%
For jrI = 0 To jrN
	jrRow = "<tr>" & JrT(jrData(0, jrI)) & JrT(jrData(1, jrI)) & JrT(jrData(2, jrI)) & JrT(jrData(3, jrI)) & _
		"<td data-t=""d"">" & jrData(4, jrI) & "</td>" & _
		JrT(jrData(5, jrI)) & JrT(jrData(6, jrI)) & JrT(jrData(7, jrI)) & JrT(jrData(8, jrI)) & JrT(jrData(9, jrI)) & _
		JrT(jrData(10, jrI)) & JrT(jrData(11, jrI)) & JrT(jrData(12, jrI)) & JrT(jrData(13, jrI)) & JrT(jrData(14, jrI)) & JrT(jrData(15, jrI)) & _
		JrT(jrData(16, jrI)) & JrT(jrData(17, jrI)) & JrT(jrData(18, jrI)) & JrT(jrData(19, jrI)) & _
		JrN2(jrData(20, jrI)) & JrN2(jrData(21, jrI)) & JrN2(jrData(22, jrI)) & JrN2(jrData(23, jrI)) & _
		JrT(jrData(24, jrI)) & JrN2(jrData(25, jrI)) & "</tr>"
	Response.Write jrRow & vbCrLf
	If jrI Mod 2000 = 1999 Then Response.Flush   ' kirim bertahap: browser mulai menerima selagi server masih menulis
Next
Response.Write "<script language='javascript'> getExport();</script>"
%>
</tbody>
</table>
</div>

</BODY>
</HTML>
