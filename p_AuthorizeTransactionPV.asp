<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Payment Voucher Authorization</title>
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
<link rel="stylesheet" href="/css/fbp-table.css?v=20261010c" type="text/css" />
<script type="text/javascript" src="/js/fbp-table.js?v=20261009c"></script>
<style>
	/* Laporan antrean otorisasi PV */
	.apv-note { margin-top: 4px; font-size: 12px; font-weight: 400; color: #5d6d69; letter-spacing: 0; }
	#product-table td.apv-code { font-weight: 600; white-space: nowrap; text-align: left; }
	#product-table td.apv-amount { font-weight: 600; white-space: nowrap; }
	#product-table td.apv-date { white-space: nowrap; }
	.apv-next { display: inline-block; padding: 2px 10px; border-radius: 999px; background: #fff4d6; color: #8a5a00; font-weight: 600; white-space: nowrap; }
	#product-table tfoot td { background: #f6f9f8; font-weight: 700; border-top: 2px solid #1B7863; }
	@media print { .apv-next { background: none; padding: 0; } }
</style>

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

v_txtVoucherDate_From=trim(request("VOUCHERDATE_From"))
v_txtVoucherDate_To=trim(request("VOUCHERDATE_To"))
v_txtVoucherCode=trim(request("VOUCHERCODE"))
v_txtITEMCODE=trim(request("ITEMCODE"))
v_txtRequestorID=trim(request("REQUESTORID"))
v_txtRowStateId=trim(request("ROWSTATEID"))
v_txtpurchaseby=trim(request("ManagedId"))
businessunit=trim(request("businessunit"))

' Perbaikan: escape tanda kutip agar parameter tidak merusak/menyisipkan SQL
Function SqlStr(v)
	SqlStr = Replace(Trim(v & ""), "'", "''")
End Function

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & SqlStr(businessunit) & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if

txtSQL = 	"   SELECT DISTINCT stage1.businessunit fcba," & _
			"                   stage1.vouchercode," & _
			"                   stage1.voucherdate datecreated," & _
			"                   stage1.fccode authorizetemplatecode," & _
			"                   stage1.document_classification," & _
			"                   stage1.amount," & _
			"                   stage1.currentauthorizelevel," & _
			"                   stage1.authorizelevelname," & _
			"                   stage1.lastauthorizer," & _
			"                   stage1.lastactiondate," & _
			"                   stage1.authorizelevelcode," & _
			"                   aust.fcusercode," & _
			"                   stage1.accountingyear," & _
			"                   stage1.accountingperiod," & _
			"                   stage1.rowstateid," & _
			"                   stage1.bankcode," & _
			"                   stage1.bankname," & _
			"                   stage1.transtype," & _
			"                   stage1.remarks" & _
			"     FROM (SELECT abc.businessunit," & _
			"                  abc.vouchercode," & _
			"                  abc.voucherdate," & _
			"                  at.fccode," & _
			"                  abc.document_classification," & _
			"                  abc.amount," & _
			"                  abc.currentauthorizelevel," & _
			"                  abc.authorizelevelname," & _
			"                  abc.lastauthorizer," & _
			"                  abc.lastactiondate," & _
			"                  authorizestructure_api.get_nextlevelpv (" & _
			"                     'PV'," & _
			"                     abc.document_classification," & _
			"                     abc.businessunit," & _
			"                     NVL (abc.currentauthorizelevel, 0)," & _
			"                     abc.vouchercode," & _
			"                     abc.lastauthorizer)" & _
			"                     AS authorizelevelcode," & _
			"                  rowstateid," & _
			"                  accountingyear," & _
			"                  accountingperiod," & _
			"                  bankcode," & _
			"                  bankname," & _
			"                  transtype," & _
			"                  remarks" & _
			"             FROM (  SELECT iplasprod.paymentvoucher.fcba AS businessunit," & _
			"                            iplasprod.paymentvoucher.vouchercode AS vouchercode," & _
			"                            iplasprod.paymentvoucher.datecreated AS voucherdate," & _
			"                            iplasprod.paymentvoucher.document_classification," & _
			"                            SUM (iplasprod.paymentvoucherdetail.amount) AS amount," & _
			"                            iplasprod.authorizelevel.fccode" & _
			"                               AS currentauthorizelevel," & _
			"                            iplasprod.authorizelevel.fcname AS authorizelevelname," & _
			"                            iplasprod.paymentvoucher_hist.fcentry" & _
			"                               AS lastauthorizer," & _
			"                            iplasprod.paymentvoucher_hist.lastupdate" & _
			"                               AS lastactiondate," & _
			"                            iplasprod.paymentvoucher.rowstateid," & _
			"                            iplasprod.paymentvoucher.remarks," & _
			"                            iplasprod.paymentvoucher.transtype," & _
			"                            iplasprod.paymentvoucher.bankcode," & _
			"                            bank_api.get_bankcode_name (" & _
			"                               iplasprod.paymentvoucher.bankcode," & _
			"                               iplasprod.paymentvoucher.fcba)" & _
			"                               bankname," & _
			"                            iplasprod.paymentvoucher.accountingperiod," & _
			"                            iplasprod.paymentvoucher.accountingyear" & _
			"                       FROM iplasprod.paymentvoucher_hist" & _
			"                            INNER JOIN iplasprod.paymentvoucherdetail" & _
			"                               ON (    iplasprod.paymentvoucher_hist.fcba =" & _
			"                                          iplasprod.paymentvoucherdetail.fcba" & _
			"                                   AND iplasprod.paymentvoucher_hist.vouchercode =" & _
			"                                          iplasprod.paymentvoucherdetail.vouchercode)" & _
			"                            INNER JOIN iplasprod.paymentvoucher" & _
			"                               ON (    iplasprod.paymentvoucherdetail.fcba =" & _
			"                                          iplasprod.paymentvoucher.fcba" & _
			"                                   AND iplasprod.paymentvoucherdetail.vouchercode =" & _
			"                                          iplasprod.paymentvoucher.vouchercode)" & _
			"                            INNER JOIN iplasprod.authorizelevel" & _
			"                               ON iplasprod.paymentvoucher_hist.currentauthorizelevel =" & _
			"                                     iplasprod.authorizelevel.fccode" & _
			"                            INNER JOIN" & _
			"                            (  SELECT iplasprod.paymentvoucher_hist.vouchercode," & _
			"                                      MAX (" & _
			"                                         iplasprod.paymentvoucher_hist.currentauthorizelevel)" & _
			"                                         alevel" & _
			"                                 FROM iplasprod.paymentvoucher_hist" & _
			"                             GROUP BY iplasprod.paymentvoucher_hist.vouchercode)" & _
			"                            aa" & _
			"                               ON (    iplasprod.paymentvoucher_hist.currentauthorizelevel =" & _
			"                                          aa.alevel" & _
			"                                   AND iplasprod.paymentvoucher_hist.vouchercode =" & _
			"                                          aa.vouchercode)" & _
			"                   GROUP BY iplasprod.paymentvoucher.fcba," & _
			"                            iplasprod.paymentvoucher.vouchercode," & _
			"                            iplasprod.paymentvoucher.datecreated," & _
			"                            iplasprod.paymentvoucher.document_classification," & _
			"                            iplasprod.authorizelevel.fccode," & _
			"                            iplasprod.authorizelevel.fcname," & _
			"                            iplasprod.paymentvoucher.accountingperiod," & _
			"                            iplasprod.paymentvoucher.accountingyear," & _
			"                            iplasprod.paymentvoucher_hist.fcentry," & _
			"                            iplasprod.paymentvoucher_hist.lastupdate," & _
			"                            iplasprod.paymentvoucher.rowstateid," & _
			"                            iplasprod.paymentvoucher.bankcode," & _
			"                            bank_api.get_bankcode_name (" & _
			"                               iplasprod.paymentvoucher.bankcode," & _
			"                               iplasprod.paymentvoucher.fcba)," & _
			"                            iplasprod.paymentvoucher.remarks," & _
			"                            iplasprod.paymentvoucher.transtype) abc" & _
			"                  LEFT JOIN iplasprod.authorizetemplate at" & _
			"                     ON abc.document_classification = at.document_classification" & _
			"                  LEFT JOIN iplasprod.authorizestructure aus" & _
			"                     ON (    at.fccode = aus.authorizetemplatecode" & _
			"                         AND abc.currentauthorizelevel = aus.authorizelevelcode" & _
			"                         AND abc.businessunit = aus.fcba" & _
			"                         AND aus.fcusercode = abc.lastauthorizer)" & _
			"            WHERE abc.amount BETWEEN at.amountrangefrom AND at.amountrangeuntil)" & _
			"          stage1" & _
			"          INNER JOIN iplasprod.authorizestructure aust" & _
			"             ON     (    stage1.businessunit = aust.fcba" & _
			"                     AND stage1.authorizelevelcode = aust.authorizelevelcode" & _
			"                     AND stage1.fccode = aust.authorizetemplatecode )" & _
			"                AND stage1.rowstateid = 'AuthorizeOnProgress'" & _
			"                AND stage1.businessunit = '"& SqlStr(BusinessUnit) &"'"
 
			
if  len(trim(v_txtVoucherDate_From)) > 0 then
    txtSQL = txtSQL & " and stage1.voucherdate >=to_Date('" & SqlStr(v_txtVoucherDate_From) & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtVoucherDate_From,1)	
end if
if  len(trim(v_txtVoucherDate_To)) > 0 then
    txtSQL =txtSQL &  " and stage1.voucherdate <=to_Date('" & SqlStr(v_txtVoucherDate_To) & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtVoucherDate_To,1)		
end if
if  len(trim(v_txtVoucherCode)) > 0 then
    txtSQL = txtSQL & " and stage1.vouchercode='" & SqlStr(v_txtVoucherCode) & "'"
	txtSelect = txtSelect & " Voucher Code " & v_txtVoucherCode	
end if
	txtSQL = txtSQL & 	" ORDER BY stage1.transtype," & _
						"          stage1.bankcode," & _
						"          stage1.voucherdate," & _
						"          stage1.vouchercode" 


if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write(txtsql)
response.Flush()
%>
		<div id="wrapper">
		<!--#include file="print_excel.asp"-->
		<div class="fbt-summary" id="apv-summary"></div>
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PAYMENT VOUCHER AUTHORIZATION</h2>
					<div class="apv-note">Voucher yang sedang menunggu otorisasi (status AuthorizeOnProgress)</div>
				</div>
			</div>
					
		</th>
        </tr>
        		<!--<div class="printhide">-->
						<tr>
							<th colspan="11"  style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
							<%	
							
							Set rsCat = Server.CreateObject("ADODB.Recordset")
							rsCat.CursorType = 3
							rsCat.CursorLocation = 3
							rsCat.Open txtSQL, connstr

							
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr>
						<tr>
							<!--<th>No.</th>-->
							<th>Voucher Code</th>
							<th>Bank Code</th>
							<th>Bank Name</th>
							<th>Date Created</th>
							<th>Amount</th>
							<th>Last Authorize Level</th>
							<th>Last Authorize Level Name</th>
							<th>Last Authorizer</th>
							<th>Last Action Date</th>
							<th>Next Authorize Level</th>
							<th>Next Authorize User</th>
						</tr>
                        </thead>
                        <tbody>
							<%
							Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
							v_Quantity=0
							v_Amount = 0
							While Not rsCat.EOF%>
							<tr>
								<!--<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%> </td>-->
								<td class="apv-code" style="<%=cmdstyle%>"><%=rscat("vouchercode")%></td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("bankcode")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("bankname")%> </td>
								<td align="Center" class="apv-date" style="<%=cmdstyle%>"><%=rscat("datecreated")%></td>
								<td align="Right" class="apv-amount" style="<%=cmdstyle%>"><%=FormatNumber(cdbl(to_nol(rscat("amount"))),0)%></td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("currentauthorizelevel")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("authorizelevelname")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("lastauthorizer")%> </td>
								<td align="Center" class="apv-date" style="<%=cmdstyle%>"><%=rscat("lastactiondate")%></td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("authorizelevelcode")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><span class="apv-next"><%=rscat("fcusercode")%></span></td>
							</tr>
							  <%

							  v_Amount = v_Amount + cdbl(to_nol(rscat("amount")))
							  rsCat.MoveNext
							  'response.Flush()
							  no=no+1
							Wend
							if NO = 1 then%>
							<tr><td colspan="11" align="center" style="padding:16px;">Tidak ada voucher yang menunggu otorisasi untuk filter ini.</td></tr>
							<%end if
							rsCat.close()
							Set rsCat = Nothing%>
                            </tbody>
							<%if NO > 1 then%>
							<tfoot>
							<tr>
								<td colspan="4" style="<%=cmdstyle%>">Total (<span class="fbt-count"><%=NO-1%></span> voucher)</td>
								<td align="Right" class="apv-amount fbt-sum" style="<%=cmdstyle%>"><%=FormatNumber(v_Amount,0)%></td>
								<td colspan="6" style="<%=cmdstyle%>"></td>
							</tr>
							</tfoot>
							<%end if%>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
<script type="text/javascript">
// Tabel interaktif: cari, urut, pagination, ringkasan (hanya tampilan; data tetap dari server).
// Dijalankan setelah xTableHeaderFixed (event load) agar header tetap tidak ikut menggandakan toolbar.
window.addEventListener('load', function () {
	if (!window.FbpTable) return;
	var summary = document.getElementById('apv-summary');
	function fmt(n) { return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
	function cell(r, i) { return (r.cells[i].textContent || '').replace(/\s+/g, ' ').trim(); }
	FbpTable.patchExport();
	FbpTable.init(document.getElementById('product-table'), {
		pageSize: 10,
		sumCol: 4,
		onChange: function (rows) {
			var total = 0, who = {}, top = '-', topN = 0;
			rows.forEach(function (r) {
				total += parseFloat(cell(r, 4).replace(/,/g, '')) || 0;
				var u = cell(r, 10) || '-';
				who[u] = (who[u] || 0) + 1;
				if (who[u] > topN) { topN = who[u]; top = u; }
			});
			summary.innerHTML =
				'<div class="fbt-card"><div class="fbt-card-label">Voucher menunggu</div><div class="fbt-card-value">' + rows.length + '</div><div class="fbt-card-note">sesuai filter &amp; pencarian</div></div>' +
				'<div class="fbt-card"><div class="fbt-card-label">Total amount</div><div class="fbt-card-value">' + fmt(total) + '</div><div class="fbt-card-note">jumlah voucher yang tampil</div></div>' +
				'<div class="fbt-card accent"><div class="fbt-card-label">Paling banyak ditunggu</div><div class="fbt-card-value">' + (rows.length ? top : '-') + '</div><div class="fbt-card-note">' + (rows.length ? topN + ' voucher menunggu user ini' : 'tidak ada antrean') + '</div></div>';
		}
	});
});
</script>
	</BODY>
</HTML>	