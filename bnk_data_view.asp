<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/include/db_param.asp"-->
<%
' =====================================================================
' Rincian mutasi Bank/Kas per bulan (dibuka dari daftar Bank di Select_Form_Show).
' Tab Payment, Receive dan Cash Flow. Isi & perhitungan sama dengan versi lama;
' tampilan baru + tabel ala DataTables (fbp-shell -> fbp-table-auto).
' Perbaikan: semua nilai dari URL kini berparameter (dulu disambung ke SQL, termasuk
' blok PL/SQL FAC_BNK_STATUSPOST) dan BA harus di-assign ke user.
' =====================================================================
Dim bvBank, bvMonth, bvYear, bvFcba, bvErr, bvRs, bvCompany, bvBaName, bvBankName, bvControlJob
Dim bvPayRows, bvRcvRows, bvCfRows, bvNo, bvTotPay, bvTotRcv, bvOpening, bvSaldo, bvDebet, bvCredit
Dim bvPayCnt, bvRcvCnt, bvCfCnt, bvSql1, bvSql2, bvSql3, bvCfSql, bvPrint, bvPost, bvMonthName

bvBank = Trim(Request("bankcode") & "")
bvMonth = Trim(Request("month") & "")
bvYear = Trim(Request("tahun") & "")
bvFcba = Trim(Request("fcba") & "")
bvErr = ""

If bvBank = "" Or bvFcba = "" Then
	bvErr = "Parameter bank atau BA tidak lengkap."
ElseIf Not (IsNumeric(bvMonth) And IsNumeric(bvYear)) Then
	bvErr = "Periode tidak valid."
ElseIf CDbl(bvMonth) < 1 Or CDbl(bvMonth) > 12 Or CDbl(bvYear) < 2000 Or CDbl(bvYear) > 2100 Or InStr(bvMonth, ".") > 0 Or InStr(bvYear, ".") > 0 Then
	bvErr = "Periode tidak valid."
End If

If bvErr = "" And UCase(bvFcba) <> UCase(Session("fcBA") & "") Then
	Set bvRs = DbQuery("select count(*) n from user_fcba_tab where fcusercode = ? and fcba = ?", Array(Session("USERNAME"), bvFcba))
	If CLng(bvRs("n")) = 0 Then bvErr = "Anda tidak punya akses ke BA " & bvFcba & "."
	bvRs.Close
End If

Function BvH(v)
	BvH = Server.HTMLEncode(v & "")
End Function

' angka: tanpa desimal, kecuali nilainya memang berdesimal
Function BvNum(v)
	Dim n
	n = CDbl(to_nol(v))
	If n <> Int(n) Then BvNum = FormatNumber(n, 2) Else BvNum = FormatNumber(n, 0)
End Function

' sel teks panjang: tampil maks 2 baris (klik untuk buka), teks lengkap di tooltip
Function BvT(v)
	Dim t
	t = BvH(v)
	If t = "" Then BvT = "<td class=""t""></td>" Else BvT = "<td class=""t""><div class=""clip"" title=""" & t & """>" & t & "</div></td>"
End Function

Function BvState(s)
	Dim c
	c = LCase(s & "")
	If InStr(c, "cancel") > 0 Or InStr(c, "reject") > 0 Then
		BvState = "bad"
	ElseIf InStr(c, "approv") > 0 Or InStr(c, "post") > 0 Then
		BvState = "ok"
	ElseIf InStr(c, "progress") > 0 Or InStr(c, "confirm") > 0 Then
		BvState = "info"
	Else
		BvState = "warn"
	End If
End Function

' tombol aksi: edit, cetak voucher, lihat jurnal
Function BvActions(r, printFile, jurnalNo)
	BvActions = "<div class=""bv-act"">" & _
		"<a class=""bv-ic"" title=""Edit voucher"" href=""PaymentVoucherDetail.Asp?fcCode=" & Server.URLEncode(r("VoucherCode") & "") & "&transtype=" & Server.URLEncode(r("transtype") & "") & "&txtlineno=" & Server.URLEncode(r("lineno") & "") & "&xCode=" & Server.URLEncode(r("bankcode") & "") & "&fcba=" & Server.URLEncode(r("fcba") & "") & "&txtJenis=Edit""><i class=""fa fa-pencil"" aria-hidden=""true""></i><span class=""sr"">Edit</span></a>" & _
		"<a class=""bv-ic"" title=""Cetak voucher"" target=""_blank"" href=""" & printFile & ".asp?fcba=" & Server.URLEncode(bvFcba) & "&vouchercode=" & Server.URLEncode(r("VoucherCode") & "") & "&transtype=" & Server.URLEncode(r("transtype") & "") & """><i class=""fa fa-print"" aria-hidden=""true""></i><span class=""sr"">Cetak</span></a>" & _
		"<a class=""bv-ic"" title=""Lihat jurnal"" target=""_blank"" href=""bkm_jurnal_view.Asp?bkmno=" & Server.URLEncode(jurnalNo & "") & "&tglbkm=" & Server.URLEncode(r("datecreated") & "") & "&xCode=" & Server.URLEncode(r("bankcode") & "") & "&fcba=" & Server.URLEncode(r("fcba") & "") & "&tipetran=FAC-BNK""><i class=""fa fa-book"" aria-hidden=""true""></i><span class=""sr"">Jurnal</span></a>" & _
		"</div>"
End Function

If bvErr = "" Then
	DbExec "BEGIN fac_journal_fac_bnk_api.FAC_BNK_STATUSPOST(to_number(?), to_number(?), ?, 'System'); END;", Array(bvYear, bvMonth, bvFcba)

	Set bvRs = DbQuery("select fcname, fccompanyname from BusinessUnit where fcCode = ?", Array(bvFcba))
	If Not bvRs.EOF Then bvBaName = bvRs("fcname") : bvCompany = bvRs("fccompanyname")
	bvRs.Close

	Set bvRs = DbQuery("select fccode, fcname, controljob from bank where fccode = ? and fcba = ?", Array(bvBank, bvFcba))
	If Not bvRs.EOF Then bvBankName = bvRs("fcname") : bvControlJob = bvRs("controljob") & ""
	bvRs.Close
	bvMonthName = MonthName(CLng(bvMonth)) & " " & bvYear

	' ---------- Payment ----------
	Set bvRs = DbQuery(" SELECT a.vouchercode,to_char(datecreated,'dd/mm/yyyy')datecreated,referenceno,cekno,classification_api.get_fcname(document_classification) class, " & _
		" decode(nvl(b.bano,''),'',b.remarks,b.remarks||' ('||b.bano||')') remarksheader, a.remarks,nvl(a.amount,0) totalamount,rowstateid,transtype,bankcode,a.fcba,nvl(b.postingflg,'0') postingflg" & _
		" ,b.accountingperiod,b.fcentry,b.fcedit,b.accountingyear,a.lineno,a.jobcode||'-'||a.locationcode Akun " & _
		" FROM paymentvoucherdetail a " & _
		" INNER JOIN paymentvoucher b ON (a.vouchercode=b.vouchercode AND a.fcba=b.fcba)  " & _
		" LEFT JOIN bkmjournalhdr c ON (a.vouchercode=c.docno AND a.fcba=c.fcba) " & _
		" where a.fcba = ? AND bankcode = ?" & _
		" AND to_number(to_char(datecreated,'YYYY')) = to_number(?)" & _
		" AND to_number(to_char(datecreated,'MM')) = to_number(?)" & _
		" AND transtype='02-PAYMENT' " & _
		" ORDER BY datecreated DESC, a.vouchercode DESC ", Array(bvFcba, bvBank, bvYear, bvMonth))
	bvPayRows = "" : bvNo = 0 : bvTotPay = 0
	Do While Not bvRs.EOF
		bvNo = bvNo + 1
		If bvRs("postingflg") & "" = "0" Then bvPost = "Tidak" Else bvPost = "YA"
		If bvRs("rowstateid") = "Confirmed" Or bvRs("rowstateid") = "Planned" Then bvPrint = "printvoucherfin" Else bvPrint = "printvoucherfinapprove"
		bvPayRows = bvPayRows & "<tr>" & _
			"<td class=""n"">" & bvNo & "</td>" & _
			"<td class=""code"">" & BvH(bvRs("vouchercode")) & "</td>" & _
			"<td class=""nw"">" & BvH(bvRs("datecreated")) & "</td>" & _
			BvT(bvRs("remarksheader")) & _
			"<td class=""num"">" & BvNum(bvRs("totalamount")) & "</td>" & _
			"<td><span class=""bv-badge " & BvState(bvRs("rowstateid")) & """>" & BvH(bvRs("rowstateid")) & "</span></td>" & _
			"<td class=""c"">" & BvActions(bvRs, bvPrint, bvRs("VoucherCode")) & "</td>" & _
			BvT(bvRs("remarks")) & _
			BvT(bvRs("referenceno")) & _
			"<td class=""nw"">" & BvH(bvRs("accountingperiod")) & "-" & BvH(bvRs("accountingyear")) & "</td>" & _
			"<td>" & BvH(bvRs("cekno")) & "</td>" & _
			BvT(bvRs("class")) & _
			"<td class=""nw"">" & BvH(bvRs("Akun")) & "</td>" & _
			"<td class=""c"">" & bvPost & "</td>" & _
			"<td>" & BvH(bvRs("fcentry")) & "</td>" & _
			"<td>" & BvH(bvRs("fcedit")) & "</td>" & _
			"</tr>"
		bvTotPay = bvTotPay + CDbl(bvRs("totalamount"))
		bvRs.MoveNext
	Loop
	bvRs.Close
	bvPayCnt = bvNo

	' ---------- Receive (termasuk payment BA lain yang charge ke bank ini) ----------
	Set bvRs = DbQuery(" SELECT a.vouchercode,to_char(datecreated,'dd/mm/yyyy')datecreated,referenceno,cekno,classification_api.get_fcname(document_classification) class," & _
		" a.remarks,b.remarks remarksheader, nvl(a.amount,0) totalamount,rowstateid,transtype,bankcode,a.fcba,nvl(b.postingflg,'0') postingflg" & _
		" ,b.accountingperiod,b.accountingyear,b.fcentry,b.fcedit,a.lineno,nvl(c.docno,a.vouchercode) docno " & _
		" FROM paymentvoucherdetail a " & _
		" INNER JOIN paymentvoucher b ON (a.vouchercode=b.vouchercode AND a.fcba=b.fcba)  " & _
		" LEFT JOIN bkmjournalhdr c ON (a.vouchercode=c.docno AND a.fcba=c.fcba) " & _
		" where a.fcba = ? AND bankcode = ?" & _
		" AND to_number(to_char(datecreated,'YYYY')) = to_number(?)" & _
		" AND to_number(to_char(datecreated,'MM')) = to_number(?)" & _
		" AND transtype='01-RECEIVE' " & _
		" UNION " & _
		" SELECT a.vouchercode,to_char(datecreated,'dd/mm/yyyy')datecreated,referenceno,cekno,classification_api.get_fcname(document_classification) class," & _
		" a.remarks,b.remarks remarksheader,nvl(a.amount,0) totalamount,rowstateid,transtype,bankcode,a.fcba,nvl(b.postingflg,'0') postingflg" & _
		" ,b.accountingperiod,b.accountingyear,b.fcentry,b.fcedit,a.lineno,nvl(c.docno,Trim(a.vouchercode)||'-INT') docno " & _
		" FROM paymentvoucherdetail a " & _
		" INNER JOIN paymentvoucher b ON (a.vouchercode=b.vouchercode AND a.fcba=b.fcba)  " & _
		" INNER JOIN bkmjournalhdr c ON (Trim(a.vouchercode)||'-INT'=Trim(c.docno)) " & _
		" where a.chargeunit = ? AND chargecode = ?" & _
		" AND to_number(to_char(datecreated,'YYYY')) = to_number(?)" & _
		" AND to_number(to_char(datecreated,'MM')) = to_number(?)" & _
		" AND transtype='02-PAYMENT' " & _
		" ORDER BY datecreated DESC, vouchercode DESC ", Array(bvFcba, bvBank, bvYear, bvMonth, bvFcba, bvBank, bvYear, bvMonth))
	bvRcvRows = "" : bvNo = 0 : bvTotRcv = 0
	Do While Not bvRs.EOF
		bvNo = bvNo + 1
		If bvRs("postingflg") & "" = "0" Then bvPost = "Tidak" Else bvPost = "YA"
		If bvRs("rowstateid") = "Confirmed" Or bvRs("rowstateid") = "Planned" Then bvPrint = "printvoucherfin" Else bvPrint = "printvoucherinfinapprove"
		bvRcvRows = bvRcvRows & "<tr>" & _
			"<td class=""n"">" & bvNo & "</td>" & _
			"<td class=""code"">" & BvH(bvRs("vouchercode")) & "</td>" & _
			"<td class=""nw"">" & BvH(bvRs("datecreated")) & "</td>" & _
			BvT(bvRs("remarksheader")) & _
			"<td class=""num"">" & BvNum(bvRs("totalamount")) & "</td>" & _
			"<td><span class=""bv-badge " & BvState(bvRs("rowstateid")) & """>" & BvH(bvRs("rowstateid")) & "</span></td>" & _
			"<td class=""c"">" & BvActions(bvRs, bvPrint, bvRs("docno")) & "</td>" & _
			BvT(bvRs("remarks")) & _
			BvT(bvRs("referenceno")) & _
			"<td class=""nw"">" & BvH(bvRs("accountingperiod")) & "-" & BvH(bvRs("accountingyear")) & "</td>" & _
			"<td>" & BvH(bvRs("cekno")) & "</td>" & _
			BvT(bvRs("class")) & _
			"<td class=""c"">" & bvPost & "</td>" & _
			"<td>" & BvH(bvRs("fcentry")) & "</td>" & _
			"<td>" & BvH(bvRs("fcedit")) & "</td>" & _
			"</tr>"
		bvTotRcv = bvTotRcv + CDbl(bvRs("totalamount"))
		bvRs.MoveNext
	Loop
	bvRs.Close
	bvRcvCnt = bvNo

	' ---------- Cash Flow ----------
	' Saldo awal dari trial balance akun kontrol bank
	Set bvRs = DbQuery("SELECT nvl(sum(OPENING_BALANCE),0) as OPENING_BALANCE FROM fac_trialbalance WHERE FCCODE = ? AND FCBA = ? AND PERIOD_MONTH = ? AND PERIOD_year = ?", Array(bvControlJob, bvFcba, bvMonth, bvYear))
	If Not bvRs.EOF Then bvOpening = CDbl(to_nol(bvRs("OPENING_BALANCE"))) Else bvOpening = 0
	bvRs.Close

	' Penerimaan (debet): receive bank ini + payment BA lain yang charge ke bank ini; pengeluaran (credit): payment bank ini
	bvSql1 = "  SELECT p.datecreated, p.bankcode, p.transtype, pay.vouchercode, p.remarks orang, pay.jobcode, job.fcname jobname, pay.locationtype, pay.locationcode," & _
		" NVL (SUM (pay.amount), 0) AS debet, 0 as credit, pay.REFERENCE, pay.remarks" & _
		" FROM paymentvoucherdetail pay " & _
		" INNER JOIN paymentvoucher p ON pay.vouchercode=p.vouchercode and pay.fcba=p.fcba " & _
		" inner join jobmaster job on pay.jobcode = job.fccode " & _
		" LEFT JOIN bkmjournalhdr c ON (p.vouchercode=c.docno AND p.fcba=c.fcba) " & _
		" WHERE pay.fcba = ? AND p.transtype = '01-RECEIVE'" & _
		" AND p.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
		" and TO_NUMBER(TO_CHAR(p.datecreated,'mm')) = to_number(?)" & _
		" and TO_NUMBER(TO_CHAR(p.datecreated,'yyyy')) = to_number(?)" & _
		" and (p.BANKCODE = ?) " & _
		" GROUP BY p.datecreated, p.bankcode, p.transtype, pay.vouchercode, p.remarks, pay.jobcode, job.fcname, pay.locationtype, pay.locationcode, pay.REFERENCE, pay.remarks"
	bvSql2 = "  SELECT p.datecreated, p.bankcode, p.transtype, pay.vouchercode, p.remarks orang, pay.jobcode, job.fcname jobname, pay.locationtype, pay.locationcode," & _
		" NVL (SUM (pay.amount), 0) AS debet, 0 as credit, pay.REFERENCE, pay.remarks" & _
		" FROM paymentvoucherdetail pay " & _
		" INNER JOIN paymentvoucher p ON pay.vouchercode=p.vouchercode and pay.fcba=p.fcba " & _
		" inner join jobmaster job on pay.jobcode = job.fccode " & _
		" LEFT JOIN bkmjournalhdr c ON (p.vouchercode=c.docno AND p.fcba=c.fcba) " & _
		" WHERE pay.chargeunit = ? AND p.transtype = '02-PAYMENT'" & _
		" AND p.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _
		" and TO_NUMBER(TO_CHAR(p.datecreated,'mm')) = to_number(?)" & _
		" and TO_NUMBER(TO_CHAR(p.datecreated,'yyyy')) = to_number(?)" & _
		" and (pay.chargecode = ?)" & _
		" GROUP BY p.datecreated, p.bankcode, p.transtype, pay.vouchercode, p.remarks, pay.jobcode, job.fcname, pay.locationtype, pay.locationcode, pay.REFERENCE, pay.remarks"
	bvSql3 = "  SELECT p.datecreated, p.bankcode, p.transtype, pay.vouchercode, p.remarks orang, pay.jobcode, job.fcname jobname, pay.locationtype, pay.locationcode," & _
		" 0 as debet, NVL (SUM (pay.amount), 0) AS credit, pay.REFERENCE, pay.remarks" & _
		" FROM paymentvoucherdetail pay " & _
		" INNER JOIN paymentvoucher p ON pay.vouchercode=p.vouchercode and pay.fcba=p.fcba " & _
		" inner join jobmaster job on pay.jobcode = job.fccode " & _
		" LEFT JOIN bkmjournalhdr c ON (p.vouchercode=c.docno AND p.fcba=c.fcba) " & _
		" WHERE pay.fcba = ? AND p.transtype = '02-PAYMENT'" & _
		" AND p.rowstateid in('Confirmed','AuthorizeOnProgress','Approved')" & _
		" and TO_NUMBER(TO_CHAR(p.datecreated,'MM')) = to_number(?)" & _
		" and TO_NUMBER(TO_CHAR(p.datecreated,'yyyy')) = to_number(?)" & _
		" and (p.BANKCODE = ?)" & _
		" GROUP BY p.datecreated, p.bankcode, p.transtype, pay.vouchercode, p.remarks, pay.jobcode, job.fcname, pay.locationtype, pay.locationcode, pay.REFERENCE, pay.remarks"
	bvCfSql = " SELECT vouchercode, to_char(datecreated,'dd/mm/yyyy') tgl, orang, bankcode, jobcode, jobname, transtype, locationtype, locationcode, reference, remarks," & _
		" SUM (NVL (debet, 0)) debet, SUM (NVL (credit, 0)) credit" & _
		" FROM ( " & bvSql1 & " union all " & bvSql2 & " union all " & bvSql3 & ")" & _
		" GROUP BY vouchercode, datecreated, bankcode, transtype, jobcode, orang, jobname, locationtype, locationcode, reference, remarks" & _
		" ORDER BY datecreated, transtype, vouchercode, bankcode "
	Set bvRs = DbQuery(bvCfSql, Array(bvFcba, bvMonth, bvYear, bvBank, bvFcba, bvMonth, bvYear, bvBank, bvFcba, bvMonth, bvYear, bvBank))
	bvCfRows = "" : bvNo = 0 : bvSaldo = bvOpening : bvDebet = 0 : bvCredit = 0
	Do While Not bvRs.EOF
		bvNo = bvNo + 1
		bvSaldo = bvSaldo + CDbl(to_nol(bvRs("debet"))) - CDbl(to_nol(bvRs("credit")))
		bvDebet = bvDebet + CDbl(to_nol(bvRs("debet")))
		bvCredit = bvCredit + CDbl(to_nol(bvRs("credit")))
		bvCfRows = bvCfRows & "<tr>" & _
			"<td class=""n"">" & bvNo & "</td>" & _
			"<td class=""nw"">" & BvH(bvRs("tgl")) & "</td>" & _
			"<td class=""code"">" & BvH(bvRs("vouchercode")) & "</td>" & _
			"<td class=""num"">" & BvNum(bvRs("debet")) & "</td>" & _
			"<td class=""num"">" & BvNum(bvRs("credit")) & "</td>" & _
			"<td class=""num b"">" & BvNum(bvSaldo) & "</td>" & _
			BvT(bvRs("orang")) & _
			BvT(bvRs("Reference")) & _
			BvT(bvRs("Remarks")) & _
			"</tr>"
		bvRs.MoveNext
	Loop
	bvRs.Close
	bvCfCnt = bvNo
End If
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>FBPlants - Bank View <%=BvH(bvBank)%></title>
<script type="text/javascript" src="stlib.js"></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<script type="text/javascript" src="dtree.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014b"></script>
<style>
	.bv { --brand: #1B7863; --brand-dark: #12584a; --tint: #eef6f3; --line: #e1e7e5; --muted: #5d6d69; --ink: #17262a; padding: 4px 2px 10px; color: var(--ink); font-family: "Plus Jakarta Sans", system-ui, sans-serif; }
	.bv-head { display: flex; flex-wrap: wrap; align-items: flex-start; justify-content: space-between; gap: 14px 20px; margin-bottom: 16px; }
	.bv-eyebrow { font-size: 12px; font-weight: 600; letter-spacing: .3px; text-transform: uppercase; color: var(--muted); }
	.bv-head h1 { margin: 4px 0 2px; font-size: 21px; font-weight: 700; color: var(--ink); }
	.bv-head h1 .code { font: 600 13px ui-monospace, Consolas, monospace; color: var(--brand-dark); background: var(--tint); padding: 3px 8px; border-radius: 6px; vertical-align: middle; margin-left: 6px; }
	.bv-head p { margin: 0; font-size: 13px; color: var(--muted); }
	.box_both .bv .bv-btns { display: flex; flex-wrap: wrap; gap: 8px; }
	.box_both .bv .bv-btn { display: inline-flex; align-items: center; gap: 7px; height: 38px; padding: 0 14px; font: 600 13px "Plus Jakarta Sans", system-ui, sans-serif; color: var(--brand-dark); background: #fff; border: 1px solid #cfe0da; border-radius: 9px; cursor: pointer; text-decoration: none; }
	.box_both .bv .bv-btn:hover { background: var(--tint); }
	.box_both .bv .bv-btn.pri { color: #fff; background: var(--brand); border-color: var(--brand); }
	.box_both .bv .bv-btn.pri:hover { background: var(--brand-dark); }
	.bv-kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 10px; margin-bottom: 18px; }
	.bv-kpi { padding: 12px 15px; background: #fff; border: 1px solid var(--line); border-radius: 12px; }
	.bv-kpi span { display: block; font-size: 11.5px; font-weight: 600; letter-spacing: .3px; text-transform: uppercase; color: var(--muted); }
	.bv-kpi b { display: block; margin-top: 4px; font-size: 18px; font-weight: 700; font-variant-numeric: tabular-nums; }
	.bv-kpi small { display: block; margin-top: 2px; font-size: 12px; color: var(--muted); }
	.bv-kpi.in b { color: #136a4b; } .bv-kpi.out b { color: #a1342a; }
	.bv-card { background: #fff; border: 1px solid var(--line); border-radius: 14px; }
	.box_both .bv .bv-tabs { display: flex; gap: 4px; padding: 10px 12px 0; border-bottom: 1px solid var(--line); overflow-x: auto; }
	.box_both .bv .bv-tab { min-height: 0; border-radius: 0; transition: none; position: relative; display: inline-flex; align-items: center; gap: 8px; height: 42px; padding: 0 16px; font: 600 13.5px "Plus Jakarta Sans", system-ui, sans-serif; color: var(--muted); background: none; border: 0; border-bottom: 2px solid transparent; cursor: pointer; white-space: nowrap; }
	.box_both .bv .bv-tab:hover { color: var(--ink); }
	.box_both .bv .bv-tab[aria-selected="true"] { color: var(--brand-dark); border-bottom-color: var(--brand); }
	.box_both .bv .bv-tab .cnt { min-width: 22px; padding: 1px 7px; font-size: 11.5px; font-weight: 700; color: var(--muted); background: #eef2f1; border-radius: 999px; }
	.box_both .bv .bv-tab[aria-selected="true"] .cnt { color: #fff; background: var(--brand); }
	.bv-panel { padding: 14px 16px 16px; }
	.bv-panel[hidden] { display: none; }
	.bv-note { margin: 0 0 10px; font-size: 13px; color: var(--muted); }
	.bv-note b { color: var(--ink); font-variant-numeric: tabular-nums; }
	.bv-empty { padding: 34px 12px; text-align: center; color: var(--muted); font-size: 13.5px; }
	.bv-err { max-width: 640px; margin: 30px auto; padding: 22px 24px; background: #fff; border: 1px solid #f3c2bc; border-radius: 12px; color: #8a2a20; }
	.bv-err b { display: block; margin-bottom: 4px; font-size: 15px; }
	/* tabel: jarak sel lega (permintaan user), angka rata kanan */
	.bv table.data { width: 100%; }
	.bv table.data > thead > tr:last-child > th { padding: 11px 14px !important; font-size: 11.5px; font-weight: 700; letter-spacing: .3px; text-transform: uppercase; text-align: left; color: var(--muted) !important; background: #f6f9f8 !important; white-space: nowrap; }
	.bv table.data > tbody > tr > td { padding: 11px 14px !important; font-size: 13px; line-height: 1.45; vertical-align: top; text-align: left; border-bottom: 1px solid #edf1f0; }
	.bv table.data > tbody > tr:hover > td { background: #f6faf8; }
	.bv table.data > tfoot > tr > td { padding: 12px 14px !important; font-size: 13px; font-weight: 700; background: #f6f9f8; border-top: 2px solid #d5dfdc; }
	.bv table.data th.num, .bv table.data td.num { text-align: right !important; white-space: nowrap; font-variant-numeric: tabular-nums; }
	.bv table.data .c { text-align: center !important; }
	.bv table.data .n { color: var(--muted); white-space: nowrap; }
	.bv table.data .nw { white-space: nowrap; }
	.bv table.data .b { font-weight: 700; }
	/* kolom teks panjang boleh terbungkus agar kolom lain tidak ikut disembunyikan (fbp-table: td nowrap) */
	.bv table.data > tbody > tr > td.t { white-space: normal !important; min-width: 150px; max-width: 260px; overflow-wrap: anywhere; }
	.bv .clip { display: -webkit-box; -webkit-box-orient: vertical; -webkit-line-clamp: 2; line-clamp: 2; overflow: hidden; cursor: pointer; }
	.bv .clip.open { display: block; -webkit-line-clamp: unset; line-clamp: unset; }
	.bv .fbt-rd .clip { display: block; -webkit-line-clamp: unset; line-clamp: unset; cursor: default; }
	.bv table.data > thead > tr:last-child > th { white-space: nowrap !important; }
	.bv table.data > tbody > tr > td { vertical-align: middle !important; }
	.bv table.data .code { font: 600 12px ui-monospace, Consolas, monospace; color: var(--brand-dark); white-space: nowrap; }
	.bv-badge { display: inline-block; padding: 2px 9px; font-size: 11.5px; font-weight: 600; border-radius: 999px; white-space: nowrap; background: #eef1f0; color: #4a5a56; }
	.bv-badge.ok { background: #e3f4ec; color: #136a4b; } .bv-badge.info { background: #e8f1fb; color: #1f5a96; }
	.bv-badge.warn { background: #fdf3dc; color: #8a5a00; } .bv-badge.bad { background: #fdecea; color: #a1342a; }
	.bv-act { display: inline-flex; gap: 4px; }
	.bv-ic { display: inline-grid; place-items: center; width: 30px; height: 30px; color: var(--brand-dark); background: var(--tint); border: 1px solid #cfe3dc; border-radius: 8px; text-decoration: none; }
	.bv-ic:hover { background: #dff0ea; }
	.bv .sr { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); }
	.bv .fbt-dash-tools { max-width: none; }
	@media (max-width: 640px) { .bv { padding: 14px 12px 24px; } .bv-head h1 { font-size: 18px; } }
	@media print { .bv-btns, .bv-tabs { display: none !important; } .bv-panel[hidden] { display: block; } }
</style>
</head>
<body class="dash page-show page-legacy">
<!--#include file="top.jsp"-->
<div style="width:21.7%; float:left; margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
	<% if Session("fcBa") <> "" then %>
	<!--#include file="menu.asp"-->
	<% end if %>
</div>
<div class="judul_head" style="width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
<div class="judul">Module : Finance and Accounting / Bank View</div>
<div class="box_both">
<div class="bv">
<% If bvErr <> "" Then %>
	<div class="bv-err"><b>Data tidak bisa ditampilkan</b><%=BvH(bvErr)%></div>
<% Else %>
	<div class="bv-head">
		<div>
			<div class="bv-eyebrow">PT. <%=BvH(bvCompany)%> &middot; <%=BvH(bvBaName)%> &middot; Module Bank</div>
			<h1><%=BvH(bvBankName)%><span class="code"><%=BvH(bvBank)%></span></h1>
			<p>Periode <b><%=BvH(bvMonthName)%></b> (<%=BvH(bvMonth)%>/<%=BvH(bvYear)%>)</p>
		</div>
		<div class="bv-btns">
			<button type="button" class="bv-btn pri" data-new="PYV"><i class="fa fa-minus-circle" aria-hidden="true"></i> New Payment</button>
			<button type="button" class="bv-btn pri" data-new="RCV"><i class="fa fa-plus-circle" aria-hidden="true"></i> New Receive</button>
			<button type="button" class="bv-btn" id="bvExport"><i class="fa fa-file-excel-o" aria-hidden="true"></i> Export XLS</button>
			<button type="button" class="bv-btn" onclick="window.close();"><i class="fa fa-times" aria-hidden="true"></i> Tutup</button>
		</div>
	</div>

	<div class="bv-kpis">
		<div class="bv-kpi"><span>Saldo Awal</span><b><%=BvNum(bvOpening)%></b><small>Trial balance akun <%=BvH(bvControlJob)%></small></div>
		<div class="bv-kpi in"><span>Penerimaan (Debet)</span><b><%=BvNum(bvDebet)%></b><small>Cash flow</small></div>
		<div class="bv-kpi out"><span>Pengeluaran (Credit)</span><b><%=BvNum(bvCredit)%></b><small>Cash flow</small></div>
		<div class="bv-kpi"><span>Saldo Akhir</span><b><%=BvNum(bvSaldo)%></b><small>Saldo awal + debet &minus; credit</small></div>
	</div>

	<div class="bv-card">
		<div class="bv-tabs" role="tablist" aria-label="Rincian bank">
			<button type="button" class="bv-tab" role="tab" id="t-pay" aria-controls="p-pay" aria-selected="true">Payment <span class="cnt"><%=bvPayCnt%></span></button>
			<button type="button" class="bv-tab" role="tab" id="t-rcv" aria-controls="p-rcv" aria-selected="false" tabindex="-1">Receive <span class="cnt"><%=bvRcvCnt%></span></button>
			<button type="button" class="bv-tab" role="tab" id="t-cf" aria-controls="p-cf" aria-selected="false" tabindex="-1">Cash Flow <span class="cnt"><%=bvCfCnt%></span></button>
		</div>

		<!-- Payment -->
		<section class="bv-panel" role="tabpanel" id="p-pay" aria-labelledby="t-pay">
<% If bvPayRows = "" Then %>
			<div class="bv-empty">Tidak ada payment di periode ini.</div>
<% Else %>
			<p class="bv-note">Total payment: <b><%=BvNum(bvTotPay)%></b> dari <%=bvPayCnt%> baris detail.</p>
			<table class="data" id="tblPay" data-fbt-list>
				<thead><tr>
					<th>No</th><th>No. Voucher</th><th>Date</th><th>Paid To</th><th class="num">Amount</th><th>Status</th><th data-nosort>Aksi</th><th>Keterangan</th><th>Remark Header</th><th>Period</th><th>Cek No</th><th>Classification</th><th>Akun</th><th>Posting</th><th>Input By</th><th>Edit By</th>
				</tr></thead>
				<tbody>
<%=bvPayRows%>
				</tbody>
			</table>
<% End If %>
		</section>

		<!-- Receive -->
		<section class="bv-panel" role="tabpanel" id="p-rcv" aria-labelledby="t-rcv" hidden>
<% If bvRcvRows = "" Then %>
			<div class="bv-empty">Tidak ada receive di periode ini.</div>
<% Else %>
			<p class="bv-note">Total receive: <b><%=BvNum(bvTotRcv)%></b> dari <%=bvRcvCnt%> baris detail.</p>
			<table class="data" id="tblRcv" data-fbt-list>
				<thead><tr>
					<th>No</th><th>No. Voucher</th><th>Date</th><th>Receive By</th><th class="num">Amount</th><th>Status</th><th data-nosort>Aksi</th>
					<th>Keterangan</th><th>Remark Header</th><th>Period</th><th>Cek No</th><th>Classification</th><th>Posting</th><th>Input By</th><th>Edit By</th>
				</tr></thead>
				<tbody>
<%=bvRcvRows%>
				</tbody>
			</table>
<% End If %>
		</section>

		<!-- Cash Flow -->
		<section class="bv-panel" role="tabpanel" id="p-cf" aria-labelledby="t-cf" hidden>
			<p class="bv-note">Saldo awal (opening balance): <b><%=BvNum(bvOpening)%></b>. Kolom Saldo adalah saldo berjalan sesuai urutan tanggal.</p>
			<table class="data" id="tblCf" data-fbt-list>
				<thead><tr>
					<th>No</th><th>Date</th><th>No. Voucher</th><th class="num">Debet</th><th class="num">Credit</th><th class="num">Saldo</th>
					<th>Paid For / Receive From</th><th>Reference</th><th>Remark</th>
				</tr></thead>
				<tbody>
<% If bvCfRows = "" Then %>
					<tr><td colspan="9" class="bv-empty">Tidak ada mutasi di periode ini.</td></tr>
<% Else %>
<%=bvCfRows%>
<% End If %>
					<tr><td colspan="3">Ending Balance</td><td class="num"><%=BvNum(bvDebet)%></td><td class="num"><%=BvNum(bvCredit)%></td><td class="num"><%=BvNum(bvSaldo)%></td><td colspan="3"></td></tr>
				</tbody>
			</table>
		</section>
	</div>
<% End If %>
</div>
</div>
</div>
<!--#include file="bottom.jsp"-->
<script src="/distexcel/tableToExcel.js"></script>
<script>
(function () {
	'use strict';
	var BANK = <%="""" & Replace(Replace(bvBank, "\", "\\"), """", "\""") & """"%>;
	var BANKNM = <%="""" & Replace(Replace(bvBankName & "", "\", "\\"), """", "\""") & """"%>;
	// klik teks yang terpotong untuk membuka / menutup isinya
	document.addEventListener('click', function (e) { var c = e.target.closest && e.target.closest('.bv td.t .clip'); if (c) c.classList.toggle('open'); });
	var tabs = Array.prototype.slice.call(document.querySelectorAll('.bv-tab'));
	function show(tab, focus) {
		tabs.forEach(function (t) {
			var on = t === tab, p = document.getElementById(t.getAttribute('aria-controls'));
			t.setAttribute('aria-selected', on ? 'true' : 'false');
			t.tabIndex = on ? 0 : -1;
			if (p) p.hidden = !on;
		});
		if (focus) tab.focus();
		try { sessionStorage.setItem('bv-tab', tab.id); } catch (e) { /* abaikan */ }
	}
	tabs.forEach(function (t, i) {
		t.addEventListener('click', function () { show(t); });
		t.addEventListener('keydown', function (e) {
			var k = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
			if (k) { e.preventDefault(); show(tabs[(i + k + tabs.length) % tabs.length], true); }
		});
	});
	try { var last = sessionStorage.getItem('bv-tab'); if (last && document.getElementById(last)) show(document.getElementById(last)); } catch (e) { /* abaikan */ }

	Array.prototype.forEach.call(document.querySelectorAll('[data-new]'), function (b) {
		b.addEventListener('click', function () {
			var rcv = b.getAttribute('data-new') === 'RCV';
			window.open('paymentvoucherdetail.Asp?txtModule=Finance%20and%20Accounting&txtTable=' + (rcv ? '01-receivevoucher&transtype=01-receive' : '02-paymentvoucher&transtype=02-payment') +
				'&banktype=0&bankcodeurl=' + encodeURIComponent(BANK), '_blank');
		});
	});

	// Export XLS: tabel Cash Flow, semua baris (tanpa toolbar / pagination / baris detail responsif)
	var exp = document.getElementById('bvExport');
	if (exp) exp.addEventListener('click', function () {
		var src = document.getElementById('tblCf');
		if (!src || !window.TableToExcel) return;
		var clone = src.cloneNode(true);
		Array.prototype.forEach.call(clone.querySelectorAll('.fbt-toolbar-row, .fbt-filter, .fbt-empty-row, .fbt-rdetail'), function (el) { el.parentNode.removeChild(el); });
		Array.prototype.forEach.call(clone.querySelectorAll('.fbt-page'), function (el) { el.classList.remove('fbt-page'); });
		clone.removeAttribute('id');
		TableToExcel.convert(clone, { name: 'CashFlow_' + BANK + ' - ' + BANKNM + '.xlsx', sheet: { name: 'CashFlow' } });
	});
})();
</script>
</body>
</html>
