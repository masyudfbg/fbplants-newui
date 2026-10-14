<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting -LHD Listing</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
function() {
new xTableHeaderFixed('xthf-site', window);
}, false
);
</script>
<style type="text/css">
@media all {
#footer_note {width:65%;margin:auto;margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
}
@media print
{
#footer_note {margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.printhide{display:none;}
@page { 
 size:Letter landscape;
}
}

</style>
</head>
	<%
		Dim Columnx : Columnx = 0
		Const adOpenForwardOnly = 0
		Const adLockReadOnly = 1
		Const adCmdTableDirect = &H0200
		Const adUseClient = 3

		Dim tbl,doc,filterdiv,filterlhd,pcompany,fcba,actfrom,bkmto,divisi,bulansplitfrom,tglsplitfrom,thnsplitfrom,tglfrom,blnthn,q
		dim sqldate,rsdate,startdate,sqlsbi,rssbi,startdatesbi,periode,t1,filterprd,filtercomp
		cmdstyle="border-right: 1 solid #000000; border-bottom: 1 solid #000000"
		cmdstyleTop="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-top: 1 solid #000000"
		cmdstyleLeft="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-Left: 1 solid #000000"

		pcompany 	= Request("pcompany")
		fcba 		= Request("pfcba")
		bulan   	= Request("bulan")
		tahun   	= Request("tahun")

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
		' response.write "FROM" & startdate
		set rsCat = connStr.Execute(q)

		if not rsCat.eof then
			fcbaname = rsCat("fcName")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("fcType")
		end if

		bulansplitfrom  = left(bkmto,2)
		tglsplitfrom    = mid(bkmto,4,2)
		thnsplitfrom    = right(bkmto,4)
		tglfrom     	= bulansplitfrom & "/" & tglsplitfrom & "/" & thnsplitfrom
		blnthn 			= bulansplitfrom&thnsplitfrom

		if trim(fcba) = "-" then 
			filtercomp = " AND o.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') "
		else
			filtercomp = " AND o.FCBA ='" & fcba & "' "
		end if

		sqlview1 = " SELECT o.FCBA, o.INVOICEBY, o.NOINVOICE, o.AMOUNT, o.PAYMENT, o.DESCRIPTION, " & _ 
				   " case " & _
				   "	when REPLACE(REGEXP_REPLACE(o.BAPPNO, '([^|]+)', '''\1'''), '|', ',') is null then '0'  " & _
				   "	else REPLACE(REGEXP_REPLACE(o.BAPPNO, '([^|]+)', '''\1'''), '|', ',') " & _
				   " end BAPPNO  " & _ 
				   " from OPR_INV_PDO o " & _
				   " where o.TAHUN='" & tahun & "' and o.BULAN='" & bulan & "' " & _
				   " AND o.FCBA ='" & fcba & "'"  & _
				   " ORDER BY o.FCBA,o.INVOICEBY,o.NOINVOICE,o.AMOUNT,o.PAYMENT,o.DESCRIPTION"

		' response.Write sqlview1

		if Len(bulan) = 1 Then
			bln503 = "0"&bulan
		else 
			bln503 = bulan
		end if

		sql503 = " select " & _
					" CASE WHEN sum(c.finalamount) > 0 THEN sum(c.finalamount) ELSE 0 END AS TOTALAMOUNT_HK503  " & _
					" from bkmhdr a left join bkmupahdtl b on a.BKMNO = b.BKMNO  " & _
					" left join bkmproses c on b.BKMNO = c.BKMNO " & _
					" where b.actcode in ('6310101','6210101') and TO_CHAR(a.bkmdate, 'MMYYYY') = '"&bln503&tahun&"' " & _
					" and a.fcba = '"&fcba&"'"
		' set res = connStr.Execute(sql503)
		' response.write sql503

		Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open sqlview1, connStr

	%>

	<body>
		<div id="wrapper">
			<!--#include file="print_excel_blob.asp"-->
			<table border="0" width="100%" id="product-table" class="xthf-site">
				<thead>
					<tr>
						
						<th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none">
							<div id="header-wrapper">
								<div id="head_judul" style="float:left; text-align:left;">
									<h4>PT. <%=CompanyName%></h4>

									<h5>Estate: <%=fcba%></h5>
								</div>
								
							</div>
						</th>
					</tr>
					<tr>
						<th align="left" colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">LAPORAN INVOICE PDO </th>
					</tr>
					<tr>
						<th align="left" colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">Periode: <%=bulan & "-" & tahun%></th>
					</tr>
					<%NO=1%>
					<tr>
						<th>No</th>
						<th>Business Unit</th>
						<th>Invoice By</th>
						<th>No Invoice</th>
						<th>Amount</th>
						<th>Payment</th>
						<th>RKB</th>
						<th>Remark</th>
					</tr>
				</thead>
				<tbody>
					<%
						if Not rsCat.eof then
							no=1
							Dim ttlamount,ttlamountshi,ttlpayment,ttlpaymentshi,ttlRkb

								ttlamount = 0
								ttlamountshi = 0
								ttlRkb = 0
								ttlpayment = 0
								ttlpaymentshi = 0
						While Not rsCat.EOF

							if Len(bulan) = 1 Then
								bln = "0"&bulan
							else 
								bln = bulan
							end if

							if Request("pregion") = "" Then
								region = ""
							else 
								region = "and d.region = '"&Request("pregion")&"' "
							end if

							if rsCat("bappno") = "0" Then
								invcode = "invoicecode = '0' "
							else 
								invcode = "invoicecode in ("&rsCat("bappno")&")"
							end if

							if InStr(1, rscat("description"), "PDO UPAH PERIODE", vbTextCompare) > 0 Then 
								slct = "CASE WHEN sum(b.TOTALAMOUNT_HK) > 0 THEN sum(b.TOTALAMOUNT_HK) "
								tbljoin = "left join RKBUPAHDTL b on a.RKBNO = b.RKBNO left join businessunit c on a.fcba = c.fccode LEFT JOIN company d ON c.fccompanycode = d.companycode"
								subQuery = "SELECT U.ACTCODE FROM BKMUPAHDTL U " & _
										   " WHERE BKMNO IN(SELECT BKMNO FROM BKMHDR WHERE TO_CHAR(BKMDATE,'MMYYYY')='"&bln&tahun&"') " & _
										   " AND FCBA='"&fcba&"' " & _
										   " GROUP BY ACTCODE" 
								where = "a.DOCUMENT_CLASSIFICATION IN ('501', '502', '503')"
							End If

							if InStr(1, rscat("description"), "PDO CONTRACTOR", vbTextCompare) > 0 Then 
								slct = "CASE WHEN sum(b.qty*b.harga) > 0 THEN sum(b.qty*b.harga) "
								tbljoin = "left join rkbwipdtl b on a.RKBNO = b.RKBNO  left join businessunit c on a.fcba = c.fccode LEFT JOIN company d ON c.fccompanycode = d.companycode"
								subQuery = " select b.JOBCODE from workinprogressfinal a " & _
										   " left join workinprogressfinal_new b on a.AGREEMENTCODE = b.AGREEMENTCODE " & _
										   " where b.agreementcode in ( " & _
										   " select agreementcode from CONTRACTORINVOICEDTL  " & _
										   " where invoicecode = '"&invcode&"' group by agreementcode " & _
										   "  ) AND b.YEAR = "&tahun&" AND b.MONTH = '"&bulan&"' " & _
										   " GROUP BY b.JOBCODE "
								where = "a.DOCUMENT_CLASSIFICATION = '504'"
							End If
					%>
					<tr>
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>			
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoiceby")%></td>				
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("noinvoice")%></td>	
						<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount"),0)%></td>			
						<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("payment"),0)%></td>
						<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1">
							<%
								getRKB = " select " & _
										 slct & _
										 " ELSE 0 " & _
										 " END AS TOTALAMOUNT_HK  " & _
										 " from RKBHDR a " & _
										 tbljoin & _
										 " where a.fcba = '"&fcba&"' and a.PERIOD_YEAR = "&tahun&" AND a.PERIOD_MONTH = '"&bulan&"'" & _
										 region & _
										 " AND b.ACTCODE IN  " & _
										 " ( " & _
										 subQuery & _
										 " ) AND "& where &""
								' response.write getRKB
								set rsRKB = connStr.Execute(getRKB)

								if InStr(1, rscat("description"), "PDO UPAH PERIODE", vbTextCompare) > 0 Then 
									rkbamt = FormatNumber(rsRKB("TOTALAMOUNT_HK"), 0)
								End If

								if InStr(1, rscat("description"), "PDO CONTRACTOR LOKAL", vbTextCompare) > 0 Then 
									rkbamt = FormatNumber(rsRKB("TOTALAMOUNT_HK"), 0)
								End If
							%>
							<%= FormatNumber(rkbamt, 0) %>
						</td>	
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("description")%></td>
					</tr>
					<%
						ttlamountshi = cdbl(ttlamountshi)+cdbl(rscat("amount"))
						ttlpaymentshi = cdbl(ttlpaymentshi)+cdbl(rscat("payment"))
						ttlRkb = cdbl(ttlRkb)+cdbl(rkbamt)
						rsCat.MoveNext

						no=no+1
						Wend
						rsCat.close()
						rsCat=null
						END IF
					%>
					<tr bgcolor="#F4F4A0">
						<td align="Center" style="<%=cmdstyle%>" colspan= "4"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlamountshi, 0)%></strong></td>		
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlpaymentshi, 0)%></strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlRkb, 0)%></strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong></strong></td>
					</tr>
				</tbody>
			</table>
		<!--#include file="Bottom_Report.asp"---------->
		</div>
	</body>
</html>	