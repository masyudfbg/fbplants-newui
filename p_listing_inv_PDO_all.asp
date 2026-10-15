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

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "')"
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

		sqlview1 = "select FCBA," & _
					" MAX(INVOICEBY) AS INVOICEBY, " & _
					" MAX(NOINVOICE) AS NOINVOICE, " & _
					" SUM(AMOUNT) AS AMT, " & _
					" SUM(payment) AS payment, " & _
					" SUM(ttl) AS TTLRKB, " & _
					" MAX(DESCRIPTION) AS DESCRIPTION from ( " & _
					" SELECT  " & _
					" 	o.FCBA, o.INVOICEBY, o.NOINVOICE, o.AMOUNT, o.PAYMENT, o.DESCRIPTION, 0 ttl " & _
					" from OPR_INV_PDO o  " & _
					" where o.TAHUN='" & tahun & "' and o.BULAN='" & bulan & "'  " & _
					" AND o.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') AND o.description like '%PDO UPAH PERIODE%'  " & _
					" group by o.fcba, o.INVOICEBY, o.NOINVOICE, o.AMOUNT, o.PAYMENT, o.DESCRIPTION, bappno " & _
					" union all " & _
					" select  " & _
					"	fcba, '' INVOICEBY, '' NOINVOICE, 0 TOTAL_AMOUNT, "& _
					"	0 TOTAL_PAYMENT, '' DESCRIPTION, "& _
					"	sum(ttl+TOTALAMOUNT_HK503) ttl "& _
					"	from (select a.fcba, sum(b.TOTALAMOUNT_HK) ttl, 0 TOTALAMOUNT_HK503 "& _
					"		from rkbhdr a "& _
					"		left join RKBUPAHDTL b on a.RKBNO = b.RKBNO "& _
					"		where a.PERIOD_YEAR = " & tahun & " and a.PERIOD_MONTH = '" & bulan & "'  "& _
					"		and a.fcba in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') "& _
					"		and a.DOCUMENT_CLASSIFICATION IN ('501', '502') "& _
					"		group by a.fcba "& _
					"		union all "& _
					"		select  "& _
					" 		a.fcba, 0 ttl, " & _
					" 		sum(c.finalamount) TOTALAMOUNT_HK503  " & _
					" 		from bkmhdr a  " & _
					" 		left join bkmupahdtl b on a.BKMNO = b.BKMNO  " & _
					" 		left join bkmproses c on b.BKMNO = c.BKMNO  " & _
					" 		where b.actcode in ('6310101','6210101') and a.bkmdate = TO_DATE('" & bulan & "/" & tahun & "','MM/YYYY') and a.fcba in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') " & _
					" 		group by a.fcba " & _
					" 	)  " & _
					" group by fcba " & _
				") group by fcba"

		sqlview2 = "select FCBA," & _
					" MAX(INVOICEBY) AS INVOICEBY, " & _
					" MAX(NOINVOICE) AS NOINVOICE, " & _
					" SUM(AMOUNT) AS AMT, " & _
					" SUM(payment) AS payment, " & _
					" SUM(ttl) AS TTLRKB, " & _
					" MAX(DESCRIPTION) AS DESCRIPTION from ( " & _
					" SELECT  " & _
					" 	o.FCBA, o.INVOICEBY, o.NOINVOICE, o.AMOUNT, o.PAYMENT, o.DESCRIPTION, 0 ttl " & _
					" from OPR_INV_PDO o  " & _
					" where o.TAHUN='" & tahun & "' and o.BULAN='" & bulan & "'  " & _
					" AND o.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') AND o.description like '%PDO CONTRACTOR LOKAL PERIODE%'  " & _
					" group by o.fcba, o.INVOICEBY, o.NOINVOICE, o.AMOUNT, o.PAYMENT, o.DESCRIPTION, bappno " & _
					" union all " & _
					" select  " & _
					"	fcba, '' INVOICEBY, '' NOINVOICE, 0 TOTAL_AMOUNT, "& _
					"	0 TOTAL_PAYMENT, '' DESCRIPTION, "& _
					"	sum(ttl+TOTALAMOUNT_HK503) ttl "& _
					"	from (select a.fcba, sum(b.qty*b.harga) ttl, 0 TOTALAMOUNT_HK503 "& _
					"		from rkbhdr a "& _
					"		left join rkbwipdtl b on a.RKBNO = b.RKBNO "& _
					"		where a.PERIOD_YEAR = " & tahun & " and a.PERIOD_MONTH = '" & bulan & "'  "& _
					"		and a.fcba in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') "& _
					"		and a.DOCUMENT_CLASSIFICATION = '504' "& _
					"		group by a.fcba "& _
					" 	)  " & _
					" group by fcba " & _
				") where invoiceby is not null group by fcba"

		' response.Write sqlview1

		' sql503 = " select " & _
		' 			" CASE WHEN sum(c.finalamount) > 0 THEN sum(c.finalamount) ELSE 0 END AS TOTALAMOUNT_HK503  " & _
		' 			" from bkmhdr a left join bkmupahdtl b on a.BKMNO = b.BKMNO  " & _
		' 			" left join bkmproses c on b.BKMNO = c.BKMNO " & _
		' 			" where b.actcode in ('6310101','6210101') and a.bkmdate = TO_DATE('"&bulan&"/"&tahun&"','MM/YYYY') AND a.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "')"
		' set res = connStr.Execute(sql503)
		' response.write sqlview2

		Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open sqlview1, connStr

		Set rsCat2 = Server.CreateObject("ADODB.Recordset")
			rsCat2.CursorType = 3
			rsCat2.CursorLocation = 3
			rsCat2.Open sqlview2, connStr
			

	%>

	<body>
		<div id="wrapper">
			<table border="0" width="100%" id="product-table" class="xthf-site">
				<thead>
					<tr>
						
						<th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none">
							<div id="header-wrapper">
								<div id="head_judul" style="float:left; text-align:left;">
									<h4>PT. <%=CompanyName%></h4>
									<% if Not trim(fcba) = "-" then %>
										<h5>Estate: <%=fcba%></h5>
									<% end if %>
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
					<tr>
						<th align="left" colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none;"><h2>1. PDO UPAH PERIODE</h2></th>
					</tr>
					<%NO=1%>
					<tr>
						<th style="width: 3%;">No</th>
						<th style="width: 15%;">Business Unit</th>
						<th style="width: 10%;">Invoice By</th>
						<th style="width: 15%;">No Invoice</th>
						<th style="width: 10%;">Amount PDO</th>
						<th style="width: 10%;">Amount RKB</th>
						<th style="width: 10%;">Selisih (PDO | RKB)</th>
						<th style="width: 15%;">Remark</th>
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
					%>
					<tr>
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>			
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("INVOICEBY")%></td>				
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("noinvoice")%></td>	
						<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amt"),2)%></td>			
						<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("TTLRKB"),2)%></td>
						<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("amt"))-cdbl(rscat("TTLRKB")), 2)%></td>	
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("description")%></td>
					</tr>
					<%
						ttlamountshi = cdbl(ttlamountshi)+cdbl(rscat("amt"))
						ttlpaymentshi = cdbl(ttlpaymentshi)+cdbl(rscat("TTLRKB"))
						ttlRkb = cdbl(ttlRkb)+cdbl(rscat("amt"))-cdbl(rscat("TTLRKB"))
						rsCat.MoveNext

						no=no+1
						Wend
						rsCat.close()
						rsCat=null
						END IF
					%>
					<tr bgcolor="#F4F4A0">
						<td align="Center" style="<%=cmdstyle%>" colspan= "4"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlamountshi, 2)%></strong></td>		
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlpaymentshi, 2)%></strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlRkb, 2)%></strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong></strong></td>
					</tr>
				</tbody>
			</table><br><br>
			
			<table border="0" width="100%" id="product-table" class="xthf-site">
				<thead>
					<tr>
						<th align="left" colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none;"><h2>2. PDO CONTRACTOR LOKAL</h2></th>
					</tr>
					<%NO=1%>
					<tr>
						<th style="width: 3%;">No</th>
						<th style="width: 15%;">Business Unit</th>
						<th style="width: 10%;">Invoice By</th>
						<th style="width: 15%;">No Invoice</th>
						<th style="width: 10%;">Amount PDO</th>
						<th style="width: 10%;">Amount RKB</th>
						<th style="width: 10%;">Selisih (PDO | RKB)</th>
						<th style="width: 15%;">Remark</th>
					</tr>
				</thead>
				<tbody>
					<%
						if Not rsCat2.eof then
							no=1
							Dim ttlamount2,ttlamountshi2,ttlpayment2,ttlpaymentshi2,ttlRkb2

								ttlamount2 = 0
								ttlamountshi2 = 0
								ttlRkb2 = 0
								ttlpayment2 = 0
								ttlpaymentshi2 = 0
						While Not rsCat2.EOF
							if FormatNumber(cdbl(rsCat2("amt"))-cdbl(rsCat2("TTLRKB")), 0) = FormatNumber(rsCat2("amt"),0) then
								hasilRkb = 0
							else
								hasilRkb = FormatNumber(cdbl(rsCat2("amt"))-cdbl(rsCat2("TTLRKB")), 0)
							end if
					%>
					<tr>
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat2("fcba")%></td>			
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat2("invoiceby")%></td>				
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat2("noinvoice")%></td>	
						<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat2("amt"),2)%></td>			
						<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat2("TTLRKB"),2)%></td>
						<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1">
							<%=hasilRkb%>
						</td>	
						<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat2("description")%></td>
					</tr>
					<%
						ttlamountshi2 = cdbl(ttlamountshi2)+cdbl(rsCat2("amt"))
						ttlpaymentshi2 = cdbl(ttlpaymentshi2)+cdbl(rsCat2("TTLRKB"))
						ttlRkb2 = cdbl(ttlRkb2)+cdbl(ttlRkb1)
						rsCat2.MoveNext

						no=no+1
						Wend
						rsCat2.close()
						rsCat2=null
						END IF
					%>
					<tr bgcolor="#F4F4A0">
						<td align="Center" style="<%=cmdstyle%>" colspan= "4"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlamountshi2, 2)%></strong></td>		
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlpaymentshi2, 2)%></strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlRkb2, 2)%></strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong></strong></td>
					</tr>
					<tr bgcolor="#CCCCCC">
						<td align="Center" style="border-right: 1 solid #000000; border-bottom: 1 solid #000000" colspan= "4"><font face="Verdana" size="1"><strong>GRAND TOTAL</strong></td>						
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(ttlamountshi)+cdbl(ttlamountshi2), 2)%></strong></td>		
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(ttlpaymentshi)+cdbl(ttlpaymentshi2), 2)%></strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(ttlRkb)+cdbl(ttlRkb2), 2)%></strong></td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong></strong></td>
					</tr>
				</tbody>
			</table>
		<!--#include file="Bottom_Report.asp"---------->
		</div>
	</body>
</html>	