<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>FBPlants Reporting -Daftar Umur Hutang</title>
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

Dim tbl,doc,filterdiv,filterlhd,pcompany,fcba,actfrom,bkmto,divisi,bulansplitfrom,itemgrup,tglsplitfrom,thnsplitfrom,tglfrom,blnthn,q
dim sqldate,rsdate,startdate,sqlsbi,rssbi,startdatesbi,periode,t1,filterprd
dim rshutang,qq,filterstatus
cmdstyle="border-right: 1 solid #000000; border-bottom: 1 solid #000000"
cmdstyleTop="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-top: 1 solid #000000"
cmdstyleLeft="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-Left: 1 solid #000000"


    pcompany 	= Request("pcompany")
    region 		= Request("region")
    fcba 		= Request("pfcba")
    actfrom 	= Request("bkmfrom")
    bkmto   	= Request("bkmto")
    divisi   	= Request("divisi")
    t1   	    = Request("t1")
	bulan   	= Request("bulan")
	tahun   	= Request("tahun")
	statusbyr   = Request("statusbyr")
	itemgrup   = Request("itemgrup")
	

	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
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
							<div id="subheader" class="centered" style="width:50%; margin:auto;">
								<h2>LAPORAN DAFTAR UMUR HUTANG</h2>
							</div>
						</div>
					</th>
				</tr>
						
				
				<tr>
					<th align="left" colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">Periode: <%=bulan & "-" & tahun%></th>
				</tr>
						


	<%	
		

	%>
	<tr>
			<th rowspan="2" >No</th>
		    <th rowspan="2" >Company</th>
		    <th rowspan="2" >FCBA</th>
			<th rowspan="2" >Invoice No</th>
			<th rowspan="2" >Supplier</th>
			<th rowspan="2" >Tgl Input Hutang</th>
			<th rowspan="2" >Jatuh Tempo</th>
			<th rowspan="2" >No PO</th>
			<th rowspan="2" >Nilai Invoice</th>
			<th rowspan="2" >Dibayar</th>
			<th colspan="4" >Sudah Jatuh Tempo</th>
			<th rowspan="2" >Hari Outstanding</th>
	</tr>
    <tr>
                    <th >1-30Hari</th>
                    <th >31-60Hari</th>
					<th >61-90Hari</th>
                    <th >Over 90 Hari</th>
                
    </tr>
</thead>
<tbody>
<%

		
		if trim(statusbyr) = "b" then
			filterstatus = " AND H.PAYMENT=0"
		end if
		if trim(statusbyr) = "s" then
			filterstatus = " AND H.PAYMENT>0"
		end if

		if trim(itemgrup) = "101" then
			filteritem = " AND D.ITEMCODE in('101.1000.02.004','101.1000.004')"
		end if
		if trim(itemgrup) = "103" then
			filteritem = " AND BS.STOCKGROUPCODE ='103'"
		end if
		if trim(itemgrup) = "104" then
			filteritem = " AND BS.STOCKGROUPCODE ='104'"
		end if
		' if trim(fcba) = "-" then 
		' 	filterfcba = " AND H.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') "
		' else
		' 	filterfcba = " AND H.FCBA ='" & fcba & "' "
		' end if

		if trim(pcompany)="-" THEN
			filtercompany1 = " AND b.isactive='TRUE' AND b.region ='" & region & "' "
		else
		if trim(fcba) = "-" then 
			filtercompany1 = " AND b.isactive='TRUE' AND b.fccompanycode ='" & pcompany & "' "
		else 
			filtercompany1 = " AND b.isactive='TRUE' AND b.fccode ='" & fcba & "' "
		end if
    	end if

		qq = "SELECT " & _
		" DISTINCT H.INVOICECODE," & _
		" B.FCCOMPANYCODE," & _
		" B.FCCODE," & _
		" H.REFINVOICE," & _
		" H.INVOICEDATE," & _
		" H.POCODE," & _
		" S.FCNAME," & _
		" H.TOTALAMOUNT," & _
		" (H.PAYMENT)PAYMENT," & _
		" H.TOP,H.INVOICEDATE + nvl(H.TOP,0) AS TGLTEMPO," & _
		" MAX(CASE " & _
		"	WHEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0)) BETWEEN 1 AND 30 AND H.TOTALAMOUNT - H.PAYMENT > 0  THEN H.TOTALAMOUNT - H.PAYMENT " & _
		"	ELSE 0 " & _
		" END) AS HUTANG_30_HARI," & _
		" MAX(CASE " & _
		"	WHEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0)) BETWEEN 31 AND 60 AND H.TOTALAMOUNT - H.PAYMENT > 0 THEN H.TOTALAMOUNT - H.PAYMENT " & _
		"	ELSE 0 " & _
		" END) AS HUTANG_60_HARI," & _
		" MAX(CASE " & _ 
		"	WHEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0)) BETWEEN 61 AND 90 AND H.TOTALAMOUNT - H.PAYMENT > 0 THEN H.TOTALAMOUNT - H.PAYMENT " & _
		"	ELSE 0 " & _
		"END) AS HUTANG_90_HARI, " & _
		"MAX(CASE " & _
		"	WHEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0)) > 90 AND H.TOTALAMOUNT - H.PAYMENT > 0 THEN H.TOTALAMOUNT - H.PAYMENT " & _
		"	ELSE 0 " & _
		"END) AS HUTANG_LEBIH_90_HARI, " & _
	" CASE " & _
    "    WHEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0)) BETWEEN 1 AND 30 THEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0))" & _
    "    WHEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0)) BETWEEN 31 AND 60 THEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0))" & _
    "    WHEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0)) BETWEEN 61 AND 90 THEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0))" & _
    "    WHEN SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0)) > 90 THEN  SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0))" & _
    "    ELSE 0 " & _
    " END AS STATUS_JATUH_TEMPO " & _
	"FROM INVOICE H " & _
	"INNER JOIN INVOICEDETAIL D ON H.INVOICECODE = D.INVOICECODE AND H.FCBA = D.FCBA " & _
	"INNER JOIN SUPPLIER S ON S.FCCODE = H.SUPPLIERCODE AND S.FCBA = H.FCBA " & _
	"INNER JOIN BUSINESSUNIT B ON B.FCCODE = H.FCBA " & _
	"INNER JOIN BANKSTOCK BS ON BS.FCCODE = D.ITEMCODE " & _
	"WHERE " & _
	"	EXTRACT(MONTH FROM H.INVOICEDATE) <= '" & bulan & "' " & _
	"	AND EXTRACT(YEAR FROM H.INVOICEDATE) <= '" & tahun & "' " & _
	"	AND H.TOTALAMOUNT > 0  " & _
	"	AND B.ISACTIVE = 'TRUE'  " & _
	 filtercompany1 & _
	 filteritem & _
	 filterstatus & _
	"	AND (SYSDATE - (H.INVOICEDATE + NVL(H.TOP,0))) > 0 " & _
	"	AND H.TOTALAMOUNT-H.PAYMENT>0 " & _
	"GROUP BY " & _
	"	H.INVOICECODE,B.FCCOMPANYCODE, B.FCCODE,  H.REFINVOICE, " & _
	"	H.INVOICEDATE, " & _
	"	H.POCODE, " & _
	"	S.FCNAME, " & _
	"	H.TOTALAMOUNT, " & _
	"	H.PAYMENT, " & _
	"	H.TOP, " & _
	"	H.INVOICEDATE + nvl(H.TOP,0) " & _
	"ORDER BY " & _
	"	H.INVOICECODE,B.FCCOMPANYCODE, B.FCCODE"
		'response.Write(qq)
		set rshutang = connStr.Execute(qq)

		dim ttlinv,ttlinv30,ttlinv60,ttlinv90,ttlinvleih90,ttlbayar
		ttlinv = 0
		ttlbayar = 0
	   if Not rshutang.eof then
		no=1
		While Not rshutang.EOF
		
			%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("FCCOMPANYCODE")%></td>			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("FCCODE")%></td>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("INVOICECODE")%></td>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("FCNAME")%></td>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("INVOICEDATE")%></td>		
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("TGLTEMPO")%></td>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("POCODE")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshutang("TOTALAMOUNT"),0)%></td>	
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshutang("PAYMENT"),0)%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshutang("HUTANG_30_HARI"),0)%></td>							
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshutang("HUTANG_60_HARI"),0)%></td>							
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshutang("HUTANG_90_HARI"),0)%></td>							
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshutang("HUTANG_LEBIH_90_HARI"),0)%></td>							
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshutang("STATUS_JATUH_TEMPO"),0) & " Hari "%></td>	
			
			
		</tr>
	  <%
	
		ttlinv = cdbl(ttlinv)+cdbl(rshutang("TOTALAMOUNT"))
		ttlinv30 = cdbl(ttlinv30)+cdbl(rshutang("HUTANG_30_HARI"))
		ttlinv60 =  cdbl(ttlinv60)+cdbl(rshutang("HUTANG_60_HARI"))
		ttlinv90 = cdbl(ttlinv90)+cdbl(rshutang("HUTANG_90_HARI"))
		ttlinvleih90 = cdbl(ttlinvleih90)+cdbl(rshutang("HUTANG_LEBIH_90_HARI"))
		ttlbayar = cdbl(ttlbayar)+cdbl(rshutang("PAYMENT"))
	  rshutang.MoveNext

	  no=no+1
	Wend
	rshutang.close()
	rshutang=null
	END IF
%>

		<tr bgcolor="#F4F4A0">
			<td align="Center" style="<%=cmdstyle%>" colspan= "8"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlinv, 0)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlbayar, 0)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlinv30, 0)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlinv60, 0)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlinv90, 0)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlinvleih90, 0)%></strong></td>
		</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</BODY>
</HTML>	