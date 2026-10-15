<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>FBPlants Reporting -Daftar Hutang</title>
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
								<h2>LAPORAN DAFTAR HUTANG</h2>
							</div>
						</div>
					</th>
				</tr>
						
				
				<tr>
					<th align="left" colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">Sampai dengan: <%=MonthName(bulan) & "-" & tahun%></th>
				</tr>
						


	<%	
		

	%>
	<tr>
			<th>No</th>
		    <th>Company</th>
		    <th>FCBA</th>
			<th>Invoice No</th>
			<th>Invoice Supplier</th>
			<th>Tanggal Input Hutang</th>
			<th>Jatuh Tempo</th>
			<th>No PO</th>
			<th>Supplier</th>
			<th>Nilai Invoice</th>
			<th>Terbayar</th>
			<th>No Faktur Pajak</th>
			<th>User Entry</th>
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

		qq = "SELECT DISTINCT H.INVOICECODE,B.FCCOMPANYCODE,B.FCCODE,H.REFINVOICE,H.INVOICEDATE,H.POCODE,S.FCNAME,(H.TOTALAMOUNT)TOTALAMOUNT," & _ 
		" (H.PAYMENT)PAYMENT,H.TOP,H.INVOICEDATE + nvl(H.TOP,0) AS TGLTEMPO,H.NO_FAKTUR,H.FCENTRY " & _
		" FROM INVOICE H " & _
		" INNER JOIN INVOICEDETAIL D ON H.INVOICECODE=D.INVOICECODE AND H.FCBA=D.FCBA " & _
		" INNER JOIN SUPPLIER S ON S.FCCODE=H.SUPPLIERCODE AND S.FCBA=H.FCBA " & _
		" INNER JOIN BUSINESSUNIT B ON B.FCCODE=H.FCBA " & _
		" INNER JOIN BANKSTOCK BS ON BS.FCCODE=D.ITEMCODE " & _
		" WHERE EXTRACT(MONTH FROM H.INVOICEDATE)<=" & bulan & " AND EXTRACT(YEAR FROM H.INVOICEDATE)<=" & tahun & " AND H.TOTALAMOUNT >0" & _
		 filtercompany1 & _
		 filteritem & _
		 filterstatus & _
		" ORDER BY B.FCCOMPANYCODE,B.FCCODE, H.INVOICEDATE,H.INVOICECODE"
		'response.Write(qq)
		set rshutang = connStr.Execute(qq)

		dim ttlinv,ttlbayar
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
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("REFINVOICE")%></td>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("INVOICEDATE")%></td>	
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("TGLTEMPO")%></td>		
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("POCODE")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("FCNAME")%></td>			
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshutang("TOTALAMOUNT"),2)%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshutang("PAYMENT"),2)%></td>	
					
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("NO_FAKTUR")%></td>			
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshutang("FCENTRY")%></td>			
			
		</tr>
	  <%
	
		ttlinv = cdbl(ttlinv)+cdbl(rshutang("TOTALAMOUNT"))
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
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlinv, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlbayar, 2)%></strong></td>
		</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</BODY>
</HTML>	