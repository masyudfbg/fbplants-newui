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
dim sqldate,rsdate,startdate,sqlsbi,rssbi,startdatesbi,periode,t1,filterprd
cmdstyle="border-right: 1 solid #000000; border-bottom: 1 solid #000000"
cmdstyleTop="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-top: 1 solid #000000"
cmdstyleLeft="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-Left: 1 solid #000000"




   pcompany 	= Request("pcompany")
    fcba 		= Request("pfcba")
    actfrom 	= Request("bkmfrom")
    bkmto   	= Request("bkmto")
    divisi   	= Request("divisi")
    t1   	    = Request("t1")
	bulan   	= Request("bulan")
	tahun   	= Request("tahun")
	
	



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

	bulansplitfrom  = left(bkmto,2)
	tglsplitfrom    = mid(bkmto,4,2)
	thnsplitfrom    = right(bkmto,4)
	tglfrom     	= bulansplitfrom & "/" & tglsplitfrom & "/" & thnsplitfrom
	blnthn 			= bulansplitfrom&thnsplitfrom


	' sqldate		= " select to_char(start_date,'mm/dd/yyyy') start_date from empalldedstatus where fcba='" & fcba & "' and to_date('" & bkmto & "','mm/dd/yyyy') between  start_date and end_date"
	' set rsdate 	= connStr.Execute(sqldate)
	' if not rsdate.eof then
	' 	startdate = rsdate("start_date")
	' end if

	sqlexec = "BEGIN " & _ 
              "opr_pdo_bulanan_api.Proses('" & bulan & "','" & tahun & "','" & fcba & "');" & _
			  "END; " 
    connStr.Execute(sqlexec)


	if trim(divisi) <> "all" then
		filterdiv = " AND t.sectioncode='" & divisi & "'"
	end if


	if trim(fcba) = "-" then 
		filterlhd = " WHERE t.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') "
	else
	    filterlhd = " WHERE t.FCBA ='" & fcba & "' "
	end if

if t1 <> "1" then 
' Detail
sqlview1 = "SELECT t.fcba,t.fccode,t.fcname,t.sectioncode,t.fieldcode,t.fieldarea,t.fieldstatus, " & _ 
" nvl(sum(t.fisikshi),0)fisikshi, " & _ 
" nvl(sum(t.unitshi),0)unitshi, " & _ 
" nvl(sum(t.hkshi),0) hkshi, " & _ 
" nvl(sum(t.rphkshi),0)rphkshi " & _ 
" FROM ( " & _ 
"	SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus, " & _ 
"	SUM(o.output) fisikshi,SUM(o.unit) unitshi, " & _ 
"	SUM(o.hk) hkshi, round(SUM(o.amount),0) rphkshi " & _ 
"	FROM OPR_PDO_BULANAN o " & _ 
"	where o.source='0.PDO' " & _ 
"	AND o.PAYROLL_MONTH='" & bulan & "' AND o.PAYROLL_YEAR ='" & tahun & "' " & _
"	GROUP BY o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus " & _ 
"	) t " & _ 
  filterlhd & _
  filterdiv & _
" GROUP BY t.fcba,t.sectioncode,t.fieldcode,t.fieldarea,t.fieldstatus,t.fccode,t.fcname " & _ 
" ORDER BY t.fcba,t.fccode,t.sectioncode,t.fieldcode,t.fieldarea,t.fieldstatus"
else
'rekap
sqlview1 = "SELECT t.fcba,t.fccode,t.fcname,'-' sectioncode,'-' fieldcode,0 fieldarea,'-' fieldstatus, " & _ 
" nvl(sum(t.fisikshi),0)fisikshi, " & _ 
" nvl(sum(t.unitshi),0)unitshi, " & _ 
" nvl(sum(t.hkshi),0) hkshi, " & _ 
" nvl(sum(t.rphkshi),0)rphkshi " & _ 
" FROM ( " & _ 
"	SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus, " & _ 
"	SUM(o.output) fisikshi,SUM(o.unit) unitshi, " & _ 
"	SUM(o.hk) hkshi, round(SUM(o.amount),0) rphkshi " & _ 
"	FROM OPR_PDO_BULANAN o " & _ 
"	where o.source='0.PDO' " & _ 
"	AND o.PAYROLL_MONTH='" & bulan & "' AND o.PAYROLL_YEAR ='" & tahun & "' " & _
"	GROUP BY o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus " & _ 
"	) t " & _ 
  filterlhd & _
  filterdiv & _
" GROUP BY t.fcba,t.fccode,t.fcname " & _ 
" ORDER BY t.fcba,t.fccode"
end if
'response.Write(sqlview1)

	Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open sqlview1, connstr


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
						
						<th align="left" colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">LAPORAN BULANAN DIVISI / PDO </th>
						
				</tr>
				<tr>
						
						<th align="left" colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">Periode: <%=bulan & "-" & tahun%></th>
				</tr>
						


	<%	
	
	NO=1%>
	<tr>
			 <th rowspan="3">No</th>
		    <th rowspan="3">FCBA</th>
			<th rowspan="3">Section</th>
			<th rowspan="3">Blok</th>
			<th rowspan="3">Luas</th>
			<th rowspan="3">Status</th>
			<th rowspan="3">Activity</th>
			<th colspan="8">UPAH</th>
			
	</tr>
	<tr>
			<th colspan="">Ha</th>			
			<th colspan="">Qty</th>
			<th colspan="">HK</th>
			<th colspan="">Rp. HK</th>
			
		</tr>
		
</thead>
<tbody>
	<%
	
	   if Not rsCat.eof then
		
			no=1
		
		Dim ttlhk,ttlhkshi,ttlhksbi,ttlfisik,ttlunit, ttlrphkhi,ttlrphkshi,ttlrphksbi
		dim ttlfisikshi,ttlfisiksbi,ttlunitshi,ttlunitsbi
		dim ttlfisikbrhi,ttlfisikbrshi,ttlfisikbrsbi,ttlunitbrhi,ttlunitbrshi,ttlunitbrsbi
		dim ttlrpbrhi,ttlrpbrshi,ttlrpbrsbi
		dim ttlqtybahanhi,ttlqtybahanshi,ttlqtybahansbi,ttlbiayabahanhi,ttlbiayabahanshi,ttlbiayabahansbi
		dim ttlqtyvrahi,ttlqtyvrashi,ttlqtyvrasbi,ttlbiayavrahi,ttlbiayavrashi,ttlbiayavrasbi
		dim ttlrupiahhi,ttlrupiahshi,ttlrupiahsbi
		dim ttlgrndrupiahhi,ttlgrndrupiahshi,ttlgrndrupiahsbi

			ttlfisik = 0
			ttlfisikshi = 0
			
			ttlunit = 0
			ttlunitshi = 0
			
			ttlhk = 0
			ttlhkshi = 0
		
			ttlrphkhi = 0
			ttlrphkshi = 0
		
			
			 ttlgrndrupiahhi = 0
			 ttlgrndrupiahshi = 0
		

		While Not rsCat.EOF
		
			%>
		<tr>
			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sectioncode")%></td>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fieldcode")%></td>				
					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fieldarea")%></td>		
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fieldstatus")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fccode") & "-" & rscat("fcname")%></td>			
					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("fisikshi"),2)%></td>					
		
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("unitshi"),2)%></td>	
			
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("hkshi"),2)%></td>			
			
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("rphkshi"),0)%></td>	
			
		</tr>
	  <%
	  
	
		ttlfisikshi = cdbl(ttlfisikshi)+cdbl(rscat("fisikshi"))
		
	
		ttlunitshi = cdbl(ttlunit)+cdbl(rscat("unitshi"))
	
	
		ttlhkshi = cdbl(ttlhkshi)+cdbl(rscat("hkshi"))
	
	
		ttlrphkshi = cdbl(ttlrphkshi)+cdbl(rscat("rphkshi"))	

	
		ttlgrndrupiahshi = cdbl(ttlgrndrupiahshi) + cdbl(ttlrupiahshi)
	
	  rsCat.MoveNext

	  no=no+1
	Wend
	rsCat.close()
	rsCat=null
	END IF
%>

		<tr bgcolor="#F4F4A0">
			<td align="Center" style="<%=cmdstyle%>" colspan= "7"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
			
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlfisikshi, 0)%></strong></td>
			
		
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlunitshi, 0)%></strong></td>
		
			
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlhkshi, 0)%></strong></td> 
		
		
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlrphkshi, 0)%></strong></td>
			
			
		</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</BODY>
</HTML>	