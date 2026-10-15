<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Mill Production Rekap</title>

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
@media print
{
.printhide{display:none;}
}

.tebal{
	font-weight:bold;
}

.ratakiri{
	text-align:left;
}
.ratakanan{
	text-align:right;
}
.ratatengah{
	text-align:center;
}
.lebarket{
	width:780px;
}

</style>
</head>

<%



Dim tbl,doc,nday,txtSQL,rsTBS,bulanx,tahunx,subtotaltbs,areaid,areaname,noline,TAHUNBULAN
Dim ttl_day_1,ttl_day_2,ttl_day_3,ttl_day_4,ttl_day_5,ttl_day_6,ttl_day_7,ttl_day_8,ttl_day_9,ttl_day_10
Dim ttl_day_11,ttl_day_12,ttl_day_13,ttl_day_14,ttl_day_15,ttl_day_16,ttl_day_17,ttl_day_18,ttl_day_19,ttl_day_20
Dim ttl_day_21,ttl_day_22,ttl_day_23,ttl_day_24,ttl_day_25,ttl_day_26,ttl_day_27,ttl_day_28,ttl_day_29,ttl_day_30,ttl_day_31,ttl_tbs_all


cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

tahunx	= trim(REQUEST("tahun"))
bulanx	= trim(REQUEST("bulan"))
TAHUNBULAN = tahunx&bulanx
businessunit= trim(session("fcba"))


		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
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

	txtSQL = "SELECT D.AREAID,A.AREANAME, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE , 'dd')), 1,D.FFBWEIGHT   , 0)) day_1, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 2, D.FFBWEIGHT, 0)) day_2 , " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 3, D.FFBWEIGHT, 0)) day_3 , " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 4, D.FFBWEIGHT, 0)) day_4, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 5, D.FFBWEIGHT, 0)) day_5, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 6, D.FFBWEIGHT, 0)) day_6, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 7, D.FFBWEIGHT, 0)) day_7, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 8, D.FFBWEIGHT, 0)) day_8, " & _
	" SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 9, D.FFBWEIGHT, 0)) day_9, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 10, D.FFBWEIGHT, 0)) day_10, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 11, D.FFBWEIGHT, 0)) day_11, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 12, D.FFBWEIGHT, 0)) day_12, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 13, D.FFBWEIGHT, 0)) day_13, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 14, D.FFBWEIGHT, 0)) day_14, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 15, D.FFBWEIGHT, 0)) day_15, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 16, D.FFBWEIGHT, 0)) day_16, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 17, D.FFBWEIGHT, 0)) day_17, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 18, D.FFBWEIGHT, 0)) day_18, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 19, D.FFBWEIGHT, 0)) day_19, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 20, D.FFBWEIGHT, 0)) day_20, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 21, D.FFBWEIGHT, 0)) day_21, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 22, D.FFBWEIGHT, 0)) day_22, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 23, D.FFBWEIGHT, 0)) day_23, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 24, D.FFBWEIGHT, 0)) day_24, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 25, D.FFBWEIGHT, 0)) day_25, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 26, D.FFBWEIGHT, 0)) day_26, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 27, D.FFBWEIGHT, 0)) day_27, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 28, D.FFBWEIGHT, 0)) day_28, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 29, D.FFBWEIGHT, 0)) day_29, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 30, D.FFBWEIGHT, 0)) day_30, " & _
    " SUM(DECODE(TO_NUMBER(TO_CHAR(H.DOCKETDATE, 'dd')), 31, D.FFBWEIGHT, 0)) day_31 " & _
	" from LHPTIKETHDR H " & _
	" INNER JOIN LHPFFBWBTRN D ON (H.DOCKETID=D.DOCKETID AND H.FCBA=D.FCBA ) " & _
	" INNER JOIN LHPAREAMST A ON (D.AREAID=A.AREAID AND D.FCBA=A.FCBA) " & _
	" WHERE H.fcba = '" & businessunit & "' AND TO_CHAR(H.DOCKETDATE, 'yyyymm') = " & tahunx & "||LPAD(" & bulanx & ", 2, 0)  " & _
	" AND H.PRODUCTID=2   " & _
	" GROUP BY A.EPLASCODE,D.AREAID,A.AREANAME " & _
	" ORDER BY A.EPLASCODE asc"
 	Set rsTBS = Server.CreateObject("ADODB.Recordset")
	rsTBS.CursorType = 3
	rsTBS.CursorLocation = 3
	rsTBS.Open txtSQL, connstr
	if not rsTBS.EOF THEN 
		areaid = rsTBS("AREAID")
		areaname = rsTBS("AREANAME")
	END IF

'Set rsTBS = Server.CreateObject("ADODB.Recordset")
'rsTBS Open txtSQL, connstr, adOpenStatic

'response.Write(txtSQL)
'response.Flush()
%>

<body>
	<div id="wrapper">	
		<div id="header-wrapper">
		<div class="printhide">
        <!--#include file="print_excel.asp"-->
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					
					<th colspan="35" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=COmpanyName%></h4>

								<h5><%=fcbaname%></h5>
							</div>
							
						</div>
						<div id="subheader" class="ratakiri">
								
							<tr>
								<th align="left" colspan="35" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:left;">LHP REKAP</th>
							</tr>
							<tr>	
								<th align="left" colspan="35" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:left;">PERIODE: <%=tahunbulan%></th>
							</tr>
						</div>
					</th>
				</tr>
				
	
	<tr>
		 <th style="color:#fff;" class=lebarket>Keterangan</th>
		 <th style="color:#fff;">UOM</th> 
		 <th style="color:#fff;">Stock Bln Lalu</th> 
		 <%for nday = 1 to 31%>
		<th ><%=nday%></th>
		<%next%>
		 <th style="color:#fff;">Total</th> 
	</tr>
	
	
</thead>
	<tbody>

		
	<tr >
		<td align="center" class=lebarket><strong>TBS</strong></td>
		<td style="color:#fff;"></td>
		<td style="color:#fff;"></td>
		 <%for nday = 1 to 31%>
		<td></td>
		<%next%>
		<td></td>
	</tr>
	<%
		NO = 1
		ttl_day_1 = 0
		ttl_day_2= 0
		ttl_day_3= 0
		ttl_day_4= 0
		ttl_day_5= 0
		ttl_day_6= 0
		ttl_day_7= 0
		ttl_day_8= 0
		ttl_day_9= 0
		ttl_day_10= 0 
		ttl_day_11 = 0
		ttl_day_12 = 0
		ttl_day_13 = 0
		ttl_day_14 = 0 
		ttl_day_15 = 0
		ttl_day_16 = 0
		ttl_day_17 = 0
		ttl_day_18 = 0
		ttl_day_19 = 0
		ttl_day_20 = 0
		ttl_day_21 = 0
		ttl_day_22 = 0
		ttl_day_23 = 0
		ttl_day_24 = 0
		ttl_day_25 = 0
		ttl_day_26 = 0
		ttl_day_27 = 0
		ttl_day_28 = 0
		ttl_day_29 = 0
		ttl_day_30 = 0
		ttl_day_31 = 0
		ttl_tbs_all = 0
		while not rsTBS.eof


		
		subtotaltbs = cdbl(rsTBS("day_1"))+cdbl(rsTBS("day_2"))+cdbl(rsTBS("day_3"))+cdbl(rsTBS("day_4"))+cdbl(rsTBS("day_5"))+cdbl(rsTBS("day_6"))+cdbl(rsTBS("day_7"))+cdbl(rsTBS("day_8"))+cdbl(rsTBS("day_9"))+cdbl(rsTBS("day_10"))+ cdbl(rsTBS("day_11"))+cdbl(rsTBS("day_12"))+cdbl(rsTBS("day_13"))+cdbl(rsTBS("day_14"))+cdbl(rsTBS("day_15"))+cdbl(rsTBS("day_16"))+cdbl(rsTBS("day_17"))+cdbl(rsTBS("day_18"))+cdbl(rsTBS("day_19"))+cdbl(rsTBS("day_20"))+cdbl(rsTBS("day_21"))+cdbl(rsTBS("day_22"))+cdbl(rsTBS("day_23"))+cdbl(rsTBS("day_24"))+cdbl(rsTBS("day_25"))+cdbl(rsTBS("day_26"))+cdbl(rsTBS("day_27"))+cdbl(rsTBS("day_28"))+cdbl(rsTBS("day_29"))+cdbl(rsTBS("day_30"))+cdbl(rsTBS("day_31"))

		
	%>
	<tr class="<%=cClass%>">
		<td align="center" Rowspan="" class="lebarket"><%=rsTBS("areaname")%></td> 
		<td align="center">Kg</td>
		<td align="center"></td>
		<td align="right"><%=FormatNumber(rsTBS("day_1"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_2"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_3"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_4"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_5"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_6"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_7"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_8"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_9"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_10"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_11"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_12"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_13"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_14"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_15"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_16"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_17"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_18"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_19"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_20"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_21"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_22"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_23"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_24"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_25"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_26"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_27"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_28"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_29"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_30"),2)%></td>
		<td align="right"><%=FormatNumber(rsTBS("day_31"),2)%></td>
		<td align="right"><%=FormatNumber(subtotaltbs,2)%></td>
	</tr>
	
	
	<%
		
		' if areaname <> rsTBS("areaname") then
		' end if

	
		ttl_day_1 = cdbl(ttl_day_1) + cdbl(rsTBS("day_1"))
		ttl_day_2 =  cdbl(ttl_day_2) + cdbl(rsTBS("day_2"))
		ttl_day_3 = cdbl(ttl_day_3) + cdbl(rsTBS("day_3"))
		ttl_day_4 = cdbl(ttl_day_4) + cdbl(rsTBS("day_4"))
		ttl_day_5 =  cdbl(ttl_day_5) + cdbl(rsTBS("day_5"))
		ttl_day_6= cdbl(ttl_day_6) + cdbl(rsTBS("day_6"))
		ttl_day_7= cdbl(ttl_day_7) + cdbl(rsTBS("day_7"))
		ttl_day_8= cdbl(ttl_day_8) + cdbl(rsTBS("day_8"))
		ttl_day_9= cdbl(ttl_day_9) + cdbl(rsTBS("day_9"))
		ttl_day_10= cdbl(ttl_day_10) + cdbl(rsTBS("day_10"))
		ttl_day_11 = cdbl(ttl_day_11) + cdbl(rsTBS("day_11"))
		ttl_day_12 = cdbl(ttl_day_12) + cdbl(rsTBS("day_12"))
		ttl_day_13 = cdbl(ttl_day_13) + cdbl(rsTBS("day_13"))
		ttl_day_14 = cdbl(ttl_day_14) + cdbl(rsTBS("day_14"))
		ttl_day_15 = cdbl(ttl_day_15) + cdbl(rsTBS("day_15"))
		ttl_day_16 = cdbl(ttl_day_16) + cdbl(rsTBS("day_16"))
		ttl_day_17 = cdbl(ttl_day_17) + cdbl(rsTBS("day_17"))
		ttl_day_18 = cdbl(ttl_day_18) + cdbl(rsTBS("day_18"))
		ttl_day_19 = cdbl(ttl_day_19) + cdbl(rsTBS("day_19"))
		ttl_day_20 = cdbl(ttl_day_20) + cdbl(rsTBS("day_20"))
		ttl_day_21 = cdbl(ttl_day_21) + cdbl(rsTBS("day_21"))
		ttl_day_22 = cdbl(ttl_day_22) + cdbl(rsTBS("day_22"))
		ttl_day_23 = cdbl(ttl_day_23) + cdbl(rsTBS("day_23"))
		ttl_day_24 = cdbl(ttl_day_24) + cdbl(rsTBS("day_24"))
		ttl_day_25 = cdbl(ttl_day_25) + cdbl(rsTBS("day_25"))
		ttl_day_26 = cdbl(ttl_day_26) + cdbl(rsTBS("day_26"))
		ttl_day_27 = cdbl(ttl_day_27) + cdbl(rsTBS("day_27"))
		ttl_day_28 = cdbl(ttl_day_28) + cdbl(rsTBS("day_28"))
		ttl_day_29 = cdbl(ttl_day_29) + cdbl(rsTBS("day_29"))
		ttl_day_30 = cdbl(ttl_day_30) + cdbl(rsTBS("day_30"))
		ttl_day_31 = cdbl(ttl_day_31) + cdbl(rsTBS("day_31"))
		ttl_tbs_all = cdbl(ttl_tbs_all) + cdbl(subtotaltbs)
		NO = NO + 1	
		rsTBS.MoveNext
		WEND
		rsTBS.Close
		set rsTBS = Nothing
	
	%>

	<tr class="tebal ratakanan" style="background-color:#F4F4A0;">
		<td colspan=3 align=right>Total TBS Terima</td>
		<td align="right"><%=FormatNumber(ttl_day_1,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_2,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_3,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_4,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_5,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_6,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_7,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_8,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_9,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_10,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_11,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_12,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_13,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_14,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_15,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_16,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_17,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_18,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_19,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_20,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_21,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_22,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_23,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_24,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_25,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_26,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_27,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_28,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_29,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_30,2)%></td>
		<td align="right"><%=FormatNumber(ttl_day_31,2)%></td>
		<td align="right"><%=FormatNumber(ttl_tbs_all,2)%></td>
	</tr>

	<!--------------------------- BATAS LHP TBS  --------------------------------------->

	<%

	   dim sqlproduction,rsproduction,arr_prod,arr_prod_eof,xi
	   sqlproduction = "SELECT NVL(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.TBSolah    , 0)),0) tbsolah_day1, " & _
	   " NVL(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.JAMKERJA     , 0)),0) JAMKERJA_day1, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.TBSSISA     , 0)) tbssisa_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.BREAKDOWN     , 0)) BREAKDOWN_day1, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 1,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 1,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 1,P.JAMKERJA , 0)))THROUGPUT_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.HASILCPO     , 0)) HASILCPO_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.RENDEMENCPO     , 0)) RENDEMENCPO_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.CPOTANKISI     , 0)) CPOTANKISI_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.HASILKERNEL     , 0)) HASILKERNEL_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.PKTANKISI     , 0)) PKTANKISI_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day1, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.RENDEMENCKG     , 0)) RENDEMENCKG_day1, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.TBSolah    , 0)) tbsolah_day2, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.JAMKERJA     , 0)) JAMKERJA_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.TBSSISA     , 0)) tbssisa_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.BREAKDOWN     , 0)) BREAKDOWN_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.HASILCPO     , 0)) HASILCPO_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.RENDEMENCPO     , 0)) RENDEMENCPO_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.CPOTANKISI     , 0)) CPOTANKISI_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.HASILKERNEL     , 0)) HASILKERNEL_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.PKTANKISI     , 0)) PKTANKISI_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.RENDEMENCKG     , 0)) RENDEMENCKG_day2, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 2,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 2,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 2,P.JAMKERJA , 0)))THROUGPUT_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.TBSolah    , 0)) tbsolah_day3, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.JAMKERJA     , 0)) JAMKERJA_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.TBSSISA     , 0)) tbssisa_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.BREAKDOWN     , 0)) BREAKDOWN_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.HASILCPO     , 0)) HASILCPO_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.RENDEMENCPO     , 0)) RENDEMENCPO_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.CPOTANKISI     , 0)) CPOTANKISI_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.HASILKERNEL     , 0)) HASILKERNEL_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.PKTANKISI     , 0)) PKTANKISI_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.RENDEMENCKG     , 0)) RENDEMENCKG_day3, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 3,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 3,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 3,P.JAMKERJA , 0)))THROUGPUT_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.TBSolah    , 0)) tbsolah_day4, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.JAMKERJA     , 0)) JAMKERJA_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.TBSSISA     , 0)) tbssisa_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.BREAKDOWN     , 0)) BREAKDOWN_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.HASILCPO     , 0)) HASILCPO_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.RENDEMENCPO     , 0)) RENDEMENCPO_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.CPOTANKISI     , 0)) CPOTANKISI_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.HASILKERNEL     , 0)) HASILKERNEL_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.PKTANKISI     , 0)) PKTANKISI_day4, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.RENDEMENCKG     , 0)) RENDEMENCKG_day4, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 4,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 4,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 4,P.JAMKERJA , 0)))THROUGPUT_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.TBSolah    , 0)) tbsolah_day5, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.JAMKERJA     , 0)) JAMKERJA_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.TBSSISA     , 0)) tbssisa_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.BREAKDOWN     , 0)) BREAKDOWN_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.HASILCPO     , 0)) HASILCPO_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.RENDEMENCPO     , 0)) RENDEMENCPO_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.CPOTANKISI     , 0)) CPOTANKISI_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.HASILKERNEL     , 0)) HASILKERNEL_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.PKTANKISI     , 0)) PKTANKISI_day5, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.RENDEMENCKG     , 0)) RENDEMENCKG_day5, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 5,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 5,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 5,P.JAMKERJA , 0)))THROUGPUT_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.TBSolah    , 0)) tbsolah_day6, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.JAMKERJA     , 0)) JAMKERJA_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.TBSSISA     , 0)) tbssisa_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.BREAKDOWN     , 0)) BREAKDOWN_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.HASILCPO     , 0)) HASILCPO_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.RENDEMENCPO     , 0)) RENDEMENCPO_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.CPOTANKISI     , 0)) CPOTANKISI_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.HASILKERNEL     , 0)) HASILKERNEL_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.PKTANKISI     , 0)) PKTANKISI_day6, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.RENDEMENCKG     , 0)) RENDEMENCKG_day6, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 6,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 6,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 6,P.JAMKERJA , 0)))THROUGPUT_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.TBSolah    , 0)) tbsolah_day7, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.JAMKERJA     , 0)) JAMKERJA_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.TBSSISA     , 0)) tbssisa_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.BREAKDOWN     , 0)) BREAKDOWN_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.HASILCPO     , 0)) HASILCPO_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.RENDEMENCPO     , 0)) RENDEMENCPO_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.CPOTANKISI     , 0)) CPOTANKISI_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.HASILKERNEL     , 0)) HASILKERNEL_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.PKTANKISI     , 0)) PKTANKISI_day7, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.RENDEMENCKG     , 0)) RENDEMENCKG_day7, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 7,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 7,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 7,P.JAMKERJA , 0)))THROUGPUT_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.TBSolah    , 0)) tbsolah_day8, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.JAMKERJA     , 0)) JAMKERJA_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.TBSSISA     , 0)) tbssisa_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.BREAKDOWN     , 0)) BREAKDOWN_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.HASILCPO     , 0)) HASILCPO_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.RENDEMENCPO     , 0)) RENDEMENCPO_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.CPOTANKISI     , 0)) CPOTANKISI_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.HASILKERNEL     , 0)) HASILKERNEL_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.PKTANKISI     , 0)) PKTANKISI_day8, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.RENDEMENCKG     , 0)) RENDEMENCKG_day8, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 8,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 8,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 8,P.JAMKERJA , 0)))THROUGPUT_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.TBSolah    , 0)) tbsolah_day9, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.JAMKERJA     , 0)) JAMKERJA_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.TBSSISA     , 0)) tbssisa_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.BREAKDOWN     , 0)) BREAKDOWN_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.HASILCPO     , 0)) HASILCPO_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.RENDEMENCPO     , 0)) RENDEMENCPO_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.CPOTANKISI     , 0)) CPOTANKISI_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.HASILKERNEL     , 0)) HASILKERNEL_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.PKTANKISI     , 0)) PKTANKISI_day9, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.RENDEMENCKG     , 0)) RENDEMENCKG_day9, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 9,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 9,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 9,P.JAMKERJA , 0)))THROUGPUT_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.TBSolah    , 0)) tbsolah_day10, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.JAMKERJA     , 0)) JAMKERJA_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.TBSSISA     , 0)) tbssisa_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.BREAKDOWN     , 0)) BREAKDOWN_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.HASILCPO     , 0)) HASILCPO_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.RENDEMENCPO     , 0)) RENDEMENCPO_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.CPOTANKISI     , 0)) CPOTANKISI_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.HASILKERNEL     , 0)) HASILKERNEL_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.PKTANKISI     , 0)) PKTANKISI_day10, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.RENDEMENCKG     , 0)) RENDEMENCKG_day10, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 10,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 10,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 10,P.JAMKERJA , 0)))THROUGPUT_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.TBSolah    , 0)) tbsolah_day11, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.JAMKERJA     , 0)) JAMKERJA_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.TBSSISA     , 0)) tbssisa_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.BREAKDOWN     , 0)) BREAKDOWN_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.HASILCPO     , 0)) HASILCPO_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.RENDEMENCPO     , 0)) RENDEMENCPO_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.CPOTANKISI     , 0)) CPOTANKISI_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.HASILKERNEL     , 0)) HASILKERNEL_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.PKTANKISI     , 0)) PKTANKISI_day11, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.RENDEMENCKG     , 0)) RENDEMENCKG_day11, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 11,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 11,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 11,P.JAMKERJA , 0)))THROUGPUT_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.TBSolah    , 0)) tbsolah_day12, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.JAMKERJA     , 0)) JAMKERJA_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.TBSSISA     , 0)) tbssisa_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.BREAKDOWN     , 0)) BREAKDOWN_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.HASILCPO     , 0)) HASILCPO_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.RENDEMENCPO     , 0)) RENDEMENCPO_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.CPOTANKISI     , 0)) CPOTANKISI_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.HASILKERNEL     , 0)) HASILKERNEL_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.PKTANKISI     , 0)) PKTANKISI_day12, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.RENDEMENCKG     , 0)) RENDEMENCKG_day12, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 12,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 12,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 12,P.JAMKERJA , 0)))THROUGPUT_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.TBSolah    , 0)) tbsolah_day13, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.JAMKERJA     , 0)) JAMKERJA_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.TBSSISA     , 0)) tbssisa_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.BREAKDOWN     , 0)) BREAKDOWN_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.HASILCPO     , 0)) HASILCPO_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.RENDEMENCPO     , 0)) RENDEMENCPO_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.CPOTANKISI     , 0)) CPOTANKISI_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.HASILKERNEL     , 0)) HASILKERNEL_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.PKTANKISI     , 0)) PKTANKISI_day13, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.RENDEMENCKG     , 0)) RENDEMENCKG_day13, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 13,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 13,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 13,P.JAMKERJA , 0)))THROUGPUT_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.TBSolah    , 0)) tbsolah_day14, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.JAMKERJA     , 0)) JAMKERJA_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.TBSSISA     , 0)) tbssisa_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.BREAKDOWN     , 0)) BREAKDOWN_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.HASILCPO     , 0)) HASILCPO_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.RENDEMENCPO     , 0)) RENDEMENCPO_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.CPOTANKISI     , 0)) CPOTANKISI_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.HASILKERNEL     , 0)) HASILKERNEL_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.PKTANKISI     , 0)) PKTANKISI_day14, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.RENDEMENCKG     , 0)) RENDEMENCKG_day14, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 14,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 14,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 14,P.JAMKERJA , 0)))THROUGPUT_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.TBSolah    , 0)) tbsolah_day15, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.JAMKERJA     , 0)) JAMKERJA_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.TBSSISA     , 0)) tbssisa_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.BREAKDOWN     , 0)) BREAKDOWN_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.HASILCPO     , 0)) HASILCPO_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.RENDEMENCPO     , 0)) RENDEMENCPO_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.CPOTANKISI     , 0)) CPOTANKISI_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.HASILKERNEL     , 0)) HASILKERNEL_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.PKTANKISI     , 0)) PKTANKISI_day15, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.RENDEMENCKG     , 0)) RENDEMENCKG_day15, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 15,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 15,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 15,P.JAMKERJA , 0)))THROUGPUT_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.TBSolah    , 0)) tbsolah_day16, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.JAMKERJA     , 0)) JAMKERJA_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.TBSSISA     , 0)) tbssisa_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.BREAKDOWN     , 0)) BREAKDOWN_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.HASILCPO     , 0)) HASILCPO_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.RENDEMENCPO     , 0)) RENDEMENCPO_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.CPOTANKISI     , 0)) CPOTANKISI_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.HASILKERNEL     , 0)) HASILKERNEL_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.PKTANKISI     , 0)) PKTANKISI_day16, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.RENDEMENCKG     , 0)) RENDEMENCKG_day16, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 16,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 16,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 16,P.JAMKERJA , 0)))THROUGPUT_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.TBSolah    , 0)) tbsolah_day17, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.JAMKERJA     , 0)) JAMKERJA_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.TBSSISA     , 0)) tbssisa_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.BREAKDOWN     , 0)) BREAKDOWN_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.HASILCPO     , 0)) HASILCPO_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.RENDEMENCPO     , 0)) RENDEMENCPO_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.CPOTANKISI     , 0)) CPOTANKISI_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.HASILKERNEL     , 0)) HASILKERNEL_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.PKTANKISI     , 0)) PKTANKISI_day17, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.RENDEMENCKG     , 0)) RENDEMENCKG_day17, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 17,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 17,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 17,P.JAMKERJA , 0)))THROUGPUT_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.TBSolah    , 0)) tbsolah_day18, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.JAMKERJA     , 0)) JAMKERJA_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.TBSSISA     , 0)) tbssisa_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.BREAKDOWN     , 0)) BREAKDOWN_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.HASILCPO     , 0)) HASILCPO_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.RENDEMENCPO     , 0)) RENDEMENCPO_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.CPOTANKISI     , 0)) CPOTANKISI_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.HASILKERNEL     , 0)) HASILKERNEL_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.PKTANKISI     , 0)) PKTANKISI_day18, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.RENDEMENCKG     , 0)) RENDEMENCKG_day18, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 18,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 18,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 18,P.JAMKERJA , 0)))THROUGPUT_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.TBSolah    , 0)) tbsolah_day19, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.JAMKERJA     , 0)) JAMKERJA_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.TBSSISA     , 0)) tbssisa_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.BREAKDOWN     , 0)) BREAKDOWN_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.HASILCPO     , 0)) HASILCPO_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.RENDEMENCPO     , 0)) RENDEMENCPO_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.CPOTANKISI     , 0)) CPOTANKISI_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.HASILKERNEL     , 0)) HASILKERNEL_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.PKTANKISI     , 0)) PKTANKISI_day19, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.RENDEMENCKG     , 0)) RENDEMENCKG_day19, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 19,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 19,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 19,P.JAMKERJA , 0)))THROUGPUT_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.TBSolah    , 0)) tbsolah_day20, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.JAMKERJA     , 0)) JAMKERJA_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.TBSSISA     , 0)) tbssisa_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.BREAKDOWN     , 0)) BREAKDOWN_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.HASILCPO     , 0)) HASILCPO_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.RENDEMENCPO     , 0)) RENDEMENCPO_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.CPOTANKISI     , 0)) CPOTANKISI_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.HASILKERNEL     , 0)) HASILKERNEL_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.PKTANKISI     , 0)) PKTANKISI_day20, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.RENDEMENCKG     , 0)) RENDEMENCKG_day20, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 20,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 20,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 20,P.JAMKERJA , 0)))THROUGPUT_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.TBSolah    , 0)) tbsolah_day21, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.JAMKERJA     , 0)) JAMKERJA_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.TBSSISA     , 0)) tbssisa_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.BREAKDOWN     , 0)) BREAKDOWN_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.HASILCPO     , 0)) HASILCPO_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.RENDEMENCPO     , 0)) RENDEMENCPO_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.CPOTANKISI     , 0)) CPOTANKISI_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.HASILKERNEL     , 0)) HASILKERNEL_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.PKTANKISI     , 0)) PKTANKISI_day21, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.RENDEMENCKG     , 0)) RENDEMENCKG_day21, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 21,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 21,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 21,P.JAMKERJA , 0)))THROUGPUT_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.TBSolah    , 0)) tbsolah_day22, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.JAMKERJA     , 0)) JAMKERJA_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.TBSSISA     , 0)) tbssisa_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.BREAKDOWN     , 0)) BREAKDOWN_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.HASILCPO     , 0)) HASILCPO_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.RENDEMENCPO     , 0)) RENDEMENCPO_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.CPOTANKISI     , 0)) CPOTANKISI_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.HASILKERNEL     , 0)) HASILKERNEL_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.PKTANKISI     , 0)) PKTANKISI_day22, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.RENDEMENCKG     , 0)) RENDEMENCKG_day22, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 22,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 22,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 22,P.JAMKERJA , 0)))THROUGPUT_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.TBSolah    , 0)) tbsolah_day23, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.JAMKERJA     , 0)) JAMKERJA_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.TBSSISA     , 0)) tbssisa_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.BREAKDOWN     , 0)) BREAKDOWN_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.HASILCPO     , 0)) HASILCPO_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.RENDEMENCPO     , 0)) RENDEMENCPO_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.CPOTANKISI     , 0)) CPOTANKISI_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.HASILKERNEL     , 0)) HASILKERNEL_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.PKTANKISI     , 0)) PKTANKISI_day23, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.RENDEMENCKG     , 0)) RENDEMENCKG_day23, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 23,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 23,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 23,P.JAMKERJA , 0)))THROUGPUT_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.TBSolah    , 0)) tbsolah_day24, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.JAMKERJA     , 0)) JAMKERJA_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.TBSSISA     , 0)) tbssisa_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.BREAKDOWN     , 0)) BREAKDOWN_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.HASILCPO     , 0)) HASILCPO_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.RENDEMENCPO     , 0)) RENDEMENCPO_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.CPOTANKISI     , 0)) CPOTANKISI_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.HASILKERNEL     , 0)) HASILKERNEL_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.PKTANKISI     , 0)) PKTANKISI_day24, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.RENDEMENCKG     , 0)) RENDEMENCKG_day24, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 24,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 24,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 24,P.JAMKERJA , 0)))THROUGPUT_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.TBSolah    , 0)) tbsolah_day25, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.JAMKERJA     , 0)) JAMKERJA_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.TBSSISA     , 0)) tbssisa_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.BREAKDOWN     , 0)) BREAKDOWN_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.HASILCPO     , 0)) HASILCPO_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.RENDEMENCPO     , 0)) RENDEMENCPO_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.CPOTANKISI     , 0)) CPOTANKISI_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.HASILKERNEL     , 0)) HASILKERNEL_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.PKTANKISI     , 0)) PKTANKISI_day25, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.RENDEMENCKG     , 0)) RENDEMENCKG_day25, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 25,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 25,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 25,P.JAMKERJA , 0)))THROUGPUT_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.TBSolah    , 0)) tbsolah_day26, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.JAMKERJA     , 0)) JAMKERJA_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.TBSSISA     , 0)) tbssisa_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.BREAKDOWN     , 0)) BREAKDOWN_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.HASILCPO     , 0)) HASILCPO_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.RENDEMENCPO     , 0)) RENDEMENCPO_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.CPOTANKISI     , 0)) CPOTANKISI_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.HASILKERNEL     , 0)) HASILKERNEL_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.PKTANKISI     , 0)) PKTANKISI_day26, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.RENDEMENCKG     , 0)) RENDEMENCKG_day26, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 26,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 26,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 26,P.JAMKERJA , 0)))THROUGPUT_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.TBSolah    , 0)) tbsolah_day27, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.JAMKERJA     , 0)) JAMKERJA_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.TBSSISA     , 0)) tbssisa_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.BREAKDOWN     , 0)) BREAKDOWN_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.HASILCPO     , 0)) HASILCPO_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.RENDEMENCPO     , 0)) RENDEMENCPO_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.CPOTANKISI     , 0)) CPOTANKISI_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.HASILKERNEL     , 0)) HASILKERNEL_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.PKTANKISI     , 0)) PKTANKISI_day27, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.RENDEMENCKG     , 0)) RENDEMENCKG_day27, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 27,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 27,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 27,P.JAMKERJA , 0)))THROUGPUT_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.TBSolah    , 0)) tbsolah_day28, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.JAMKERJA     , 0)) JAMKERJA_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.TBSSISA     , 0)) tbssisa_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.BREAKDOWN     , 0)) BREAKDOWN_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.HASILCPO     , 0)) HASILCPO_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.RENDEMENCPO     , 0)) RENDEMENCPO_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.CPOTANKISI     , 0)) CPOTANKISI_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.HASILKERNEL     , 0)) HASILKERNEL_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.PKTANKISI     , 0)) PKTANKISI_day28, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.RENDEMENCKG     , 0)) RENDEMENCKG_day28, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 28,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 28,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 28,P.JAMKERJA , 0)))THROUGPUT_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.TBSolah    , 0)) tbsolah_day29, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.JAMKERJA     , 0)) JAMKERJA_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.TBSSISA     , 0)) tbssisa_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.BREAKDOWN     , 0)) BREAKDOWN_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.HASILCPO     , 0)) HASILCPO_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.RENDEMENCPO     , 0)) RENDEMENCPO_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.CPOTANKISI     , 0)) CPOTANKISI_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.HASILKERNEL     , 0)) HASILKERNEL_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.PKTANKISI     , 0)) PKTANKISI_day29, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.RENDEMENCKG     , 0)) RENDEMENCKG_day29, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 29,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 29,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 29,P.JAMKERJA , 0)))THROUGPUT_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.TBSolah    , 0)) tbsolah_day30, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.JAMKERJA     , 0)) JAMKERJA_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.TBSSISA     , 0)) tbssisa_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.BREAKDOWN     , 0)) BREAKDOWN_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.HASILCPO     , 0)) HASILCPO_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.RENDEMENCPO     , 0)) RENDEMENCPO_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.CPOTANKISI     , 0)) CPOTANKISI_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.HASILKERNEL     , 0)) HASILKERNEL_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.PKTANKISI     , 0)) PKTANKISI_day30, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.RENDEMENCKG     , 0)) RENDEMENCKG_day30, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 30,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 30,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 30,P.JAMKERJA , 0)))THROUGPUT_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.TBSolah    , 0)) tbsolah_day31, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.JAMKERJA     , 0)) JAMKERJA_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.TBSSISA     , 0)) tbssisa_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.BREAKDOWN     , 0)) BREAKDOWN_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.HASILCPO     , 0)) HASILCPO_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.RENDEMENCPO     , 0)) RENDEMENCPO_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.RENDEMENKERNEL     , 0)) RENDEMENKERNEL_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.CPOTANKISI     , 0)) CPOTANKISI_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.HASILKERNEL     , 0)) HASILKERNEL_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.PKTANKISI     , 0)) PKTANKISI_day31, " & _
	     " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.CANGKANGPRODUKSI     , 0)) CANGKANGPRODUKSI_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.CANGKANGKIRIM     , 0)) CANGKANGKIRIM_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.RENDEMENCKG     , 0)) RENDEMENCKG_day31, " & _
	   " DECODE(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 31,P.TBSolah , 0)),0,0,SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 31,P.TBSolah , 0))/SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE , 'dd')), 31,P.JAMKERJA , 0)))THROUGPUT_day31 " & _
	   " FROM PRODUCTIONTRN P " & _
	   " WHERE FCBA='" & businessunit & "' AND TO_CHAR(P.PRODDATE, 'yyyymm') = " & tahunx & "||LPAD(" & bulanx & ", 2, 0) "
 		Set rsproduction = Server.CreateObject("ADODB.Recordset")
		rsproduction.Open sqlproduction, connstr
		'Response.Write sqlproduction
	
		dim no, tbssisakemarin,subtotalrestan,hariolah,tbsolahloop,tbsx
		dim tbsolahhari1,tbsolahhari2,tbsolahhari3,tbsolahhari4,tbsolahhari5
		dim tbsolahhari6,tbsolahhari7,tbsolahhari8,tbsolahhari9,tbsolahhari10
		dim tbsolahhari11,tbsolahhari12,tbsolahhari13,tbsolahhari14,tbsolahhari15
		dim tbsolahhari16,tbsolahhari17,tbsolahhari18,tbsolahhari19,tbsolahhari20
		dim tbsolahhari21,tbsolahhari22,tbsolahhari23,tbsolahhari24,tbsolahhari25
		dim  bsolahhari26,tbsolahhari27,tbsolahhari28,tbsolahhari29,tbsolahhari30,tbsolahhari31
		dim  hari1,hari2,hari3,hari4,hari5,hari6,hari7,hari8,hari9,hari10,hari11,hari12,hari13
		dim hari14,hari15,hari16,hari17,hari18,hari19,hari20,hari21,hari22,hari23,hari24,hari25
		dim hari26,hari27,hari28,hari29,hari30,hari31
		dim m_util1,m_util2,m_util3,m_util4,m_util5,m_util6,m_util7,m_util8,m_util9,m_util10,m_util11,m_util12,m_util13
		dim m_util14,m_util15,m_util16,m_util17,m_util18,m_util19,m_util20,m_util21,m_util22,m_util23,m_util24,m_util25
		dim m_util26,m_util27,m_util28,m_util29,m_util30,m_util31

		dim m_jjgkosong1,m_jjgkosong2,m_jjgkosong3,m_jjgkosong4,m_jjgkosong5,m_jjgkosong6,m_jjgkosong7,m_jjgkosong8,m_jjgkosong9,m_jjgkosong10,m_jjgkosong11,m_jjgkosong12,m_jjgkosong13
		dim m_jjgkosong14,m_jjgkosong15,m_jjgkosong16,m_jjgkosong17,m_jjgkosong18,m_jjgkosong19,m_jjgkosong20,m_jjgkosong21,m_jjgkosong22,m_jjgkosong23,m_jjgkosong24,m_jjgkosong25
		dim m_jjgkosong26,m_jjgkosong27,m_jjgkosong28,m_jjgkosong29,m_jjgkosong30,m_jjgkosong31

		no = 1
		while not rsproduction.eof
			tbsolahhari1 = rsproduction("tbsolah_day1")
			tbsolahhari2 = rsproduction("tbsolah_day2")
			tbsolahhari3 = rsproduction("tbsolah_day3")
			tbsolahhari4 = rsproduction("tbsolah_day4")
			tbsolahhari5 = rsproduction("tbsolah_day5")
			tbsolahhari6 = rsproduction("tbsolah_day6")
			tbsolahhari7 = rsproduction("tbsolah_day7")
			tbsolahhari8 = rsproduction("tbsolah_day8")
			tbsolahhari9 = rsproduction("tbsolah_day9")
			tbsolahhari10 = rsproduction("tbsolah_day10")
			tbsolahhari11 = rsproduction("tbsolah_day11")
			tbsolahhari12 = rsproduction("tbsolah_day12")
			tbsolahhari13 = rsproduction("tbsolah_day13")
			tbsolahhari14 = rsproduction("tbsolah_day14")
			tbsolahhari15 = rsproduction("tbsolah_day15")
			tbsolahhari16 = rsproduction("tbsolah_day16")
			tbsolahhari17 = rsproduction("tbsolah_day17")
			tbsolahhari18 = rsproduction("tbsolah_day18")
			tbsolahhari19 = rsproduction("tbsolah_day19")
			tbsolahhari20 = rsproduction("tbsolah_day20")
			tbsolahhari21 = rsproduction("tbsolah_day21")
			tbsolahhari22 = rsproduction("tbsolah_day22")
			tbsolahhari23 = rsproduction("tbsolah_day23")
			tbsolahhari24 = rsproduction("tbsolah_day24")
			tbsolahhari25 = rsproduction("tbsolah_day25")
			tbsolahhari26 = rsproduction("tbsolah_day26")
			tbsolahhari27 = rsproduction("tbsolah_day27")
			tbsolahhari28 = rsproduction("tbsolah_day28")
			tbsolahhari29 = rsproduction("tbsolah_day29")
			tbsolahhari30 = rsproduction("tbsolah_day30")
			tbsolahhari31 = rsproduction("tbsolah_day31")
			if cdbl(tbsolahhari1)>0 then 
				hari1 = 1 
				m_util1 = cdbl(tbsolahhari1)/1200000
				m_jjgkosong1 = cdbl(tbsolahhari1)*21.5/100
			else 
				hari1 = 0
			end if 
			if cdbl(tbsolahhari2)>0 then 
				hari2= 1
				m_util2 = cdbl(tbsolahhari2)/1200000
				m_jjgkosong2 = cdbl(tbsolahhari2)*21.5/100
			else 
				hari2=0
			end if 
			if cdbl(tbsolahhari3)>0 then 
				hari3= 1
				m_util3 = cdbl(tbsolahhari3)/1200000
				m_jjgkosong3 = cdbl(tbsolahhari3)*21.5/100
			else 
				hari3=0
			end if 
			if cdbl(tbsolahhari4)>0 then 
				hari4= 1
				m_util4 = cdbl(tbsolahhari4)/1200000
				m_jjgkosong4 = cdbl(tbsolahhari4)*21.5/100
			else 
				hari4=0
			end if 
			if cdbl(tbsolahhari5)>0 then 
				hari5= 1
				m_util5 = cdbl(tbsolahhari5)/1200000
				m_jjgkosong5 = cdbl(tbsolahhari5)*21.5/100
			else 
				hari5=0
			end if 
			if cdbl(tbsolahhari6)>0 then 
				hari6= 1
				m_util6 = cdbl(tbsolahhari6)/1200000
				m_jjgkosong6 = cdbl(tbsolahhari6)*21.5/100
			else 
				hari6=0
			end if 
			if cdbl(tbsolahhari7)>0 then 
				hari7= 1
				m_util7 = cdbl(tbsolahhari7)/1200000
				m_jjgkosong7 = cdbl(tbsolahhari7)*21.5/100
			else 
				hari7=0
			end if 
			if cdbl(tbsolahhari8)>0 then 
				hari8= 1
				m_util8 = cdbl(tbsolahhari8)/1200000
				m_jjgkosong8 = cdbl(tbsolahhari8)*21.5/100
			else 
				hari8=0
			end if 
			if cdbl(tbsolahhari9)>0 then 
				hari9= 1
				m_util9 = cdbl(tbsolahhari9)/1200000
				m_jjgkosong9 = cdbl(tbsolahhari9)*21.5/100
			else 
				hari9=0
			end if 
			if cdbl(tbsolahhari10)>0 then 
				hari10= 1
				m_util10 = cdbl(tbsolahhari10)/1200000
				m_jjgkosong10 = cdbl(tbsolahhari10)*21.5/100
			else 
				hari10=0
			end if 
			if cdbl(tbsolahhari11)>0 then 
				hari11= 1
				m_util11 = cdbl(tbsolahhari11)/1200000
				m_jjgkosong11 = cdbl(tbsolahhari11)*21.5/100
			else 
				hari11=0
			end if 
			if cdbl(tbsolahhari12)>0 then 
				hari12= 1
				m_util12 = cdbl(tbsolahhari12)/1200000
				m_jjgkosong12 = cdbl(tbsolahhari12)*21.5/100
			else 
				hari12=0
			end if 
			if cdbl(tbsolahhari13)>0 then 
				hari13= 1
				m_util13 = cdbl(tbsolahhari13)/1200000
				m_jjgkosong13 = cdbl(tbsolahhari13)*21.5/100
			else 
				hari13=0
			end if 
			if cdbl(tbsolahhari14)>0 then 
				hari14= 1
				m_util14 = cdbl(tbsolahhari14)/1200000
				m_jjgkosong14 = cdbl(tbsolahhari14)*21.5/100
			else 
				hari14=0
			end if 
			if cdbl(tbsolahhari15)>0 then 
				hari15= 1
				m_util15 = cdbl(tbsolahhari15)/1200000
				m_jjgkosong15 = cdbl(tbsolahhari15)*21.5/100
			else 
				hari15=0
			end if 
			if cdbl(tbsolahhari16)>0 then 
				hari16= 1
				m_util16 = cdbl(tbsolahhari16)/1200000
				m_jjgkosong16 = cdbl(tbsolahhari16)*21.5/100
			else 
				hari16=0
			end if 
			if cdbl(tbsolahhari17)>0 then 
				hari17= 1
				m_util17 = cdbl(tbsolahhari17)/1200000
				m_jjgkosong17 = cdbl(tbsolahhari17)*21.5/100
			else 
				hari17=0
			end if 
			if cdbl(tbsolahhari18)>0 then 
				hari18= 1
				m_util18 = cdbl(tbsolahhari18)/1200000
				m_jjgkosong18 = cdbl(tbsolahhari18)*21.5/100
			else 
				hari18=0
			end if 
			if cdbl(tbsolahhari19)>0 then 
				hari19= 1
				m_util19 = cdbl(tbsolahhari19)/1200000
				m_jjgkosong19 = cdbl(tbsolahhari19)*21.5/100
			else 
				hari19=0
			end if 
			if cdbl(tbsolahhari20)>0 then 
				hari20= 1
				m_util20 = cdbl(tbsolahhari20)/1200000
				m_jjgkosong20 = cdbl(tbsolahhari20)*21.5/100
			else 
				hari20=0
			end if 
			if cdbl(tbsolahhari21)>0 then 
				hari21= 1
				m_util21 = cdbl(tbsolahhari21)/1200000
				m_jjgkosong21 = cdbl(tbsolahhari21)*21.5/100
			else 
				hari21=0
			end if 
			if cdbl(tbsolahhari22)>0 then 
				hari22= 1
				m_util22 = cdbl(tbsolahhari22)/1200000
				m_jjgkosong22 = cdbl(tbsolahhari22)*21.5/100
			else 
				hari22=0
			end if 
			if cdbl(tbsolahhari23)>0 then 
				hari23= 1
				m_util23 = cdbl(tbsolahhari23)/1200000
				m_jjgkosong23 = cdbl(tbsolahhari23)*21.5/100
			else 
				hari23=0
			end if 
			if cdbl(tbsolahhari24)>0 then 
				hari24= 1
				m_util24 = cdbl(tbsolahhari24)/1200000
				m_jjgkosong24 = cdbl(tbsolahhari24)*21.5/100
			else 
				hari24=0
			end if 
			if cdbl(tbsolahhari25)>0 then 
				hari25= 1
				m_util25 = cdbl(tbsolahhari25)/1200000
				m_jjgkosong25 = cdbl(tbsolahhari25)*21.5/100
			else 
				hari25=0
			end if 
			if cdbl(tbsolahhari26)>0 then 
				hari26= 1
				m_util26 = cdbl(tbsolahhari26)/1200000
				m_jjgkosong26 = cdbl(tbsolahhari26)*21.5/100
			else 
				hari26=0
			end if 
			if cdbl(tbsolahhari27)>0 then 
				hari27= 1
				m_util27 = cdbl(tbsolahhari27)/1200000
				m_jjgkosong27 = cdbl(tbsolahhari27)*21.5/100
			else 
				hari27=0
			end if 
			if cdbl(tbsolahhari28)>0 then 
				hari28= 1
				m_util28 = cdbl(tbsolahhari28)/1200000
				m_jjgkosong28 = cdbl(tbsolahhari28)*21.5/100
			else 
				hari28=0
			end if 
			if cdbl(tbsolahhari29)>0 then 
				hari29= 1
				m_util29 = cdbl(tbsolahhari29)/1200000
				m_jjgkosong29 = cdbl(tbsolahhari29)*21.5/100
			else 
				hari29=0
			end if 
			if cdbl(tbsolahhari30)>0 then 
				hari30= 1
				m_util30 = cdbl(tbsolahhari30)/1200000
				m_jjgkosong30 = cdbl(tbsolahhari30)*21.5/100
			else 
				hari30=0
			end if 
			if cdbl(tbsolahhari31)>0 then 
				hari31= 1
				m_util31 = cdbl(tbsolahhari31)/1200000
				m_jjgkosong31 = cdbl(tbsolahhari31)*21.5/100
			else 
				hari31=0
			end if 
			
			
			' untuk perhitungan subtotal restan di ambil dari kolom 1,3,5 dan seterusnya sesuai urutan kolomnya
			'subtotalrestan = cdbl(arr_prod(1, xi))+cdbl(arr_prod(3, xi))  +cdbl(arr_prod(5, xi))+cdbl(arr_prod(7, xi))+cdbl(arr_prod(9, xi))+cdbl(arr_prod(11, xi)) +cdbl(arr_prod(13, xi))+cdbl(arr_prod(15, xi))+cdbl(arr_prod(17, xi))+cdbl(arr_prod(19, xi)) 
			
			subtotalrestan= 0
			'  +cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi)) " & _ 
			'  +cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi)) " & _ 
			'  +cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi)) " & _  
			' " +cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))" & _ 
			' " +cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))+cdbl(arr_prod(19, xi))
		%>				
	<tr class=ratakanan>
		<td colspan= class=ratakanan>TBS OLah</td>
		<td colspan= class=ratatengah>KG</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbsolah_day31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>

	<tr class=ratakanan>
		<td colspan= class=ratakanan lebarket>TBS Restan</td>
		<td colspan= class=ratatengah>KG</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("tbssisa_day31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>

	<tr class=ratakanan>
		<td colspan= class=ratakanan>Jam Olah Efektif</td>
		<td colspan= class=ratatengah>JAM</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("JAMKERJA_day31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>

	<tr class=ratakanan>
		<td colspan= class=ratakanan>Total Breakdown</td>
		<td colspan= class=ratatengah>JAM</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("BREAKDOWN_day31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>

	<tr class=ratakanan>
		<td colspan= class=ratakanan>Throughput</td>
		<td colspan= class=ratatengah>TON/JAM</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("THROUGPUT_day31"),2)%></td>
		
		<td colspan= class="ratakanan">0</td>
	</tr>

	<tr class=ratakanan>
		<td colspan= class=ratakanan>Hari Olah</td>
		<td colspan= class=ratatengah>Hari</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari1)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari2)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari3)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari4)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari5)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari6)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari7)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari8)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari9)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari10)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari11)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari12)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari13)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari14)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari15)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari16)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari17)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari18)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari19)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari20)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari21)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari22)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari23)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari24)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari25)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari26)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari27)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari28)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari29)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari30)%></td>
		<td colspan= class="ratakanan"><%=FbpNum2(hari31)%></td>
		
		<td colspan= class="ratakanan">0</td>
	</tr>
	<tr class=ratakanan>
		<td colspan= class=ratakanan>Mill  Utilitation</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util1,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util2,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util3,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util4,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util5,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util6,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util7,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util8,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util9,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util10,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util11,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util12,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util13,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util14,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util15,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util16,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util17,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util18,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util19,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util20,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util21,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util22,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util23,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util24,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util25,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util26,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util27,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util28,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util29,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util30,2)%></td>
		<td colspan= class="ratakanan"><%= formatnumber(m_util31,2)%></td>
		
		<td colspan= class="ratakanan">0</td>
	</tr>
	
	<tr>
		<td colspan=3 align=left><strong>GRADING TBS INTI</strong></td>
	</tr>

		<%
		dim sqlgradinginti,rsgradinginti
		sqlgradinginti = "SELECT KET, " & _
    " nvl(TGL01,0) daygradinginti_1, " & _
    " nvl(TGL02,0) daygradinginti_2 , " & _
    " nvl(TGL03,0) daygradinginti_3 , " & _
    " nvl(TGL04,0) daygradinginti_4, " & _
    " nvl(TGL05,0) daygradinginti_5, " & _
    " nvl(TGL06,0) daygradinginti_6, " & _
    " nvl(TGL07,0) daygradinginti_7, " & _
    " nvl(TGL08,0) daygradinginti_8, " & _
	" nvl(TGL09,0) daygradinginti_9, " & _
    " nvl(TGL10,0) daygradinginti_10, " & _
    " nvl(TGL11,0) daygradinginti_11, " & _
    " nvl(TGL12,0) daygradinginti_12, " & _
    " nvl(TGL13,0) daygradinginti_13, " & _
    " nvl(TGL14,0) daygradinginti_14, " & _
    " nvl(TGL15,0) daygradinginti_15, " & _
    " nvl(TGL16,0) daygradinginti_16, " & _
    " nvl(TGL17,0) daygradinginti_17, " & _
    " nvl(TGL18,0) daygradinginti_18, " & _
    " nvl(TGL19,0) daygradinginti_19, " & _
    " nvl(TGL20,0) daygradinginti_20, " & _
    " nvl(TGL21,0) daygradinginti_21, " & _
    " nvl(TGL22,0) daygradinginti_22, " & _
    " nvl(TGL23,0) daygradinginti_23, " & _
    " nvl(TGL24,0) daygradinginti_24, " & _
    " nvl(TGL25,0) daygradinginti_25, " & _
    " nvl(TGL26,0) daygradinginti_26, " & _
    " nvl(TGL27,0) daygradinginti_27, " & _
    " nvl(TGL28,0) daygradinginti_28, " & _
    " nvl(TGL29,0) daygradinginti_29, " & _
    " nvl(TGL30,0) daygradinginti_30, " & _
    " nvl(TGL31,0) daygradinginti_31 " & _
	" from GRADINGTBSINTI " & _
	" WHERE fcba = '" & businessunit & "' AND TAHUNBULAN = '" & TAHUNBULAN & "' "
	'Response.Write sqlgradinginti
		Set rsgradinginti = Server.CreateObject("ADODB.Recordset")
		rsgradinginti.Open sqlgradinginti, connstr

		while not rsgradinginti.eof

	%>

	<tr class=ratakanan>
		<td colspan= class=ratakanan><%=rsgradinginti("ket")%></td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradinginti("daygradinginti_31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>
	<%
		rsgradinginti.MoveNext
		WEND
		rsgradinginti.Close
		set rsgradinginti = Nothing
	%>
	<tr>
		<td colspan=3 align=left><strong>GRADING TBS LUAR</strong></td>
	</tr>

	<%
		dim sqlgradingluar,rsgradingluar
		sqlgradingluar = "SELECT KET, " & _
    " nvl(TGL01,0) daygradingluar_1, " & _
    " nvl(TGL02,0) daygradingluar_2 , " & _
    " nvl(TGL03,0) daygradingluar_3 , " & _
    " nvl(TGL04,0) daygradingluar_4, " & _
    " nvl(TGL05,0) daygradingluar_5, " & _
    " nvl(TGL06,0) daygradingluar_6, " & _
    " nvl(TGL07,0) daygradingluar_7, " & _
    " nvl(TGL08,0) daygradingluar_8, " & _
	" nvl(TGL09,0) daygradingluar_9, " & _
    " nvl(TGL10,0) daygradingluar_10, " & _
    " nvl(TGL11,0) daygradingluar_11, " & _
    " nvl(TGL12,0) daygradingluar_12, " & _
    " nvl(TGL13,0) daygradingluar_13, " & _
    " nvl(TGL14,0) daygradingluar_14, " & _
    " nvl(TGL15,0) daygradingluar_15, " & _
    " nvl(TGL16,0) daygradingluar_16, " & _
    " nvl(TGL17,0) daygradingluar_17, " & _
    " nvl(TGL18,0) daygradingluar_18, " & _
    " nvl(TGL19,0) daygradingluar_19, " & _
    " nvl(TGL20,0) daygradingluar_20, " & _
    " nvl(TGL21,0) daygradingluar_21, " & _
    " nvl(TGL22,0) daygradingluar_22, " & _
    " nvl(TGL23,0) daygradingluar_23, " & _
    " nvl(TGL24,0) daygradingluar_24, " & _
    " nvl(TGL25,0) daygradingluar_25, " & _
    " nvl(TGL26,0) daygradingluar_26, " & _
    " nvl(TGL27,0) daygradingluar_27, " & _
    " nvl(TGL28,0) daygradingluar_28, " & _
    " nvl(TGL29,0) daygradingluar_29, " & _
    " nvl(TGL30,0) daygradingluar_30, " & _
    " nvl(TGL31,0) daygradingluar_31 " & _
	" from GRADINGTBSLUAR " & _
	" WHERE fcba = '" & businessunit & "' AND TAHUNBULAN = '" & TAHUNBULAN & "' "
	'Response.Write sqlgradingluar
		Set rsgradingluar = Server.CreateObject("ADODB.Recordset")
		rsgradingluar.Open sqlgradingluar, connstr

		while not rsgradingluar.eof

	%>

	<tr class=ratakanan>
		<td colspan= class=ratakanan><%=rsgradingluar("ket")%></td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsgradingluar("daygradingluar_31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>
	<%

		rsgradingluar.MoveNext
		WEND
		rsgradingluar.Close
		set rsgradingluar = Nothing
	%>

	<tr>
		<td colspan=35 align=left><strong>PRODUKSI DAN STOCK</strong></td>
	</tr>
	<tr>
		<td colspan=35 align=left><strong>CPO</strong></td>
	</tr>
	<tr>
		<td colspan= class=ratakanan><strong>Total Produksi</strong></td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILCPO_day31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>
	<tr>
		<td colspan= class=ratakanan>Rendemen</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCPO_day31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>
	<tr>
		<td colspan= align=left>Stock Akhir</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CPOTANKISI_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>	
	<tr>
		<td colspan=35 align=left><strong>KERNEL</strong></td>
	</tr>
	<tr>
		<td colspan= class=ratakanan><strong>Total Produksi</strong></td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("HASILKERNEL_day31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>
	<tr>
		<td colspan= class=ratakanan>Rendemen</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENKERNEL_day31"),2)%></td>
		<td colspan= class="ratakanan">0</td>
	</tr>
	<tr>
		<td colspan= align=left>Stock Akhir</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("PKTANKISI_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>

	<%
		dim sqlkualitascpo,rskualitascpo
		sqlkualitascpo = " SELECT TANKID, NVL(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.FFATANKCPO    , 0)),0) FFATANKCPO_day1, " & _
	   " NVL(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.AIRTANKCPO     , 0)),0) AIRTANKCPO_day1, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day1, " & _
	  
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.FFATANKCPO    , 0)) FFATANKCPO_day2, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.AIRTANKCPO     , 0)) AIRTANKCPO_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day2, " & _
	   
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.FFATANKCPO    , 0)) FFATANKCPO_day3, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.AIRTANKCPO     , 0)) AIRTANKCPO_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day3, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.FFATANKCPO    , 0)) FFATANKCPO_day4, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.AIRTANKCPO     , 0)) AIRTANKCPO_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day4, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.FFATANKCPO    , 0)) FFATANKCPO_day5, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.AIRTANKCPO     , 0)) AIRTANKCPO_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day5, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.FFATANKCPO    , 0)) FFATANKCPO_day6, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.AIRTANKCPO     , 0)) AIRTANKCPO_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day6, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.FFATANKCPO    , 0)) FFATANKCPO_day7, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.AIRTANKCPO     , 0)) AIRTANKCPO_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day7, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.FFATANKCPO    , 0)) FFATANKCPO_day8, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.AIRTANKCPO     , 0)) AIRTANKCPO_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day8, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.FFATANKCPO    , 0)) FFATANKCPO_day9, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.AIRTANKCPO     , 0)) AIRTANKCPO_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day9, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.FFATANKCPO    , 0)) FFATANKCPO_day10, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.AIRTANKCPO     , 0)) AIRTANKCPO_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day10, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.FFATANKCPO    , 0)) FFATANKCPO_day11, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.AIRTANKCPO     , 0)) AIRTANKCPO_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day11, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.FFATANKCPO    , 0)) FFATANKCPO_day12, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.AIRTANKCPO     , 0)) AIRTANKCPO_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day12, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.FFATANKCPO    , 0)) FFATANKCPO_day13, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.AIRTANKCPO     , 0)) AIRTANKCPO_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day13, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.FFATANKCPO    , 0)) FFATANKCPO_day14, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.AIRTANKCPO     , 0)) AIRTANKCPO_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day14, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.FFATANKCPO    , 0)) FFATANKCPO_day15, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.AIRTANKCPO     , 0)) AIRTANKCPO_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day15, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.FFATANKCPO    , 0)) FFATANKCPO_day16, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.AIRTANKCPO     , 0)) AIRTANKCPO_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day16, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.FFATANKCPO    , 0)) FFATANKCPO_day17, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.AIRTANKCPO     , 0)) AIRTANKCPO_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day17, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.FFATANKCPO    , 0)) FFATANKCPO_day18, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.AIRTANKCPO     , 0)) AIRTANKCPO_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day18, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.FFATANKCPO    , 0)) FFATANKCPO_day19, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.AIRTANKCPO     , 0)) AIRTANKCPO_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day19, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.FFATANKCPO    , 0)) FFATANKCPO_day20, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.AIRTANKCPO     , 0)) AIRTANKCPO_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day20, " & _
     
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.FFATANKCPO    , 0)) FFATANKCPO_day21, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.AIRTANKCPO     , 0)) AIRTANKCPO_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day21, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.FFATANKCPO    , 0)) FFATANKCPO_day22, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.AIRTANKCPO     , 0)) AIRTANKCPO_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day22, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.FFATANKCPO    , 0)) FFATANKCPO_day23, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.AIRTANKCPO     , 0)) AIRTANKCPO_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day23, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.FFATANKCPO    , 0)) FFATANKCPO_day24, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.AIRTANKCPO     , 0)) AIRTANKCPO_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day24, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.FFATANKCPO    , 0)) FFATANKCPO_day25, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.AIRTANKCPO     , 0)) AIRTANKCPO_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day25, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.FFATANKCPO    , 0)) FFATANKCPO_day26, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.AIRTANKCPO     , 0)) AIRTANKCPO_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day26, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.FFATANKCPO    , 0)) FFATANKCPO_day27, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.AIRTANKCPO     , 0)) AIRTANKCPO_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day27, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.FFATANKCPO    , 0)) FFATANKCPO_day28, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.AIRTANKCPO     , 0)) AIRTANKCPO_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day28, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.FFATANKCPO    , 0)) FFATANKCPO_day29, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.AIRTANKCPO     , 0)) AIRTANKCPO_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day29, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.FFATANKCPO    , 0)) FFATANKCPO_day30, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.AIRTANKCPO     , 0)) AIRTANKCPO_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day30, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.FFATANKCPO    , 0)) FFATANKCPO_day31, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.AIRTANKCPO     , 0)) AIRTANKCPO_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.KOTORANTANKCPO     , 0)) KOTORANTANKCPO_day31 " & _
      
	   " FROM PRODUCTIONTANKCPO P " & _
	   " WHERE FCBA='" & businessunit & "' AND TO_CHAR(P.PRODDATE, 'yyyymm') = " & tahunx & "||LPAD(" & bulanx & ", 2, 0) GROUP BY P.TANKID"
	   Set rskualitascpo = Server.CreateObject("ADODB.Recordset")
		rskualitascpo.Open sqlkualitascpo, connstr
	
	%>
	<tr>
		<td colspan=35 align=left><strong>KUALITAS CPO</strong></td>
	</tr>
		<%
		
			while not rskualitascpo.eof
		%>
	<tr>
		<td colspan= align=left>FFA St <%=rskualitascpo("TANKID")%></td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("FFATANKCPO_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>
	<tr>
		<td colspan= align=left>MOIST St <%=rskualitascpo("TANKID")%></td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("AIRTANKCPO_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>
	<tr>
		<td colspan= align=left>DIRT St <%=rskualitascpo("TANKID")%></td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitascpo("KOTORANTANKCPO_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>
	<%

		rskualitascpo.MoveNext
		WEND
		rskualitascpo.Close
		set rskualitascpo = Nothing
	%>
	<!------------------ BATAS KUALITAS ----------------------------------------------------->

	<%
	
		dim sqlkualitaspk,rskualitaspk
		sqlkualitaspk = " SELECT NVL(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.BULKINGSILOFFA    , 0)),0) BULKINGSILOFFA_day1, " & _
	   " NVL(SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.BULKINGSILOAIR     , 0)),0) BULKINGSILOAIR_day1, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day1, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 1,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day1, " & _
	  
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day2, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day2, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 2,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day2, " & _
	   
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day3, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day3, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 3,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day3, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day4, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day4, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 4,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day4, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day5, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day5, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 5,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day5, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day6, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day6, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 6,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day6, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day7, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day7, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 7,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day7, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day8, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day8, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 8,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day8, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day9, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day9, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 9,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day9, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day10, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day10, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 10,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day10, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day11, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day11, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 11,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day11, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day12, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day12, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 12,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day12, " & _
       
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day13, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day13, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 13,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day13, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day14, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day14, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 14,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day14, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day15, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day15, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 15,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day15, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day16, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day16, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 16,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day16, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day17, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day17, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 17,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day17, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day18, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day18, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 18,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day18, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day19, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day19, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 19,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day19, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day20, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day20, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 20,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day20, " & _
     
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day21, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day21, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 21,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day21, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day22, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day22, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 22,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day22, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day23, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day23, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 23,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day23, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day24, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day24, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 24,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day24, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day25, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day25, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 25,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day25, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day26, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day26, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 26,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day26, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day27, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day27, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 27,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day27, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day28, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day28, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 28,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day28, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day29, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day29, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 29,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day29, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day30, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day30, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 30,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day30, " & _
      
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.BULKINGSILOFFA    , 0)) BULKINGSILOFFA_day31, " & _
	   " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.BULKINGSILOAIR     , 0)) BULKINGSILOAIR_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.BULKINGSILOKOTORAN     , 0)) BULKINGSILOKOTORAN_day31, " & _
       " SUM(DECODE(TO_NUMBER(TO_CHAR(P.PRODDATE  , 'dd')), 31,P.BULKINGSILOBROKEN     , 0)) BULKINGSILOBROKEN_day31 " & _
      
	   " FROM PRODUCTIONBULKINGSILO P " & _
	   " WHERE FCBA='" & businessunit & "' AND TO_CHAR(P.PRODDATE, 'yyyymm') = " & tahunx & "||LPAD(" & bulanx & ", 2, 0)"
	   Set rskualitaspk = Server.CreateObject("ADODB.Recordset")
		rskualitaspk.Open sqlkualitaspk, connstr

	%>

	<tr>
		<td colspan=35 align=left><strong>KUALITAS KERNEL</strong></td>
	</tr>

	<%
		
			while not rskualitaspk.eof
		%>
	<tr>
		<td colspan= align=left>FFA</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOFFA_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>

	<tr>
		<td colspan= align=left>MOIST</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOAIR_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>
	<tr>
		<td colspan= align=left>DIRT</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOKOTORAN_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>
	<tr>
		<td colspan= align=left>BROKEN</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rskualitaspk("BULKINGSILOBROKEN_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>

	<%

		rskualitaspk.MoveNext
		WEND
		rskualitaspk.Close
		set rskualitaspk = Nothing
	%>
	<!------------------ BATAS KUALITAS KERNEL----------------------------------------------------->

	<tr>
		<td colspan=35 align=left><strong>CANGKANG</strong></td>
	</tr>
	<tr>
		<td colspan= align=left>CANGKANG PRODUKSI</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGPRODUKSI_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>

	<tr>
		<td colspan= align=left>RENDEMEN</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("RENDEMENCKG_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>

	<tr>
		<td colspan= align=left>PENGIRIMAN CANGKANG</td>
		<td colspan= class=ratakanan>%</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day1"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day2"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day3"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day4"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day5"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day6"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day7"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day8"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day9"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day10"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day11"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day12"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day13"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day14"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day15"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day16"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day17"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day18"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day19"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day20"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day21"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day22"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day23"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day24"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day25"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day26"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day27"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day28"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day29"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day30"),2)%></td>
		<td colspan= class="ratakanan"><%=FormatNumber(rsproduction("CANGKANGKIRIM_day31"),2)%></td>
		<td colspan= class=ratakanan>0</td>
	</tr>
		
	<!------------------ BATAS KUALITAS ckg----------------------------------------------------->
	<tr>
		<td colspan=35 align=left><strong></strong></td>
	</tr>
	<tr class=ratakiri>
		<td colspan= class=ratakiri>PRODUKSI JJG KOSONG</td>
		<td colspan= class=ratatengah>KG</td>
		<td colspan= class=ratakanan></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong1,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong2,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong3,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong4,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong5,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong6,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong7,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong8,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong9,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong10,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong11,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong12,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong13,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong14,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong15,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong16,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong17,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong18,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong19,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong20,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong21,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong22,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong23,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong24,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong25,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong26,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong27,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong28,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong29,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong30,2)%></td>
		<td colspan= class="ratakanan"><%= FormatNumber(m_jjgkosong31,2)%></td>
		
		<td colspan= class="ratakanan">0</td>
	</tr>

	<%
		no = no + 1
		rsproduction.MoveNext
		WEND
		rsproduction.Close
		set rsproduction = Nothing
	%>
	</tbody>
	</div>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</BODY>
	</div>
</HTML>	