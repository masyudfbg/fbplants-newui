<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Invoice</title>
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
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>

<%
Dim fcba,tahun,kegiatan1,kegiatan2
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

CompanyName= trim(request("company"))
fcba= trim(request("fcba"))
tahun= trim(request("tahun"))
kegiatan1= trim(request("kegiatan1"))
kegiatan2= trim(request("kegiatan2"))


Dim sqlview,rsview 
sqlview = "select h.unitkerja,h.fieldcode,h.activitycode,j.fccode || ' - ' || j.fcname coaname,h.rate,sum(d.qty) qtymaterial,d.ratestock,sum(h.mandays) mandays, " & _
" nvl(sum(h.boronganpremi),0) boronganpremi from monthlyworkplan h,job j,MONTHLYWORKPLANDET d " & _
" where h.activitycode=j.fccode and h.fcba=j.fcba AND h.rkbno=d.rkbno and h.fcba=d.fcba and h.PERIOD_YEAR='" & tahun & "' and j.fccode='" & kegiatan1 & "' " & _ 
" group by h.unitkerja,h.fieldcode,h.activitycode,j.fccode || ' - ' || j.fcname,h.rate,(d.qty),d.ratestock  order by h.unitkerja asc"
set rsview = connStr.Execute(sqlview)
'response.Write(sqlview)

%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcba%></h5>
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
					<h2>BUDGET PROSES</h2>
				</div>
			</div>
				</th>
			</tr>

						<tr>
							<th rowspan="2">Kegiatan</th>
							<th rowspan="2">Divisi</th>
							<th rowspan="2">Blok</th>
							<th colspan="2">HK</th>
							<th colspan="">BORONGAN</th>
							<th colspan="2">MATERIAL</th>
						</tr>
						<tr>	
							<th>QTY</th>
							<th>Rp</th>
							<th>Rp</th>
							<th>QTY</th>
							<th>Rp</th>
						</tr>
                        
						<tbody>
						<%
						
						dim boronganpremi,qtymaterial,ratestock,rpmaterial, ttlqtyhk,ttlrphk,ttlrpborongan,ttlqtymaterial,ttlrpmaterial
						ttlqtyhk = 0
						ttlrphk	 = 0
						ttlrpborongan = 0
						ttlqtymaterial = 0
						ttlrpmaterial=0
						While Not rsview.EOF

							dim hk,rate,rphk
							hk = rsview("mandays")
							rate = rsview("rate")
							boronganpremi = rsview("boronganpremi")
							qtymaterial = rsview("qtymaterial")
							ratestock = rsview("ratestock")
							rphk = cdbl(hk)*cdbl(rate)
							rpmaterial = cdbl(qtymaterial)*cdbl(ratestock)
						%>
						<tr>					
							<td><%=rsview("coaname")%></td>							
							<td><%=rsview("unitkerja")%></td>							
							<td><%=rsview("fieldcode")%></td>							
							<td style="text-align: right;"><%=FormatNumber(cdbl(to_nol(rsview("mandays"))),2)%></td>							
							<td style="text-align: right;"><%=FormatNumber(cdbl(rphk),2)%></td>							
							<td style="text-align: right;"><%=FormatNumber(boronganpremi)%></td>
							<td style="text-align: right;"><%=FormatNumber(cdbl(qtymaterial),2)%></td>							
							<td style="text-align: right;"><%=FormatNumber(cdbl(rpmaterial),2)%></td>							
						</tr>
						<%
						
						ttlqtyhk = cdbl(ttlqtyhk) + cdbl(hk)
						ttlrphk	 = cdbl(ttlrphk)+cdbl(rphk)
						ttlrpborongan = cdbl(ttlrpborongan)+cdbl(boronganpremi)
						ttlqtymaterial = cdbl(ttlqtymaterial)+cdbl(qtymaterial)
						ttlrpmaterial= cdbl(ttlrpmaterial)+cdbl(rpmaterial)

						rsview.MoveNext
						Wend
						rsview.close()
						
						%>
						</tbody>
						<tr>
							<td colspan="3" style="text-align: right;font-weight: bold;">Total</td>
							<td  style="text-align: right;font-weight: bold;"><%=ttlqtyhk%></td>
							<td  style="text-align: right;font-weight: bold;"><%=FormatNumber(cdbl(ttlrphk),2)%></td>
							<td  style="text-align: right;font-weight: bold;"><%=FormatNumber(cdbl(ttlrpborongan),2)%></td>
							<td  style="text-align: right;font-weight: bold;"><%=FormatNumber(cdbl(ttlqtymaterial),2)%></td>
							<td  style="text-align: right;font-weight: bold;"><%=FormatNumber(cdbl(ttlrpmaterial),2)%></td>
						</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	