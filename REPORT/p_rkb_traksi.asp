<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtMonth, v_txtYear
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim ScriptName, v_txtsectionname, v_txtGangCode, v_txtStatus
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName, no

v_txtYear		= Request("tahun")
v_txtMonth		= Request("bulan")
v_txtfcba 			= Request("fcba")




strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE'" & _
               "       and fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("fcba")
Else
	varDefaultBa = Session("fcba")
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			"     , FCMANAGER" & _
			"     , FCKTU" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varManagerName = rsBaInfo("FCMANAGER")
	varKTUName = rsBaInfo("FCKTU")
Else
	varFcbaName = ""
	varCompanyName = ""
	varManagerName = ""
	varKTUName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

txtComp         = Request("txtCompany")



			
%>
<html>
<head>
	<title>FBPlants Reporting - Rencana Kerja Bulanan Traksi </title>
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
	<script language="JavaScript" src="/js/date-picker.js"></script> 
	<script language="JavaScript">
	
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
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
</head>

<body>
	<div id="wrapper">	
		<div id="header-wrapper">
			<div class="printhide">
			<input type="button" value="Print" onClick="window.print()" style="font-family: 'FontAwesome';margin-top:-10px;">
				<!--#include file="../print_excel.asp"-->
				
			</div>
			
		</div>
		
	<table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
			<tr>
				<th colspan="30" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
						<h4>PT. <%=varCompanyName%></h4>
						<h5><%=varFcbaName%></h5>
					</div>
					<div id="subheader" class="centered">
						<h2>LAPORAN RENCANA KERJA BULANAN TRAKSI</h2>
						<br/>
						<h4>
							PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%>
						</h4> 

					</div>
					</div>
				</th>
			</tr>
		
		<tr>	
			<th rowspan="2" style="color:#fff;">No.</th>
			<th rowspan="2" style="color:#fff;">No. RKB</th>
			<th rowspan="2" style="color:#fff;">Vehicle Code</th>
			<th rowspan="2" style="color:#fff;">KM/HM</th>
			<th rowspan="2" style="color:#fff;">Job Code</th>
			<th rowspan="2" style="color:#fff;">Job Name</th>
			<th rowspan="2" style="color:#fff;">UOM</th>
			<th rowspan="2" style="color:#fff;">Norm. Jarak</th> 
			<th rowspan="2" style="color:#fff;">Norm. Qty</th> 
			<th rowspan="2" style="color:#fff;">Rate</th> 
			<th rowspan="2" style="color:#fff;">Amount</th> 
			
		</tr>
		
	</thead>
	
	<%
		
		dim sqldtl1,rsdtl1
		dim ttljarak,ttlqty,ttlamount

		sqldtl1 = "Select RH.RKBNO,RH.SECTIONCODE VEHICLECODE,RH.KMHM, R.ACTCODE,J.FCNAME,NVL(R.JARAK,0)JARAK , NVL(R.QTY ,0)QTY , NVL(R.RATE,0)RATE,  " & _ 
		" NVL(R.AMOUNT ,0)AMOUNT ,R.LINENO,J.UNITOFMEASUREMENT UOM " & _ 
		" from RKBHDR RH " & _
		" LEFT JOIN RKBVRADTL R ON RH.RKBNO=R.RKBNO AND RH.FCBA=R.FCBA " & _ 
		" LEFT JOIN JOBMASTER J ON R.ACTCODE=J.FCCODE " & _ 
		" WHERE RH.PERIOD_YEAR='" & v_txtYear & "'  and RH.PERIOD_MONTH='" & v_txtMonth & "' AND RH.DOCUMENT_CLASSIFICATION='502'  AND RH.fcba='" & v_txtfcba & "' " & _ 
		" ORDER BY  RH.RKBNO,RH.SECTIONCODE"
		set rsdtl1 = connStr.Execute(sqldtl1)
		'Response.Write sqldtl1
		no = 1
		ttljarak = 0
		ttlqty = 0
		ttlamount = 0
		While Not rsdtl1.EOF
			
	%>

		<tr>
			<td align="center"><%=no%></td>
			<td><%=rsdtl1("RKBNO")%></td>
			<td><%=rsdtl1("VEHICLECODE")%></td> 
			<td align="left"><%=rsdtl1("KMHM")%></td>
			<td align="left"><%=rsdtl1("ACTCODE")%></td>
			<td align="left"><%=rsdtl1("FCNAME")%></td>
			<td align="left"><%=rsdtl1("UOM")%></td>
			<td align="right"><%=formatnumber(rsdtl1("JARAK"),0)%></td>
			<td align="right"><%=formatnumber(rsdtl1("QTY"),0)%></td>
			<td align="right"><%=formatnumber(rsdtl1("RATE"),2)%></td>
			<td align="right"><%=formatnumber(rsdtl1("AMOUNT"),2)%></td>
			
		</tr>
		
	<%
			ttljarak = cdbl(ttljarak) + cdbl(rsdtl1("JARAK"))
			ttlqty	 = cdbl(ttlqty) +  cdbl(rsdtl1("QTY"))
			ttlamount = cdbl(ttlamount) +  cdbl(rsdtl1("AMOUNT"))
			no = no + 1 
			rsdtl1.MoveNext

			Wend
	%>

		<tr style="background-color:#27AE60;">	 
			<td align="center" colspan="7" ><strong>Grand Total</strong></td>
			<td align="right"><strong><%=formatnumber(ttljarak,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(ttlqty,0)%></strong></td>
			<td align="right"><strong></strong></td>
			<td align="right"><strong><%=formatnumber(ttlamount,0)%></strong></td>
			
			
		</tr> 
</table> 
<script src="/plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		
	</script>
   </body>
   </html>