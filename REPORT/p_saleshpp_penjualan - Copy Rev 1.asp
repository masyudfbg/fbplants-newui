<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
maxCol = 0
add_col_bud_prod_qty = ""
d_span = 26
	
v_txtFcba = Request("txtFcba")
v_txtYear = Request("txtYear")
v_txtmonth = Request("txtMonth")
v_txtDate = Request("txtDate")

If "" & v_txtDate = "" Then v_txtDate = TextDate(date)

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Costbook"))
end if

if  v_txtfcba <> "--ALL--"  and v_txtfcba <> "" then
	whereCons = whereCons & " And fcba = '" & v_txtfcba & "' "
end if
' selectCons = selectCons & ",fcba"
' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
set rsCat = connStr.Execute(q)
if not rsCat.eof then
	fcbaname = rsCat("fcName")
	CompanyName=rsCat("FCCOMPANYNAME")
	MgrName = rsCat("FCMANAGER")
	KTUName = rsCat("FCKTU")
	CropName=rsCat("FCCROP")
	FcbaType=rsCat("fcType")
else
    nSpanTotal = nSpanTotal - 1 
end if

strDefaultBa = 	" SELECT" & _
				" FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'M'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

arr_field_eof = true
strPusingan	= 	"   SELECT wb.fcba," & _
				"		   '1' judul," & _
				"          'Dari TBS Sendiri' source," & _
				"          (SUM (wb.weightin) - SUM (wb.weightout)) AS qty" & _
				"     FROM weighbridgeticket wb" & _
				"          INNER JOIN" & _
				"          (SELECT fcba, customercode" & _
				"             FROM customer" & _
				"            WHERE     fcba = 'PSR'" & _
				"                  AND supplier = 'Y'" & _
				"                  AND customercode IN ('KS00001', 'KS00002')) cs" & _
				"             ON (cs.fcba = wb.fcba AND cs.customercode = wb.cust_customercode)" & _
				"    WHERE     wb.fcba = 'PSR'" & _
				"          AND TO_NUMBER (TO_CHAR (wb.dateout, 'mm')) <= '09'" & _
				"          AND TO_NUMBER (TO_CHAR (wb.dateout, 'yyyy')) = '2019'" & _
				"          AND wb.product_productcode IN ('TBS')" & _
				" GROUP BY wb.fcba" & _
				" UNION ALL" & _
				"   SELECT wb.fcba," & _
				"		   '1' judul," & _
				"          'Dari TBS Luar' source," & _
				"          (SUM (wb.weightin) - SUM (wb.weightout)) AS qty" & _
				"     FROM weighbridgeticket wb" & _
				"          INNER JOIN" & _
				"          (SELECT fcba, customercode" & _
				"             FROM customer" & _
				"            WHERE     fcba = 'PSR'" & _
				"                  AND supplier = 'Y'" & _
				"                  AND customercode NOT IN ('KS00001', 'KS00002')) cs" & _
				"             ON (cs.fcba = wb.fcba AND cs.customercode = wb.cust_customercode)" & _
				"    WHERE     wb.fcba = 'PSR'" & _
				"          AND TO_NUMBER (TO_CHAR (wb.dateout, 'mm')) <= '09'" & _
				"          AND TO_NUMBER (TO_CHAR (wb.dateout, 'yyyy')) = '2019'" & _
				"          AND wb.product_productcode IN ('TBS')" & _
				" GROUP BY wb.fcba" & _
				" UNION ALL" & _
				"   SELECT mp.fcba,'2' judul, 'CPO' source, SUM (mp.production_today) AS qty" & _
				"     FROM mill_production mp" & _
				"    WHERE     mp.fcba = 'PSR'" & _
				"          AND TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy')) = 2019" & _
				"          AND TO_NUMBER (TO_CHAR (mp.tdate, 'mm')) = 09" & _
				"          AND mp.product_code IN ('CPO')" & _
				" GROUP BY mp.fcba" & _
				" UNION ALL" & _
				"   SELECT mp.fcba, '2' judul, 'KER' source, SUM (mp.production_today) AS qty" & _
				"     FROM mill_production mp" & _
				"    WHERE     mp.fcba = 'PSR'" & _
				"          AND TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy')) = 2019" & _
				"          AND TO_NUMBER (TO_CHAR (mp.tdate, 'mm')) = 09" & _
				"          AND mp.product_code IN ('KER')" & _
				" GROUP BY mp.fcba" 

response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

' If Not qORS.EOF Then
' 	arr_field = qORS.getrows()
' 	arr_field_eof = false
' End If
' qORS.Close
' set qORS = Nothing

%>

<%	
    strTitle = 	" select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
			    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
				" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
				" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
				" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
				" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
						
	Set rsTitle = Server.CreateObject("ADODB.Recordset")
	'response.write rsTitle
	rsTitle.Open strTitle, connstr, 3,3	

	if not rsTitle.EOF then
		no = 0
		titleGroupPlantation = rsTitle ("titleGP")			
		titlePlantation = rsTitle ("titleP")
		titleGroupRegion = rsTitle ("titleGR")
		titleRegion = rsTitle ("titleR")
		titleCompany = rsTitle ("titleC")
		titleSite = rsTitle ("titleS")
	else
		titleGroupPlantation = "Group Plantation"			
		titlePlantation = "Plantation"
		titleGroupRegion = "Group Region"
		titleRegion = "Region"
		titleCompany = "Company"
		titleSite = "Site"
	end if
	rsTitle.Close
	set rsTitle = Nothing
			
%>
<html>

<head>
<title>ePlaS Report - Sales Harga Pokok Penjualan</title>
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
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<script>
 $(function() {
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}

th.rotate-270 {
  height: 200px; /* 1 */
  white-space: nowrap;
  min-width:40px;
  max-width:40px;
  overflow:hidden;
}
th.rotate-270 > div {
	transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-webkit-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-ms-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
    width:15px; /* 4 */
	font-size:10px;
	font-weight:normal;
}

/*
1. Mengatur tinggi table head.
2. mengatur posisi X div
3. mengatur posisi Y div.
4. mengatur lebarnya DIV yang akan mempengaruhi Lebarnya Table Head.
*/
</style>


<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
</head>

<body>
	<div id="wrapper">
	 <!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
				<div class="printhide">
					<table id="form-table" class="centered-margin">
						<tr>
							<td>
								<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
									<%
									Dim rsFcba
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.CursorLocation = adUseClient
									strsql =    " SELECT '--ALL--' as fcba from DUAL UNION " & _  
									            " select fccode fcba from businessunit " & _
												" where " & _
												" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL AND fcType = 'M'" 
									  
									
									rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsFcba.eof then
										DO while not rsFcba.eof
											if v_txtFcba = rsFcba("fcba") then
											%>			
												<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
											<% Else %>	
												<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
											<% End If
											
											rsFcba.movenext
										loop
									End if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">TAHUN</font>
							</td>
								<% 
									Dim mon, Yr, MinYr, MaxYr, i, MM
									mon = Month(Session("v_Costbook"))
									If len(v_txtMonth)=0 then v_txtMonth=mon

									Yr = Year(Session("v_Costbook"))
									MinYr = 2015
									MaxYr = Yr  + 2
									If len(v_txtYear)= 0 then v_txtYear = Yr
								%>
							<td width="4">
								<select name="txtYear">
									<% For i = MinYr to MaxYr %>
										<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
									<% Next %>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">BULAN</font>
							</td>
							<td width="2">
								<select name="txtMonth">
								<% 
									For i = 1 to 12		
										if i <= 9 then 
											iBulan = 0&i
										else
											iBulan = i
										end if
										i	= iBulan
								%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
								<%
									Next
								%>
								</select>
							</td> 
							<td>
								<input type="submit" value="View" />
							</td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
		<tr>
	        <th colspan="<%=d_span%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header-wrapper"> 
					<div id="subheader" class="centered">
						<h2>PT. <%=Session("Company")%></h2>
						<h3>BEBAN POKOK PENJUALAN MINYAK SAWIT DAN INTI SAWIT</h3>
						<h3>UNTUK TAHUN YANG BERAKHIR PADA TANGGAL-TANGGAL</h3> 
						<h4><%=v_txtdate%></h4>
					</div>
		        </div>  
			</th>
		</tr> 
    </thead> 
</table>
<div style="margin-left: 200px;">
	<h4>I. HASIL PRODUKSI MINYAK SAWIT DAN INTI SAWIT</h4>
</div>
<br> 
<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<font face="verdana" size="3">Pemakaian Tandan Buah Segar</font>
	<font face="verdana" size="3" style="float: right; margin-right: 250px;">KG</font>
</div>
<br>
<% if Not qORS.eof then
	   NO = 1
	   nCountofField = qORS.fields.count - 1
	end if

	While Not qORS.EOF
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass = "default"
		end if	

	Currjudul = qORS("judul")
	If Currjudul <> Prevjudul Then
%>
		<%If Len(Trim(Prevjudul)) = 0 Then 
		    totaljudul = 0 
		%>
		<%Else%> 
			<br>
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3" style="font-weight: bold;"> Jumlah Tandan Buah Segar yang Diproses</font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px; font-weight: bold;"><%=formatnumber(cdbl(to_nol(totaljudul)),0)%></font>
			</div>
			<br>
				<hr>
			<br> 
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3">Hasil Produksi</font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px;">KG</font>
			</div>
			<br>
		<%
			totaljudul = 0 
		%> 
		<%End If%>
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3"><%=qORS("source")%></font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px;"><%=formatnumber(cdbl(to_nol(qORS("qty"))),0)%></font>
			</div>  
		<%
			totaljudul = totaljudul + cdbl(to_nol(qORS("qty"))) 
		%>
	<%Else%>
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3"><%=qORS("source")%></font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px;"><%=formatnumber(cdbl(to_nol(qORS("qty"))),0)%></font>
			</div>  
		<%
			totaljudul = totaljudul + cdbl(to_nol(qORS("qty"))) 
		%>
<%End If%>

	
<%	qORS.MoveNext
 	Prevjudul= Currjudul
	Wend
	qORS.close()
	qORS=null%>
			<br>
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3" style="font-weight: bold;"> Jumlah Hasil Produksi Minyak Sawit dan Inti Sawit</font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px; font-weight: bold;"><%=formatnumber(cdbl(to_nol(totaljudul)),0)%></font>
			</div> 
		<br>
			<hr>
		<br>

<%
	strPusingan2	= 	"   SELECT wb.fcba," & _
					"		   '1' judul," & _
					"          'Dari TBS Sendiri' source," & _
					"          (SUM (wb.weightin) - SUM (wb.weightout)) AS qty" & _
					"     FROM weighbridgeticket wb" & _
					"          INNER JOIN" & _
					"          (SELECT fcba, customercode" & _
					"             FROM customer" & _
					"            WHERE     fcba = 'PSR'" & _
					"                  AND supplier = 'Y'" & _
					"                  AND customercode IN ('KS00001', 'KS00002')) cs" & _
					"             ON (cs.fcba = wb.fcba AND cs.customercode = wb.cust_customercode)" & _
					"    WHERE     wb.fcba = 'PSR'" & _
					"          AND TO_NUMBER (TO_CHAR (wb.dateout, 'mm')) <= '09'" & _
					"          AND TO_NUMBER (TO_CHAR (wb.dateout, 'yyyy')) = '2019'" & _
					"          AND wb.product_productcode IN ('TBS')" & _
					" GROUP BY wb.fcba" & _
					" UNION ALL" & _
					"   SELECT wb.fcba," & _
					"		   '1' judul," & _
					"          'Dari TBS Luar' source," & _
					"          (SUM (wb.weightin) - SUM (wb.weightout)) AS qty" & _
					"     FROM weighbridgeticket wb" & _
					"          INNER JOIN" & _
					"          (SELECT fcba, customercode" & _
					"             FROM customer" & _
					"            WHERE     fcba = 'PSR'" & _
					"                  AND supplier = 'Y'" & _
					"                  AND customercode NOT IN ('KS00001', 'KS00002')) cs" & _
					"             ON (cs.fcba = wb.fcba AND cs.customercode = wb.cust_customercode)" & _
					"    WHERE     wb.fcba = 'PSR'" & _
					"          AND TO_NUMBER (TO_CHAR (wb.dateout, 'mm')) <= '09'" & _
					"          AND TO_NUMBER (TO_CHAR (wb.dateout, 'yyyy')) = '2019'" & _
					"          AND wb.product_productcode IN ('TBS')" & _
					" GROUP BY wb.fcba" & _
					" UNION ALL" & _
					"   SELECT mp.fcba,'2' judul, 'CPO' source, SUM (mp.production_today) AS qty" & _
					"     FROM mill_production mp" & _
					"    WHERE     mp.fcba = 'PSR'" & _
					"          AND TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy')) = 2019" & _
					"          AND TO_NUMBER (TO_CHAR (mp.tdate, 'mm')) = 09" & _
					"          AND mp.product_code IN ('CPO')" & _
					" GROUP BY mp.fcba" & _
					" UNION ALL" & _
					"   SELECT mp.fcba, '2' judul, 'KER' source, SUM (mp.production_today) AS qty" & _
					"     FROM mill_production mp" & _
					"    WHERE     mp.fcba = 'PSR'" & _
					"          AND TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy')) = 2019" & _
					"          AND TO_NUMBER (TO_CHAR (mp.tdate, 'mm')) = 09" & _
					"          AND mp.product_code IN ('KER')" & _
					" GROUP BY mp.fcba" 

	response.write strPusingan2
	Set qORS = Server.CreateObject("ADODB.Recordset")
	qORS.Open strPusingan2, connstr, adOpenStatic
%>
<div style="margin-left: 200px;">
	<h4>II. MUTASI STOCK MINYAK SAWIT DAN INTI SAWIT</h4>
</div>
<br> 
<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<font face="verdana" size="3">Pemakaian Tandan Buah Segar</font>
	<font face="verdana" size="3" style="float: right; margin-right: 250px;">KG</font>
</div>
<br>
<% if Not qORS.eof then
	   NO = 1
	   nCountofField = qORS.fields.count - 1
	end if

	While Not qORS.EOF
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass = "default"
		end if	

	Currjudul = qORS("judul")
	If Currjudul <> Prevjudul Then
%>
		<%If Len(Trim(Prevjudul)) = 0 Then 
		    totaljudul = 0 
		%>
		<%Else%> 
			<br>
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3" style="font-weight: bold;"> Jumlah Tandan Buah Segar yang Diproses</font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px; font-weight: bold;"><%=formatnumber(cdbl(to_nol(totaljudul)),0)%></font>
			</div>
			<br>
				<hr>
			<br> 
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3">Hasil Produksi</font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px;">KER (KG)</font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px;">CPO (KG)</font>
			</div>
			<br>
		<%
			totaljudul = 0 
		%> 
		<%End If%>
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3"><%=qORS("source")%></font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px;"><%=formatnumber(cdbl(to_nol(qORS("qty"))),0)%></font>
			</div>  
		<%
			totaljudul = totaljudul + cdbl(to_nol(qORS("qty"))) 
		%>
	<%Else%>
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3"><%=qORS("source")%></font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px;"><%=formatnumber(cdbl(to_nol(qORS("qty"))),0)%></font>
			</div>  
		<%
			totaljudul = totaljudul + cdbl(to_nol(qORS("qty"))) 
		%>
<%End If%>

	
<%	qORS.MoveNext
 	Prevjudul= Currjudul
	Wend
	qORS.close()
	qORS=null%>
			<br>
			<div style="margin-left: 200px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font face="verdana" size="3" style="font-weight: bold;"> Jumlah Hasil Produksi Minyak Sawit dan Inti Sawit</font>
				<font face="verdana" size="3" style="float: right; margin-right: 250px; font-weight: bold;"><%=formatnumber(cdbl(to_nol(totaljudul)),0)%></font>
			</div> 
		<br>
			<hr>
		<br>
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

