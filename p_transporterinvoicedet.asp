<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Transporter Invoice Detail</title>
<link rel="stylesheet" href="css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src='js/x.js'></script>
<script language="JavaScript" src='js/xtableheaderfixed.js'></script>
        <script language="JavaScript">
        xAddEventListener(window, 'load',
          function() {
            new xTableHeaderFixed('xthf-site', window);
            new xTableHeaderFixed('xthf-green', window);
            new xTableHeaderFixed('xthf-blue', window);
          }, false
        );
        </script>
		
	<script language="JavaScript" src="js/jquery.printPage.js"></script> 
	<script>
					function toggle(id){
						var e = document.getElementById(id);
						if(e.style.display == 'inline-block'){
							e.style.display = 'none';
						}else if(e.style.display == 'none'){
							e.style.display = 'inline-block';
						}
					}
	</script>
	<script>
	 function printDiv(divName) {
	 var printContents = document.getElementById(divName).innerHTML;
     var originalContents = document.body.innerHTML;

     document.body.innerHTML = printContents;

     window.print();

     document.body.innerHTML = originalContents;
}
</script>
<style type="text/css">
@media all {
#product-table td{padding:5px;}
.page-break	{ display: block; }
a.button_print { display:block; }	
#wrapper{background: #e2e2e2; /* Old browsers */
background: -moz-linear-gradient(-45deg,  #e2e2e2 0%, #dbdbdb 50%, #d1d1d1 51%, #fefefe 100%); /* FF3.6+ */
background: -webkit-gradient(linear, left top, right bottom, color-stop(0%,#e2e2e2), color-stop(50%,#dbdbdb), color-stop(51%,#d1d1d1), color-stop(100%,#fefefe)); /* Chrome,Safari4+ */
background: -webkit-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Chrome10+,Safari5.1+ */
background: -o-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Opera 11.10+ */
background: -ms-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* IE10+ */
background: linear-gradient(135deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* W3C */
filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#e2e2e2', endColorstr='#fefefe',GradientType=1 ); /* IE6-9 fallback on horizontal gradient */
}
a.button_print {display: inline-block;}
.letter { 
background-color:#fff; 
box-shadow:0px 0px 4px rgba(0,0,0,0.4); 
border-radius:4px;
margin: auto;
width: 1200px;
padding-top:30px;
padding-bottom:30px;
padding-left:30px;
padding-right:30px;
}
#header-wrapper {display:block} 
} 
</style>		
</head>

<body>
<%
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

	v_txtFcba = Request("fcba") 
 	v_txtACCOUNTINGYEARPERIOD=trim(request("ACCOUNTINGYEARPERIOD"))
	invoicecode = trim(request("invoicecode")) 

	varDefaultBa = Session("fcBA")   
	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If "" & txttdate = "" Then txttdate = textDate(date)
	txttdate = txtdatefrom & "" & txtdateuntil

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
		
   	if  v_txtfcba <> "" then
       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
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
	
		strGrnInv	=	"   SELECT invoicecode," & _
						"          salesshipmentno," & _
						"          lineno," & _
						"          wbt_chitno," & _
						"          acceptedqtybruto," & _
						"          deliveredqty," & _
						"          varianceqty," & _
						"          amountbeforeloss," & _
						"          losspercentage," & _
						"          maxtolerance," & _
						"          penaltyqty," & _
						"          penaltyamount," & _
						"          amountafterloss," & _
						"          remarks" & _
						"     FROM transporterinvoicedtl" & _
						"    WHERE invoicecode ='"& invoicecode &"'" & _
						" ORDER BY invoicecode," & _
						"          lineno," & _
						"          wbt_chitno"

	'response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 

%>
<div id="wrapper">
<div class="navi_to_print">
<input type="button" value="Print" onclick="printDiv('note')" />

</div>
<div class="letter">
<div id="note">
<style>
@media all {
td{padding:5px;}
table.PAYSLIP td{padding:5px;}
.page-break hr {display:block;border: 1px dashed red;}
.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
#note {margin-left:auto;padding:5px;}
#info_note .judul { text-align:center; font-size:16px;padding-top:0px;padding-bottom:0px;}
#info_note .left { float: left; width:35%;}
#info_note .right {  float: right; width:35%;}
#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:5px;}
#info_note .point span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point .name { width:100px; float:left; margin-top :10px;}
#info_note .point .isi  { float:left; height:26px; line-height:25px;margin-top:7px; margin-left:10px;font-size:10px;}

#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point_right .name {width:25%; float:left;  margin-top :10px;}
#info_note .point_right .isi  { float:left; height:26px; line-height:25px; margin-left:10px;margin-top:7px;font-size:10px;}
.clear { clear:both; content:'';}
table .long {width:400px;}
table .small {width:40px;}
table th span { font-style:italic; font-weight:normal;}
.line { height:1px; background:#000; content:' ';margin: auto;}
.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
#product-table th { background-color:#fff; color:#000;}
#product-table td { border-color:#000; color:#000; font-size:9px;}
#product-table td.none { border: none;}
#footer_note .block { font-size:10px; float:left; width:25%; height:90px; background-color:#fff; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:90%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.terbilang{font-size:10px;float:left; text-align:left;font-style: italic;margin-left:20px; padding-top:10px;}
.title_terbilang {width:100px;float:left; text-align:left;font-style: italic;font-size:10px;}
.page{width:100%;}
}
@media print
{
td{padding:5px;}
table.PAYSLIP td{padding:5px;}
.page-break hr {display:block;border: 1px dashed red;}
.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
#note {margin-left:auto;padding:5px;}
#info_note .judul { text-align:center; font-size:16px;padding-top:0px;padding-bottom:0px;}
#info_note .left { float: left; width:35%;}
#info_note .right {  float: right; width:35%;}
#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:5px;}
#info_note .point span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point .name { width:100px; float:left; margin-top :10px;}
#info_note .point .isi  { float:left; height:26px; line-height:25px;margin-top:7px; margin-left:10px;font-size:10px;}

#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point_right .name {width:25%; float:left;  margin-top :10px;}
#info_note .point_right .isi  { float:left; height:26px; line-height:25px; margin-left:10px;margin-top:7px;font-size:10px;}
.clear { clear:both; content:'';}
table .long {width:400px;}
table .small {width:40px;}
table th span { font-style:italic; font-weight:normal;}
.line { height:1px; background:#000; content:' ';margin: auto;}
.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
#product-table th { background-color:#fff; color:#000;}
#product-table td { border-color:#000; color:#000; font-size:9px;}
#product-table td.none { border: none;}
#footer_note .block { font-size:10px; float:left; width:19%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.terbilang{font-size:10px;float:left; text-align:left;font-style: italic;  padding-top:10px;}
.title_terbilang {width:100px;float:left; text-align:left;font-style: italic;font-size:10px;}
.page{width:98.5%;}
}

</style>
<div class="page" >  
	<table border="0" width="100%" id="product-table">
		<thead>
			<tr>
				<th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
					   <div id="head_judul" style="float:left; text-align:left;">
							<h4>PT. <%=CompanyName%></h4>
							<h5><%=FcbaName%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;"> 
							<h1 class="judul">TRANSPORTER INVOICE DETAIL </h1> 
						</div> 	
						<div style="clear:both;"></div>
					</div>
				</th>
			</tr> 
			<!--<div class="printhide">--> 
			<tr>
				<th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
			</tr>
			<!--</div>--> 
				<th>Invoice Code</th>
				<th>Sales Shipment No</th>
				<th>Line No</th> 
				<th>WBT Chitno</th>
				<th>Delivery Qty</th>  
				<th>Accept Qty Bruto</th>
				<th>Variance Qty</th>
				<th>Amount Before Loss</th>
				<th>Loss Percentage</th>
				<th>Max Tolerance</th>
				<th>Penalty Qty</th>
				<th>Penalty Amount</th>
				<th>Amount After Loss </th>
				<th>Remarks</th> 
			</tr>
		</thead>
	<tbody>
<%
	 
	if Not rsGrnInv.eof then
	   	nCountofField=rsGrnInv.fields.count-1
		v_deliveredqty=0
		v_acceptedqtybruto=0
		v_varianceqty = 0
		v_amountbeforeloss = 0
		v_losspercentage = 0
		v_maxtolerance= 0
		v_penaltyqty= 0
		v_penaltyamount= 0
		v_amountafterloss= 0
	end if 

	while not rsGrnInv.eof 
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
	%> 
		<tr class="<%=cClass%>">
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("invoicecode")%></td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("salesshipmentno")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("lineno")%></td> 
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("wbt_chitno")%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rsGrnInv("deliveredqty"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rsGrnInv("acceptedqtybruto"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rsGrnInv("varianceqty"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rsGrnInv("amountbeforeloss"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rsGrnInv("losspercentage"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rsGrnInv("maxtolerance"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rsGrnInv("penaltyqty"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rsGrnInv("penaltyamount"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rsGrnInv("amountafterloss"))),2)%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("remarks")%></td>
		</tr>
<% 			
			v_deliveredqty= v_deliveredqty + Trim(cdbl(to_nol(rsGrnInv("deliveredqty"))))
			v_acceptedqtybruto= v_acceptedqtybruto + Trim(cdbl(to_nol(rsGrnInv("acceptedqtybruto"))))
			v_varianceqty = v_varianceqty + Trim(cdbl(to_nol(rsGrnInv("varianceqty"))))
			v_amountbeforeloss = v_amountbeforeloss + Trim(cdbl(to_nol(rsGrnInv("amountbeforeloss"))))
			v_losspercentage = v_losspercentage + Trim(cdbl(to_nol(rsGrnInv("losspercentage"))))
			v_maxtolerance = v_maxtolerance + Trim(cdbl(to_nol(rsGrnInv("maxtolerance"))))
			v_penaltyqty = v_penaltyqty + Trim(cdbl(to_nol(rsGrnInv("penaltyqty"))))
			v_penaltyamount = v_penaltyamount + Trim(cdbl(to_nol(rsGrnInv("penaltyamount"))))
			v_amountafterloss = v_amountafterloss + Trim(cdbl(to_nol(rsGrnInv("amountafterloss"))))
			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>		<tr bgcolor="#F4F4A0" style="font-weight: bold;">
			<td align="Center" style="<%=cmdstyleLeft%>" colspan="4"> <font face="Verdana" size="1">T o t a l</td> 
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_deliveredqty,2)%></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_acceptedqtybruto,2)%></td> 
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_varianceqty,2)%></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amountbeforeloss,2)%></td> 
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_losspercentage,2)%></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_maxtolerance,2)%></td> 
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_penaltyqty,2)%></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_penaltyamount,2)%></td> 
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amountafterloss,2)%></td> 
			<td align="Right" style="<%=cmdstyleLeft%>"  > </td>
		</tr>
	</tbody>  
</table> 
</body>
</html>






