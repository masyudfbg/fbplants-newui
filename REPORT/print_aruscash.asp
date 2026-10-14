<%validatesession%>
<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Payment Voucher</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
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
		
	<script language="JavaScript" src="/js/jquery.printPage.js"></script> 
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

v_txtDATECREATED_From= trim(request("DATECREATED_From"))
v_txtDATECREATED_To=trim(request("DATECREATED_To"))
v_txtACCOUNTINGYEARPERIOD=trim(request("ACCOUNTINGYEARPERIOD"))
v_txtBANKCODE=trim(request("BANKCODE"))
txtControlJob=""
nOpening=0
' session.lcid = 0421    1057


q="SELECT CONTROLJOB FROM BANK WHERE FCCODE ='" & v_txtBANKCODE & _
  "' AND  FCBA= '" & Session("fcBA") & "'"
Set rsCARI = Server.CreateObject("ADODB.Recordset")
rsCARI.CursorType = 3
rsCARI.CursorLocation = 3
rsCARI.Open q, connstr	 
if Not rsCARI.eof Then
   txtControlJob=rsCARI("CONTROLJOB")
else
   txtControlJob=""   
End if
rsCari.Close
set rsCari=nothing

q="SELECT nvl(sum(OPENING_BALANCE),0) as OPENING_BALANCE  FROM TRIALBALANCE WHERE FCCODE ='" & txtControlJob & _
  "' AND  FCBA= '" & Session("fcBA") & "'" & _
  "  AND PERIOD_MONTH = to_char(to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY'),'MM')" & _
  "  AND PERIOD_year = to_char(to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY'),'YYYY')"
Set rsCARI = Server.CreateObject("ADODB.Recordset")
rsCARI.CursorType = 3
rsCARI.CursorLocation = 3
' response.Write q
response.Flush()
rsCARI.Open q, connstr	 
if Not rsCARI.eof Then
   nOpening=rsCARI("OPENING_BALANCE")
else
   nOpening=0
End if
rsCari.Close
set rsCari=nothing

' Amount diminuskan
' Karena kalau debet adalah kas bekurang
' Kalau Kredit maka kas bertambah
'-------------------------------- Control Job  -----------------------------------------------------------------------------------
txtSQL1 =  	 "  SELECT p.datecreated," & _ 
			 "                    p.bankcode," & _ 
			 "                    p.transtype," & _ 
			 "                    pay.vouchercode," & _
			 "                    p.remarks orang," & _			 
			 "                    pay.jobcode," & _ 
			 "                    job.fcname jobname," & _ 
			 "                    pay.locationtype," & _ 
			 "                    pay.locationcode," & _ 
			 "                    NVL (SUM (pay.amount), 0) AS debet," & _ 
			 "                    0 as credit," & _ 
			 "                    pay.REFERENCE," & _ 
			 "                    pay.remarks" & _ 
			 "               FROM paymentvoucherdetail pay, job job, paymentvoucher p" & _ 
			 "              WHERE     (    (    (pay.jobcode = job.fccode)" & _ 
			 "                              AND (pay.FCBA = job.FCBA))" & _ 
			 "                         AND (    (pay.vouchercode = p.vouchercode)" & _ 
			 "                              AND (pay.FCBA = p.FCBA)))" & _ 
			 "                    AND pay.fcba = '"& session("fcba") &"'" & _ 
			 "                    AND p.transtype = '01-RECEIVE'" & _ 
			 "                    AND p.rowstateid = 'Approved'" 

if  len(trim(v_txtDATECREATED_From)) > 0 then
    txtSQL1 = txtSQL1 & " and p.datecreated >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect1 = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
    txtSQL1 =txtSQL1 &  " and p.datecreated <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect1 = txtSelect1 & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtBANKCODE)) > 0 then
    txtSQL1 = txtSQL1 & " and (p.BANKCODE='" & v_txtBANKCODE & "')"
	txtSelect1 = txtSelect1 & " Bank Code " & v_txtBANKCODE
end if

if  len(trim(v_txtACCOUNTINGYEARPERIOD)) > 0 then
    txtSQL1 = txtSQL1 & " and p.accountingyear||lpad(p.accountingperiod,2,0) =" & v_txtACCOUNTINGYEARPERIOD 
	txtSelect1 = txtSelect1 & " , ACCOUNTINGYEAR PERIOD " & left(v_txtACCOUNTINGYEARPERIOD,4) & " - " & Right(v_txtACCOUNTINGYEARPERIOD,2)
end if
txtsql1=(txtsql1) &  "           GROUP BY p.datecreated," & _ 
					 "                    p.bankcode," & _ 
					 "                    p.transtype," & _ 
					 "                    pay.vouchercode," & _
					 "                    p.remarks," & _	
					 "                    pay.jobcode," & _ 
					 "                    job.fcname," & _ 
					 "                    pay.locationtype," & _ 
					 "                    pay.locationcode," & _ 
					 "                    pay.REFERENCE," & _ 
					 "                    pay.remarks" 

'-------------------------------- Jurnal ----------------------------------------------------------------------------------------

txtSQL2 =  	 "             SELECT p.datecreated," & _ 
			 "                    p.bankcode," & _ 
			 "                    p.transtype," & _ 
			 "                    pay.vouchercode," & _ 
			 "                    p.remarks orang," & _	
			 "                    pay.jobcode," & _ 
			 "                    job.fcname jobname," & _ 
			 "                    pay.locationtype," & _ 
			 "                    pay.locationcode," & _ 
			 "                    0 as debet," & _ 
			 "                    NVL (SUM (pay.amount), 0) AS credit," & _ 
			 "                    pay.REFERENCE," & _ 
			 "                    pay.remarks" & _ 
			 "               FROM paymentvoucherdetail pay, job job, paymentvoucher p" & _ 
			 "              WHERE     (    (    (pay.jobcode = job.fccode)" & _ 
			 "                              AND (pay.FCBA = job.FCBA))" & _ 
			 "                         AND (    (pay.vouchercode = p.vouchercode)" & _ 
			 "                              AND (pay.FCBA = p.FCBA)))" & _ 
			 "                    AND pay.fcba = '"& session("fcba") &"'" & _ 
			 "                    AND p.transtype = '02-PAYMENT'" & _ 
			 "                    AND p.rowstateid = 'Approved'" 
			 

if  len(trim(v_txtDATECREATED_From)) > 0 then
    txtSQL2 = txtSQL2 & " and p.datecreated >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect2 = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
    txtSQL2 =txtSQL2 &  " and p.datecreated <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect2 = txtSelect2 & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtBANKCODE)) > 0 then
    txtSQL2 = txtSQL2 & " and (p.BANKCODE='" & v_txtBANKCODE & "')"
	txtSelect2 = txtSelect2 & " Bank Code " & v_txtBANKCODE
end if
if  len(trim(v_txtACCOUNTINGYEARPERIOD)) > 0 then
    txtSQL2 = txtSQL2 & " and p.accountingyear||lpad(p.accountingperiod,2,0) =" & v_txtACCOUNTINGYEARPERIOD 
end if
txtsql2=(txtsql2) &  "           GROUP BY p.datecreated," & _ 
					 "                    p.bankcode," & _ 
					 "                    p.transtype," & _ 
					 "                    pay.vouchercode," & _ 
					 "                    p.remarks," & _	
					 "                    pay.jobcode," & _ 
					 "                    job.fcname," & _ 
					 "                    pay.locationtype," & _ 
					 "                    pay.locationcode," & _ 
					 "                    pay.REFERENCE," & _ 
					 "                    pay.remarks"  

txtsql = "   SELECT vouchercode," & _ 
		 "          datecreated," & _ 
		 "          orang," & _ 
		 "          bankcode," & _ 
		 "          jobcode," & _ 
		 "          jobname," & _ 
		 "          transtype," & _ 
		 "          locationtype," & _ 
		 "          locationcode," & _ 
		 "          reference," & _ 
		 "          remarks," & _ 
		 "          SUM (DECODE (transtype, '01-RECEIVE', debet, 0)) debet," & _ 
		 "          SUM (DECODE (transtype, '02-PAYMENT', credit, 0)) credit" & _ 
		 "     FROM ( "& txtsql1 &" union all "& txtsql2 &")" & _
		 " GROUP BY vouchercode," & _ 
		 "          datecreated," & _ 
		 "          bankcode," & _ 
		 "          transtype," & _ 
		 "          jobcode," & _ 
		 "          orang," & _ 
		 "          jobname," & _ 
		 "          locationtype," & _ 
		 "          locationcode," & _ 
		 "          reference," & _ 
		 "          remarks ORDER BY datecreated, transtype, vouchercode, bankcode " 

txtSelect = txtSelect1
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected All Data"
end if
' response.Write(txtsql)
' response.Flush()

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
				<th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
					   <div id="head_judul" style="float:left; text-align:left;">
							<h4>PT. <%=Session("COMPANY")%></h4>
							<h5><%=Session("BA")%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h1 class="judul">ARUS CASH</h1>
						</div> 	
						<div style="clear:both;"></div>
					</div>
				</th>
			</tr> 
			<!--<div class="printhide">--> 
			<tr>
				<th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
			</tr>
			<!--</div>--> 
				<th>Date</th>
				<th> Voucher Code </th>
				<th> Paid For / Receive From</th>
				<th> Reference</th>
				<th> Remarks</th>
				<th> Debet</th>
				<th> Credit</th>
				<th> Saldo</th>
			</tr>
		</thead>
	<tbody>
		<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo	
	if  len(trim(v_txtACCOUNTINGYEARPERIOD)) > 0 then
        Set rsCat = connStr.Execute(txtSQL)
	
	
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1
    v_Amount = 0
    v_Saldo = 0%>
	<tr>
		<td align="left" style="<%=cmdstyleLeft%>" colspan="7"><font face="Verdana" size="1">Opening Balance.........</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nOpening,2)%></td>
	</tr>
	
	<%
	v_Amount=cdbl(to_nol(nOpening))
	v_Saldo=cdbl(to_nol(nOpening))
	While Not rsCat.EOF
	v_Saldo=v_Saldo + cdbl(to_nol(rsCat("debet"))) - cdbl(to_nol(rsCat("credit")))
	%>
		<tr> 
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Datecreated")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Vouchercode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("orang")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Reference")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Remarks")%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("debet"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("credit"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_Saldo,2)%></td>
		</tr>
	  <%
	  v_Amount=v_Amount+trim(cdbl(to_nol(rsCat("debet")))) - trim(cdbl(to_nol(rsCat("credit"))))
	  ' response.write v_Amount
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr>
	<td align="left" style="<%=cmdstyleLeft%>" colspan="7"> <font face="Verdana" size="1">Ending Balance.........</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Amount,2)%></td>
	</tr>
	</tbody>
		
    <%		
	end if %>
	</tbody>
</table>
	<div id="footer_note" style="width:100%; margin-left:10%;">	
		<div class="block"> <font face="Verdana" size="1" style="font-weight: bold;">Prepared By,
			<br/><br/> 
			<font face="Verdana" size="1" style="font-weight: bold;"><br/><br/>KASIR
		</div>
		<div class="block">Checked by,
			<br/><br/> 
			<font face="Verdana" size="1" style="font-weight: bold;"><br/><br/><%=Session("KTU")%>
		</div>
		<div class="block">Approved by,
			<br/><br/>
			<font face="Verdana" size="1" style="font-weight: bold;"><br/><br/><%=Session("FCMANAGER")%>
		</div>
	</div> 
	<div class="clear"></div> 
</body>
</html>






