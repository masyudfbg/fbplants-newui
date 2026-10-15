<%'Option Explicit
%>
<%'On Error GOTO CmdErrorlog 
%>
<%response.buffer=True%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Print Invoice PO</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
        <script language="JavaScript">
        xAddEventListener(window, 'load',
          function() {
            new xTableHeaderFixed('xthf-site', window);
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
	 //var printContents = document.getElementById(divName).innerHTML;
     //var originalContents = document.body.innerHTML;

     //document.body.innerHTML = printContents;

     window.print();

     //document.body.innerHTML = originalContents;
}
</script>
<style type="text/css">
@media all {
#product-table td{padding:5px;padding-bottom: 0; border-top:0px;border-bottom:0px;
    padding-top: 0;}
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
width: 98%;
padding-top:30px;
padding-bottom:30px;
padding-left:1%;
padding-right:1%;
}
#header-wrapper {display:block}
.navi_to_print {
display:block;
margin: auto;
width: 98%;
padding-top:30px;
}
}
@media print
{
table td{padding:3px;}
#header-wrapper {display:none}
.button_print {display: none;}
.printhide{display:none;}
.page-break	{ display: block; page-break-before: always; }
.product-table{width: 100%;}
#wrapper{background:#fff;padding:none;}
#product-table th {font-size:14px; }
#product-table td {font-size:14px; padding:2px;padding-bottom: 0;
    padding-top: 0;order-top:0px;border-bottom:0px;}
a.button_print {display:none;}
.letter { 
background-color:#fff; 
box-shadow:none; 
border-radius:0px;
width: 100%;
margin: none;
padding-top:0px;
padding-bottom:0px;
padding-left:0px;
padding-right:0px;
}
.navi_to_print {
display:none;
margin: auto;
width: 100%;
padding-top:30px;
}
}
</style>
</head>

<body>
<%
' Const adOpenForwardOnly = 0
' Const adLockReadOnly = 1
' Const adCmdTableDirect = &H0200
' Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleStock="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; width:100px"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>

 <%
 Dim txtNoReceved,txtDateReceived,txtRemarks,q,nCountofField,no,i,txtRetensi
 Dim Remarks,dptxt

   Dim rsReceived,intRecordsPerPage,cColor,txtSupplier
   set rsReceived = server.CreateObject("Adodb.recordset")
   q="select * from contractorinvoicehdr  " & _
     " where INVOICECODE='" & request("fcCode") & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsReceived = Server.CreateObject("ADODB.Recordset")
   rsReceived.Open q, connstr
   if not rsReceived.eof then
      txtNoReceved = rsReceived("INVOICECODE")
	  txtDateReceived = rsReceived("TRANSDATE")
	  txtSupplier = cmdFindTitle_BA("Contractor","fcCode",rsReceived("CONTRACTORCODE"))
	  txtHeaderRemarks="From Invoice Date : " & rsReceived("TRANSDATE")
	  txtRetensi= rsReceived("isretention_invoice")
   End if
   rsReceived.close
   set rsReceived = Nothing%>

<div id="wrapper">
	<div class="navi_to_print">
	<input type="button" value="Print" onclick="printDiv('note')"/>
</div>

<div class="letter">
	<div id="note">
		<style>
		@media all {
		td{padding:5px;}
		table.PAYSLIP td{padding:1px;}
		.page-break hr {display:block;border: 1px dashed red;}
		.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
		#note {margin-left:auto;padding:5px;}
		#info_note .judul { text-align:center; font-size:20px;}
		#info_note .left { float: left; width:35%;}
		#info_note .right {  float: right; width:35%;}
		#info_note .point { font-size:11px; font-weight:bold; width:100%;  margin-top:2px;}
		#info_note .point span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point .name { width:100px; float:left; margin-top: 10px;}
		#info_note .point .isi  { float:left; height:36px; line-height:35px; margin-left:10px;}

		#info_note .point_right { font-size:11px; font-weight:bold; width:100%;  margin-top:2px; text-align:left;}
		#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point_right .name { width:100px; float:left; margin-top: 10px;}
		#info_note .point_right .isi  { float:left; height:36px; line-height:35px; margin-left:10px;}
		.clear { clear:both; content:'';}
		table .long {width:400px;}
		table .small {width:40px;}
		table th span { font-style:italic; font-weight:normal;}
		.line { height:1px; background:#000; content:' ';margin: auto;}
		.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
		#product-table th { background-color:#fff; color:#000;border:1px #000000 solid ;}
		#product-table td { border-color:#000; color:#000;padding-bottom: 0;order-top:0px;border-bottom:0px;
    padding-top: 0;}
		#product-table td.none { border: none;}
		#footer_note .block { font-size:10px; float:left; width:19%; height:150px; background-color:#fff; margin-left:1%; text-align:center;}
		.tembusan{font-size:10px;}
		input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
		hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
		hr{ margin-top: 5px;}
		}
		
		</style>
		
		<header id="info_note" >
				<div class="judul"><strong>CONTRACTOR INVOICE</strong></div>
					<table width="100%" >
					<!-- Left -->
					<td valign="top">
					<div class="point">
							<div class="name">
								PT
							</div>
							<div class="isi">
							: <%=session("company")%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point">
							<div class="name">
							Kebun/Pabrik
							</div>
							<div class="isi">
							: <%=session("FCBA")%> - <%=Session("BA")%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point">
							<div class="name">
							Supplier
							</div>
							<div class="isi">
							: <%=txtSupplier%>
							</div>
						</div>
						<div class="clear"></div>
					<!-- right -->
					<td width="400px" valign="top">			
						<div class="point_right">
							<div class="name">
							Nomor 
							</div>
							<div class="isi">
							: <%=txtNoReceved%> <%if txtRetensi="TRUE" Then %>(RETENSI)<%end if%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point_right">
							<div class="name">
							Tanggal 
							</div>
							<div class="isi">
							:  <%=txtDateReceived%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point_right">
							<div class="name">
							</div>
							<div class="isi">
							</div>
						</div>
					</td>
					
				</table>
		</header>
		<div class="clear"></div>
	        <style>
			.group_checkbox {margin-bottom:10px;margin-top:25px;}
			.group_checkbox ul {width:100%;}
			.group_checkbox ul li { 
				list-style-type:none;
				float:left;
			}
			.group_checkbox ul li p{
				font-weight:bold;
				font-size:12px;
				float:left;
				display:initial;
			}
			.group_checkbox ul li span{
				float:left;
				margin-top:0px;
				margin-left:10px;
			}
			.group_checkbox ul li span img{
				width:21px;
			}
		</style>
		 
		<table border="0" width="100%" id="product-table" style="border-bottom:1px #000000 solid;">
			<thead>
				<tr>
					<th class="small" style="border-bottom:1px #000000 solid;">
					Nomor
					</th>
					<th>Tanggal</th>
					<th>Agreement Code</th>
					<th style="font-size:10px">JOB<div class="line"></div><span>Code</span></th>
					<th style="font-size:10px">JOB<div class="line"></div><span>Nama</span></th>
					<th style="font-size:10px">Quantity</th>
					<th style="font-size:10px">Rate</th>
					<th style="font-size:10px">Amount</th>					
					</tr>
			 </thead>
			 <tbody>
	

   <%
   recordTotalAmount = " SELECT NVL(inv.totalamount_afterretention,0) amountbeforevat" & _
						"        , NVL(inv.grandtotal_aftertax, 0) totalAmount" & _
						"        , NVL(inv.amountppn, 0) totalppn" & _
						"        , NVL(inv.amountpph, 0) totalpph" & _
						"        , NVL(inv.dp, 0) dp" & _
						" FROM contractorinvoicehdr inv" & _
						" WHERE inv.invoicecode = '" & request("fcCode") & "'" & _ 
						"       AND inv.fcba = '" & session("fcBA") & "'" 
						
	' response.Write recordTotalAmount
	Set rsrecordTotalAmount = Server.CreateObject("ADODB.Recordset")
	rsrecordTotalAmount.Open recordTotalAmount, connstr

	if Not rsrecordTotalAmount.Eof then
		p_amountbeforevat = cdbl(to_nol(rsrecordTotalAmount("amountbeforevat")))		
		p_totalamount = cdbl(to_nol(rsrecordTotalAmount("totalAmount")))
		p_totalppn = cdbl(to_nol(rsrecordTotalAmount("totalppn")))
		p_totalpph = cdbl(to_nol(rsrecordTotalAmount("totalpph")))
		dptxt = cdbl(to_nol(rsrecordTotalAmount("dp")))
    Else
		p_amountbeforevat = 0
		p_totalppn = 0
		p_totalpph = 0
		p_totalamount = 0
		dptxt = 0
	End if
	
   Dim rsPurchaseReceive,nQty,nAmount
   set rsPurchaseReceive = server.CreateObject("Adodb.recordset")
   q="SELECT inv.transdate,invd.agreementcode,invd.jobcode,JOB_API.get_jobcode_name(invd.jobcode,invd.fcba) ||' - '|| ct.remarks jobname,nvl(sum(invd.totalfinal),0) as QUANTITY," & _
     "nvl((invd.rate),0) as Rate,nvl(sum(invd.AMOUNT),0) as Amount FROM contractorinvoicehdr inv INNER JOIN contractorinvoicedtl invd on (inv.INVOICECODE=invd.INVOICECODE and inv.fcba=invd.fcba ) " & _ 
	 "INNER JOIN contractor sp on (inv.contractorcode=sp.fccode and inv.fcba=sp.fcba) " & _
	 "INNER JOIN contractagreementdetail ct ON (ct.fcba=invd.fcba AND ct.agreementcode=invd.agreementcode and invd.agreementlineno=ct.lineno ) " & _
     "WHERE  " & _
     " inv.INVOICECODE='" & request("fcCode") &  "' and " & _
	 " inv.fcBA='" & session("fcBA") & "'" & _
	 " group by inv.transdate ,invd.agreementcode,invd.jobcode,invd.fcba,ct.remarks,invd.rate " & _
	 " order by invd.jobcode,invd.rate "
  ' response.Write q	
   Set rsPurchaseReceive = Server.CreateObject("ADODB.Recordset")
   rsPurchaseReceive.CursorType = 3
   rsPurchaseReceive.CursorLocation = 3
   rsPurchaseReceive.Open q, connstr%>
   <%
   if Not rsPurchaseReceive.eof then
      nCountofField=rsPurchaseReceive.fields.count-1
	  no=1
	  nQty=0
	  nAmount=0
	  While Not rsPurchaseReceive.EOF

		  nQty = nQty+CDBL(rsPurchaseReceive("QUANTITY"))
		  nAmount = nAmount+CDBL(rsPurchaseReceive("AMOUNT"))
		  'if no  mod 2 = 1 then
          '   cColor = "#EEEEEE"
		  'else
		  '   cColor = "#E5E5E5"
		  'end if
		  %>
		   <tr>
					  <td align="center" style="<%=cmdstyleLeft%>"><%=no%>&nbsp;</td>
					  <td align="center"  style="<%=cmdstyleStock%>" title="" valign="middle"><%=rsPurchaseReceive("transdate")%>&nbsp;</td>
					  <td align="center"  style="<%=cmdstyleStock%>" title="" valign="middle"><%=rsPurchaseReceive("agreementcode")%>&nbsp;</td>
					  <td align="LEFT"  style="width:50px;<%=cmdstyle%>" title="" valign="middle"><%=rsPurchaseReceive("jobcode")%>&nbsp;</td>
					  <td align="LEFT"  style="width:200px;<%=cmdstyle%>" title="" valign="middle"><font face="Verdana"><%=rsPurchaseReceive("jobname")%></td>
					  <td align="Right"  style="width:50px;<%=cmdstyle%> " title="" valign="middle"><%=FormatNumber(rsPurchaseReceive("QUANTITY"),2)%>&nbsp;</td>
					  <td align="Right"  style="width:50px;<%=cmdstyle%> " title="" valign="middle"><%=FormatNumber(rsPurchaseReceive("Rate"),2)%>&nbsp;</td>
					  <td align="Right"  style="width:50px;<%=cmdstyle%> " title="" valign="middle"><%=FormatNumber(rsPurchaseReceive("Amount"),2)%>&nbsp;</td>					  
		  </tr>
		  <%
		  rsPurchaseReceive.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  
	
  <%
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseReceive.Close
   set rsPurchaseReceive = Nothing
   
   
   %>
   	</tbody>
	</table>
		
   	
	 <table border="1" width="100%" id="product-table" >
	 <thead>
	  <tr>
	  <td colspan="6"  align="right"><font face="Verdana">SUB Total</td>
	  <td style="width:110px" align="right"><font face="Verdana"><%=FbpNum2(CheckValue(nAmount))%>&nbsp;</td>
	  </tr>	
	  <tr>
	  <td colspan="6" align="right"><font face="Verdana">VAT</td>
	  <td style="width:110px;<%=cmdstyle%>" align="right"><font face="Verdana"><%=CheckValue(p_totalppn)%>&nbsp;</td>
	  </tr>				    
	  <tr>
	  <td colspan="6"  align="right"><font face="Verdana">PPH</td>
	  <td style="width:110px;<%=cmdstyle%>" align="right"><font face="Verdana"><%=CheckValue(p_totalpph)%>&nbsp;</td>
	  </tr>				    
	  <tr>
	  <td colspan="6"  align="right"><font face="Verdana">DP</td>
	  <td style="width:110px;<%=cmdstyle%>" align="right"><font face="Verdana"><%=CheckValue(dptxt)%>&nbsp;</td>
	  </tr>				    
	  <tr>
	  <td colspan="6"  align="right"><font face="Verdana">TOTAL</td>
	  <td style="width:110px;<%=cmdstyle%>" align="right"><font face="Verdana"><%=CheckValue(p_totalamount)%>&nbsp;</td>
	  </tr>				    
   	</tbody>
	</table>
	
		   <br/>

</body>
</html>
 