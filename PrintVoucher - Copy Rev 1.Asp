<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
	<%
	function f_title(transtype_, banktype_)
		result = ""
		if Ucase(banktype_) = "JOURNAL" then result = "Journal Voucher" else result = PCase(transtype_) & " Voucher"
		
		f_title = result
	end function
	%>
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
width: 980px;
padding-top:30px;
padding-bottom:30px;
padding-left:30px;
padding-right:30px;
}
#header-wrapper {display:block}
.navi_to_print {
display:block;
margin: auto;
width: 980px;
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
.product-table{width: 21cm;}
#wrapper{background:#fff;padding:0px;}
#product-table td {font-size:10px; padding:2px;}
a.button_print {display:none;}
.letter { 
background-color:#fff; 
box-shadow:none; 
border-radius:0px;
padding:0px;
margin: none;
width: 100%;
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
dim txtTable,txtModule,v_txtDATECREATED,v_txtVOUCHERCODE,txtSQL,txtSelect,nCountofField,NO,v_Amount
txtTable = request("txtTable")
txtModule = request("txtModule")
v_txtDATECREATED= trim(request("DATECREATED"))
v_txtVOUCHERCODE=  trim(request("fcCode"))
transtype=  trim(request("transtype"))

if cstr(""&transtype) = "RECEIVE" Then
	txtSQL = " select p.datecreated" & _
			 "        , p.bankcode" & _
			 "        , bank_api.get_bankcode_name(p.bankcode,p.fcba) bankname" & _
			 "        , pay.vouchercode" & _
			 "        , pay.lineno" & _
			 "        , pay.jobcode" & _
			 "        , job.fcname" & _
			 "        , pay.locationtype" & _
			 "        , pay.locationcode" & _
			 "        , nvl(sum(pay.amount * -1),0) as amount" & _
			 "        , pay.reference" & _
			 "        , pay.remarks" & _
			 "        , p.remarks description" & _
			 "        , p.transtype" & _
			 " from paymentvoucherdetail pay" & _
			 "      , job job" & _
			 "      ,paymentvoucher p" & _
			 " where ( ((pay.jobcode = job.fccode) and (pay.fcba = job.fcba))" & _
			 "       and ((pay.vouchercode = p.vouchercode) and (pay.fcba = p.fcba)) )" & _
			 "       and pay.fcba = '" & session("fcBA") & "'"
			 
	if  len(trim(v_txtDATECREATED)) > 0 then
		txtSQL = txtSQL & " and p.datecreated >=to_Date('" & v_txtDATECREATED & "','MM/DD/YYYY')"
		txtSelect = " From " & FormatDateTime(v_txtDATECREATED,1)
	end if

	if  len(trim(v_txtVOUCHERCODE)) > 0 then
		txtSQL =txtSQL &  " and pay.VOUCHERCODE='" & v_txtVOUCHERCODE & "'"
		txtSelect = txtSelect & " Voucher No " & v_txtVOUCHERCODE
	end if 

	txtsql=(txtsql) & " group by p.datecreated" & _
					  "          , p.bankcode" & _
					  "          , BANK_API.get_Bankcode_name(p.bankcode,P.FCBA)" & _
					  "          , pay.vouchercode" & _
					  "          , pay.lineno" & _
					  "          , pay.jobcode" & _
					  "          , job.fcname" & _
					  "          , pay.locationtype" & _
					  "          , pay.locationcode" & _
					  "          , pay.REFERENCE" & _
					  "        	 , p.transtype" & _
					  "          , p.remarks" & _
					  "          , pay.remarks"
	txtsql	=	txtsql & " Order By p.datecreated, pay.vouchercode, lineno asc, amount Desc"
	if len(trim(txtSelect)) > 0 then
	   txtSelect = "Selected  by : " & txtSelect
	else
	   txtSelect = "Selected by All Data"
	end if
Else
	txtSQL = " select p.datecreated" & _
         "        , p.bankcode" & _
         "        , bank_api.get_bankcode_name(p.bankcode,p.fcba) bankname" & _
         "        , pay.vouchercode" & _
		 "        , pay.lineno" & _
         "        , pay.jobcode" & _
         "        , job.fcname" & _
         "        , pay.locationtype" & _
         "        , pay.locationcode" & _
         "        , nvl(sum(pay.amount),0) as amount" & _
         "        , pay.reference" & _
         "        , pay.remarks" & _
		 "        , p.remarks description" & _
         "        , p.transtype" & _
		 " from paymentvoucherdetail pay" & _
         "      , job job" & _
         "      ,paymentvoucher p" & _
		 " where ( ((pay.jobcode = job.fccode) and (pay.fcba = job.fcba))" & _
         "       and ((pay.vouchercode = p.vouchercode) and (pay.fcba = p.fcba)) )" & _
		 "       and pay.fcba = '" & session("fcBA") & "'"
		 
if  len(trim(v_txtDATECREATED)) > 0 then
    txtSQL = txtSQL & " and p.datecreated >=to_Date('" & v_txtDATECREATED & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDATECREATED,1)
end if

if  len(trim(v_txtVOUCHERCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.VOUCHERCODE='" & v_txtVOUCHERCODE & "'"
	txtSelect = txtSelect & " Voucher No " & v_txtVOUCHERCODE
end if 

txtsql=(txtsql) & " group by p.datecreated" & _
                  "          , p.bankcode" & _
                  "          , BANK_API.get_Bankcode_name(p.bankcode,P.FCBA)" & _
                  "          , pay.vouchercode" & _
				  "          , pay.lineno" & _
                  "          , pay.jobcode" & _
                  "          , job.fcname" & _
                  "          , pay.locationtype" & _
                  "          , pay.locationcode" & _
                  "          , pay.REFERENCE" & _
                  "        	 , p.transtype" & _
				  "          , p.remarks" & _
                  "          , pay.remarks"
txtsql	=	txtsql & " Order By p.datecreated, pay.vouchercode, lineno asc, amount Desc"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
End If
' response.Write(txtsql)
response.Flush()
Set rsCat = Server.CreateObject("ADODB.Recordset")
rsCat.CursorType = 3
rsCat.CursorLocation = 3
rsCat.Open txtSQL, connstr

if not rsCat.eof then
	  transtype = rsCat("transtype")
      vouchercode = rsCat("vouchercode")
      datecreated = rsCat("datecreated")
	  bankcode = rsCat("bankcode")
	  bankname=rsCat("bankname")
	  jobcode=rsCat("jobcode")
	  fcname = rsCat("fcname")
	  description = rsCat("description")
	  locationtype=rsCat("locationtype")
	  locationcode=rsCat("locationcode")
	  REFERENCE = rsCat("REFERENCE")
	  remarks = rsCat("remarks")
   End if
   rsCat.close
   set rsCat = Nothing 

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
#info_note .point .name { width:100px; float:left;}
#info_note .point .isi  { float:left; height:26px; line-height:25px;margin-top:7px; margin-left:10px;font-size:10px;}

#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point_right .name {width:25%; float:left;}
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
.terbilang{font-size:10px;float:left; text-align:left;font-style: italic;margin-left:20px; padding-top:10px;}
.title_terbilang {width:100px;float:left; text-align:left;font-style: italic;font-size:10px;}
.page{width:100%;}
}
@media print
{
td{padding:3px;}
table.PAYSLIP td{padding:3px;}
.page-break hr {border:none;display: none;}
.PAYSLIP td {font-family:verdana,sans-serif; font-size:7px; padding:0px;}
#note {margin-left:auto;padding:5px;}
#info_note .judul { text-align:center; font-size:14px;padding-top:0px;padding-top:0px;padding-bottom:0px;}
#info_note .left { float: left; width:35%;}
#info_note .right {  float: right; width:35%;}
#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:5px;}
#info_note .point span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point .name { width:100px; float:left;}
#info_note .point .isi  { float:left; height:26px; line-height:25px; margin-left:10px;margin-top:7px;font-size:9px;}

#info_note .point_right {font-size:10px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point_right .name {width:25%; float:left;}
#info_note .point_right .isi  { float:left; height:26px; line-height:25px;margin-top:7px; margin-left:10px;}
.clear { clear:both; content:'';}
table .long {width:400px;}
table .small {width:40px;}
table th span { font-style:italic; font-weight:normal;}
.line { height:1px; background:#000; content:' ';margin: auto;}
.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
#product-table th { background-color:#fff; color:#000;}
#product-table td { border-color:#000; color:#000; font-size:9px; }
#product-table td.none { border: none;}
#footer_note .block { font-size:10px; float:left; width:19%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.terbilang{font-size:10px;float:left; text-align:left;font-style: italic;margin-left:10px; padding-top:5px;}
.title_terbilang {width:100px;float:left; text-align:left;font-style: italic;font-size:10px;}
.page{width:6.6in;}
}
@page { 
 size:6.9in 5.5in;
 margin-left:0.15in; 
 margin-right:0.15in;
} /* width height */

</style>
<div class="page" >
	<header id="info_note" >
		<div class="judul"><strong><%=transtype%> VOUCHER</strong></div>
			<table width="100%">
			<!-- Left -->
			<td valign="top">
				<div class="point">
					<div class="name">
					Kebun/Pabrik<span>Business Unit</span>
					</div>
					<div class="isi">
					: <%=session("FCBA")%> - <%=Session("BA")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Kas/Bank<span>Cash/Bank</span>
					</div>
					<div class="isi">
					: <%=BANKCODE%> - <%=BANKNAME%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<%if cstr(""&transtype) = "PAYMENT" then%>
					<div class="name">
						Di Bayarkan Kepada<span>Paid To</span>
					</div>
					<div class="isi">
					: <%=description%> 
					</div>
					<%else if cstr(""&transtype) = "RECEIVE" then%>
					<div class="name">
						Di Terima Dari<span>Received From</span>
					</div>
					<div class="isi">
					: <%=description%> 
					</div>
					<%else if cstr(""&transtype) = "JOURNAL" then%>
					<div class="name">
						Di Jurnal Oleh<span>Journal By</span>
					</div>
					<div class="isi">
					: <%=description%> 
					</div>
					<%end if
					end if
					end if%>
				</div>
				<div class="clear"></div>	
			

			<!-- right -->
			<td valign="top">			
				<div class="point_right">
					<div class="name">
					Nomor<span>Number</span>
					</div>
					<div class="isi">
					: <%=VOUCHERCODE%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Tanggal<span>Date</span>
					</div>
					<div class="isi">
					: <%=DATECREATED%>
					</div>
				</div>
				<div class="clear"></div>
			</td>
		</table>
		

	</header>

<div class="clear"></div>

<table border="0" width="100%" id="product-table">
	<thead>
		
		<tr>
			<th rowspan="2" class="small">Nomor<div class="line"></div><span>Item</span></th>
			<th rowspan="2">Keperluan<div class="line long"></div><span>Particulars</span></th>
			<!-- <th colspan="3" valign="middle" align="center">Pembebanan</th> -->
			<th rowspan="2">Nilai<div class="line"></div><span>Amount</span></th>
		</tr>
		<!-- <tr>
			<th class="small" valign="middle" align="center">Jobcode</th>
			<th class="small" valign="middle" align="center">TT</th>
			<th class="small" valign="middle" align="center">TC</th>
		</tr> -->
	</thead>
	<tbody>
		<%
			Dim v_Opening, v_Debit, v_Credit, v_Saldo	
			Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.Open txtSQL, connstr
			if Not rsCat.eof then
				nCountofField=rsCat.fields.count-1
			end if
			NO	=	1
			v_Amount	=	0
			
			While Not rsCat.EOF
		%>
		<tr>
			<td>
				<%=NO%>
			</td>
			<td>
				<%=rsCat("REMARKS")%>
			</td>
			<td align="right">
				<%=FormatNumber(rsCat("AMOUNT"),0)%>
			</td>
		</tr>
		<%
				v_Amount= v_Amount + trim(cdbl(rsCat("Amount")))
				'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
				rsCat.MoveNext
				'response.Flush()
				' intRecordsPerPage = intRecordsPerPage - 1
		 		no=no+1
				Wend
				rsCat.close
				set rsCat = Nothing
				%>
		<tr>
			<td colspan="2" align="right">
				TOTAL&nbsp;&nbsp;&nbsp;&nbsp;<b>Rp.</b>
			</td>
			<td align="right">
				<%=FormatNumber(v_Amount,0)%>
			</td>
		</tr>
		<tr>
			<td colspan="4" align="right" style="padding:5px;">
			<%

		
			
			%>
			<div class="title_terbilang">
			The Sum Of :
			<hr/>
			Terbilang
			</div>
			<div class="terbilang" >
			<%=Terbilang(v_Amount)%> RUPIAH
			</div>
			</td>
		</tr>
	</tbody>
</table>
<div id="footer_note" style="width:95%;">	
	<div class="block">Received by,
        <br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
    <div class="block">Check No,
        <br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
    <div class="block">Cashier,
        <br/><br/><br/>
        <input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
	<div class="block">Checked by,
        <br/><br/><br/>
        <input class="isian" type="text" value="" >
        <hr/>
        <input class="isian" type="text" value="">
    </div>
	<div class="block">Approved by,
        <br/><br/><br/>
        <input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
 </div>
    
    
    <div class="clear"></div>


</div>
</div>
</div>
</div>

</body>
</html>






