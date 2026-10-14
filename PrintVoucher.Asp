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
fcba=  trim(request("fcba"))
' session.lcid = 0421    1057

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
			 "        , p.referenceno" & _
			 "        , p.cekno" & _
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
			 "       and pay.fcba = '" & fcba & "'"
			 
	if  len(trim(v_txtDATECREATED)) > 0 then
		txtSQL = txtSQL & " and p.datecreated >=to_Date('" & v_txtDATECREATED & "','MM/DD/YYYY')"
		txtSelect = " From " & FormatDateTime(v_txtDATECREATED,1)
	end if

	if  len(trim(v_txtVOUCHERCODE)) > 0 then
		txtSQL =txtSQL &  " and pay.VOUCHERCODE like '%" & v_txtVOUCHERCODE & "%'"
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
					  "          , p.referenceno" & _
				  	  "          , p.cekno" & _
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
         "        , p.referenceno" & _
		 "        , p.cekno" & _
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
		 "       and pay.fcba = '" & fcba & "'"
		 
if  len(trim(v_txtDATECREATED)) > 0 then
    txtSQL = txtSQL & " and p.datecreated >=to_Date('" & v_txtDATECREATED & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDATECREATED,1)
end if

if  len(trim(v_txtVOUCHERCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.VOUCHERCODE like '%" & v_txtVOUCHERCODE & "%'"
	txtSelect = txtSelect & " Voucher No " & v_txtVOUCHERCODE
end if 

txtsql=(txtsql) & " group by p.datecreated" & _
                  "          , p.bankcode" & _
                  "          , BANK_API.get_Bankcode_name(p.bankcode,P.FCBA)" & _
                  "          , pay.vouchercode" & _
				  "          , pay.lineno" & _
				  "          , p.referenceno" & _
				  "          , p.cekno" & _
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
	  cekno = rsCat("cekno")
	  referenceno = rsCat("referenceno")
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
#info_note .point .name { width:100px; float:left;  margin-top :10px;}
#info_note .point .isi  { float:left; height:26px; line-height:25px; margin-left:10px;margin-top:7px;font-size:9px;}

#info_note .point_right {font-size:10px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point_right .name {width:25%; float:left;  margin-top :10px;}
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
		<% if transtype = "01-RECEIVE" Then %>
			<div class="judul"><strong>RECEIVED VOUCHER</strong></div>
		<%Else%>
			<div class="judul"><strong><%=mid(transtype,4)%> VOUCHER</strong></div>
		<%End If%>
			<table width="100%">
			<!-- Left -->
			<td valign="top">
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
					Nomor 
					</div>
					<div class="isi">
					: <%=VOUCHERCODE%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Ref.No 
					</div>
					<div class="isi">
					: <%=referenceno%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Cek.No 
					</div>
					<div class="isi">
					: <%=cekno%>
					</div>
				</div>
				<div class="clear"></div>	
			<!-- right -->
			<td valign="top">			
				<div class="point_right">
					<div class="name">
					Kas/Bank 
					</div>
					<div class="isi">
					: <%=BANKCODE%> - <%=BANKNAME%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Tanggal 
					</div>
					<div class="isi">
					: <%=day(datecreated) & " " & monthname(month(datecreated),true) & " " & year(datecreated)%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<%if cstr(""&transtype) = "02-PAYMENT" then%>
					<div class="name">
						<p Font-size: 16px, line-height: 20px>
							Di Bayarkan Kepada 
						</p>
					</div>
					<div class="isi">
					: <%=description%> 
					</div>
					<%else if cstr(""&transtype) = "01-RECEIVE" then%>
					<div class="name">
						Di Terima Dari 
					</div>
					<div class="isi">
					: <%=description%> 
					</div>
					<%else if cstr(""&transtype) = "03-JOURNAL" then%>
					<div class="name">
						Di Jurnal Oleh 
					</div>
					<div class="isi">
					: <%=description%> 
					</div>
					<%end if
					end if
					end if%>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Dibayarkan Dengan
					</div>
					<div class="isi">
					: Cash/Tunan
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
			<td align="center">
				<%=NO%>
			</td>
			<td>
				<%=rsCat("REMARKS")%>
			</td>
			<td align="right">
				<%=FormatNumber(rsCat("AMOUNT"),2)%>
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
				<%=FormatNumber(v_Amount,2)%>
			</td>
		</tr>
		<tr>
			<td colspan="4" align="right" style="padding:5px;">
			<%
				Function Val( myString )
				' Val Function for VBScript (aka ParseInt Function in VBScript).
				' By Denis St-Pierre.
				' Natively VBScript has no function to extract numbers from a string.
				' Based shamelessly on MS' Helpfile example on RegExp object.
				' CAVEAT: Returns only the *last* match found
				'         (or, with objRE.Global = False, only the *first* match)
				 
					Dim colMatches, objMatch, objRE, strPattern
				 
					' Default if no numbers are found
					Val = 0
				 
					strPattern = "[-+0-9]+"       ' Numbers positive and negative; use
												  ' "ˆ[-+0-9]+" to emulate Rexx' Value()
												  ' function, which returns 0 unless the
												  ' string starts with a number or sign.
					Set objRE = New RegExp        ' Create regular expression object.
					objRE.Pattern    = strPattern ' Set pattern.
					objRE.IgnoreCase = True       ' Set case insensitivity.
					objRE.Global     = True       ' Set global applicability:
												  '   True  => return last match only,
												  '   False => return first match only.
					Set colMatches = objRE.Execute( myString )  ' Execute search.
					For Each objMatch In colMatches             ' Iterate Matches collection.
						Val = objMatch.Value
					Next
					Set objRE= Nothing
				End Function
				 
				Function Terbilang(Angka)
				   Dim n1
					Dim nangka, Strjmlhuruf, intpecahan  , strpecahan, urai, Bil1, strtot, bil2
					Dim x, Y, z, i  
				   
					For i = 1 To Len(Angka)
						If Not Mid(Angka, i, 1) = "." Then
							nangka = nangka + Mid(Angka, i, 1)
						End If
					Next
				   
					If Len(LTrim(nangka)) = 0 Then
						terbit = "Nol Rupiah"
						Exit Function
					End If
					Strjmlhuruf = LTrim(RTrim((nangka)))
					intpecahan = Val(Right(Mid(Angka, 15, 2), 2))
					If (intpecahan = 0) Then
						strpecahan = ""
					End If
					x = 0
					Y = 0
					urai = ""
				   
					While (x < Len(Strjmlhuruf))
						x = x + 1
						strtot = Mid(Strjmlhuruf, x, 1)
						Y = Y + Val(strtot)
						z = Len(Strjmlhuruf) - x + 1
				   
						Select Case Val(strtot)
						Case 1
							If (z = 1 Or z = 7 Or z = 10 Or z = 13) Then
								Bil1 = "Satu "
							ElseIf (z = 4) Then
								If (x = 1) Then
									Bil1 = "Se"
								Else
									Bil1 = "Satu "
								End If
							ElseIf (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
								x = x + 1
								strtot = Mid(Strjmlhuruf, x, 1)
								z = Len(Strjmlhuruf) - x + 1
								bil2 = ""
				   
							Select Case Val(strtot)
							Case 0
								Bil1 = "Sepuluh "
							Case 1
								Bil1 = "Sebelas "
							Case 2
								Bil1 = "Dua Belas "
							Case 3
								Bil1 = "Tiga Belas "
							Case 4
								Bil1 = "Empat Belas "
							Case 5
								Bil1 = "Lima Belas "
							Case 6
								Bil1 = "Enam Belas "
							Case 7
								Bil1 = "Tujuh Belas "
							Case 8
								Bil1 = "Delapan Belas "
							Case 9
								Bil1 = "Sembilan Belas "
							End Select
							Else
								Bil1 = "Se"
							End If
				   
						Case 2
							Bil1 = "Dua "
						Case 3
							Bil1 = "Tiga "
						Case 4
							Bil1 = "Empat "
						Case 5
							Bil1 = "Lima "
						Case 6
							Bil1 = "Enam "
						Case 7
							Bil1 = "Tujuh "
						Case 8
							Bil1 = "Delapan "
						Case 9
							Bil1 = "Sembilan "
						Case Else
							Bil1 = ""
						End Select
				   
					If (Val(strtot) > 0) Then
						If (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
							bil2 = "Puluh "
						ElseIf (z = 3 Or z = 6 Or z = 9 Or z = 12 Or z = 15) Then
							bil2 = "Ratus "
						Else
							bil2 = ""
						End If
					Else
						bil2 = ""
					End If
					If (Y > 0) Then
						Select Case z
						Case 4
							bil2 = bil2 + "Ribu "
							Y = 0
						Case 7
							bil2 = bil2 + "Juta "
							Y = 0
						Case 10
							bil2 = bil2 + "Milyar "
							Y = 0
						Case 13
							bil2 = bil2 + "Trilyun "
							Y = 0
						End Select
					End If
					urai = urai + Bil1 + bil2
				Wend
				   
				urai = urai + strpecahan
				Terbilang = urai & ""
				Terbilang = UCase(Terbilang)
				n1 = UCase(Left(Terbilang, 1))
				Terbilang = n1 & Mid(Terbilang, 2, Len(Terbilang) - 1)
				End Function
		
			
			%>
			<div class="title_terbilang">
			The Sum Of :
			<hr/>
			Terbilang
			</div>
			<div class="terbilang" >
			<%  txtSQL= "select upper(replace(terbilang_indo("& v_Amount &"),'-',' ')) terbilang from dual"
			    Set rsCats = Server.CreateObject("ADODB.Recordset")
				rsCats.Open txtSQL, connstr
				if Not rsCats.eof then
					v_terbilang =  rsCats("terbilang") 
				end if
				rsCats.close
				set rsCats = Nothing
			%>
			<%=v_terbilang%> RUPIAH
			</div>
			</td>
		</tr>
	</tbody>
</table>
<%
   set rsPurchaseOrders = server.CreateObject("Adodb.recordset")
   qprd = 	"   SELECT vouchercode," & _
			"          MAX (comments) comments," & _
			"          fcba," & _
			"          fcentry," & _
			"          currentauthorizelevel," & _
			"          MAX (lastupdate) lastupdate" & _
			"     FROM (  SELECT vouchercode," & _
			"                    '' comments," & _
			"                    fcba," & _
			"                    fcentry," & _
			"                    MAX (currentauthorizelevel) currentauthorizelevel," & _
			"                    MAX (lastupdate) lastupdate" & _
			"               FROM paymentvoucher_hist" & _
			"              WHERE vouchercode = '"& vouchercode &"' AND fcba = '"& fcba &"'" & _
			"           GROUP BY vouchercode, fcba, fcentry" & _
			"           UNION ALL" & _
			"             SELECT vouchercode," & _
			"                    comments," & _
			"                    fcba," & _
			"                    fcentry," & _
			"                    MAX (currentauthorizelevel) currentauthorizelevel," & _
			"                    MAX (lastupdate) lastupdate" & _
			"               FROM pv_authorization_comments" & _
			"              WHERE vouchercode = '"& vouchercode &"' AND fcba = '"& fcba &"'" & _
			"           GROUP BY vouchercode," & _
			"                    fcba," & _
			"                    fcentry," & _
			"                    comments)" & _
			" GROUP BY vouchercode," & _
			"          fcba," & _
			"          fcentry," & _
			"          currentauthorizelevel" & _
			" ORDER BY currentauthorizelevel"

  ' response.Write qprd	
   Set rsPurchaseOrders = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrders.Open qprd, connstr
   if not rsPurchaseOrders.eof then
      comments = rsPurchaseOrders("comments")
	  fcentry = rsPurchaseOrders("fcentry")
	  lastupdate = rsPurchaseOrders("lastupdate")
   End if
   rsPurchaseOrders.close
   set rsPurchaseOrders = Nothing %>
   <table border="0" width="100%" id="product-table" style="border-bottom:1px #000000 solid;">
		<thead>
			<tr>
				<th>Authorizer</th>
				<th>Action Authorize</th>
				<th>Comments</th>
			</tr>
		 </thead>
		 <tbody>
			

		 <%	Set rsPurchaseOrders = Server.CreateObject("ADODB.Recordset")
			rsPurchaseOrders.Open qprd, connstr
			if Not rsPurchaseOrders.eof then	
			nCountofField=rsPurchaseOrders.fields.count-1
		%> 
			  
			  <%
			  While Not rsPurchaseOrders.EOF
				%>
				  
			<tr>
				<td align="center" style="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; width:170px;"><%=rsPurchaseOrders("fcentry")%>&nbsp;</td>
				<td align="center" style="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; width:205px;"><%=rsPurchaseOrders("lastupdate")%>&nbsp;</td>
				<td align="left"  style="<%=cmdstyle%>"><%=rsPurchaseOrders("comments")%>&nbsp;</td>
			</tr>
			
			  <%
				  rsPurchaseOrders.MoveNext
				  intRecordsPerPage = intRecordsPerPage - 1
			  Wend  %> 
			  
		<%	  
		   end if 
		   rsPurchaseOrders.Close
		   set rsPurchaseOrders = Nothing%> 
			
		</tbody>
	</table>
<div id="footer_note" style="width:95%;">	
	<div class="block">Penerima,
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






