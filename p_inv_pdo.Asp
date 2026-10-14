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
<title>FBPlants - Print PDO Invoice</title>
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
Dim txtNoReceved,txtDateReceived,txtRemarks,q,nCountofField,no,i
Dim Remarks
DIM  txtnoinvoice,txttgl,txtsumber,amount
dim sql,rsa,bulanx,tahunx,pcompany,pfcba,tglawalinv,tglakhirinv
dim sql1,sql2,invby

bulanx = request("bulan")
tahunx = request("tahun")
pcompany  = request("pcompany")
pfcba  = request("pfcba")
tglawalinv  = request("tglawalinv")
tglakhirinv  = request("tglakhirinv")

if pcompany = "PT.ARJ" then 
	invby   = "ARJ"
elseif pcompany = "PT.GAS" then 
	invby   = "GAS"
elseif pcompany = "PT.KSR" then 
	invby   = "KSR"
elseif pcompany = "PT.BRS" then 
	invby   = "BRS"
elseif pcompany = "PT.KAA" then 
	invby   = "KAA"
elseif pcompany = "PT.ESR" then 
	invby   = "ESR"
elseif pcompany = "PT.BIA" then 
	invby   = "BIA"
elseif pcompany = "PT.BRP" then 
	invby   = "BRP"
elseif pcompany = "PT.CAE" then 
	invby   = "CAE"
end if

'hapus dulu invoice di bulan yg di proses
sql1   = "DELETE FROM OPR_INV_PDO WHERE TAHUN='" & tahunx & "' AND BULAN='" & bulanx & "' AND FCBA='" & pfcba & "' "
connStr.Execute(sql1)

arr1 = Split(tglawalinv, "/")
arr2 = Split(tglakhirinv, "/")

hasiltgl1 = arr1(1) & "/" & arr1(0) & "/" & arr1(2)
hasiltgl2 = arr2(1) & "/" & arr2(0) & "/" & arr2(2)

sqlGet = " SELECT INVOICECODE from CONTRACTORINVOICEHDR WHERE FCBA='"& pfcba &"' AND " & _ 
		 " TRANSDATE BETWEEN TO_DATE('"& hasiltgl1 &"','DD/MM/YYYY') AND TO_DATE('" & hasiltgl2 & "','DD/MM/YYYY')"
' response.write sqlGet
Set rsTest = connStr.Execute(sqlGet)
Dim hasil
hasil = ""
Do Until rsTest.EOF
	If hasil = "" Then
		hasil = "'" & rsTest("INVOICECODE") & "'"    ' record pertama tanpa koma
	Else
		hasil = hasil & "|| '|' ||'" & rsTest("INVOICECODE") & "'"  ' record berikutnya pakai koma
	End If
	rsTest.MoveNext
Loop

' --- Tambahkan kurung ---
If hasil <> "" Then
	hasil = hasil
Else
	hasil = "(NULL)"  ' kalau query tidak ada hasil
End If

' response.write sqlGet

'insert invoice pdo
sql2 = "INSERT INTO OPR_INV_PDO( NOINVOICE, CREATEDATE, SUMBER, AMOUNT, FCENTRY, " & _ 
"       FCEDIT, FCIP, FCBA, LASTUPDATE, INVOICEBY, " & _ 
"       TAHUN, BULAN, DESCRIPTION, BAPPNO) " & _ 
" Select FCBA||'/INV/PDO/'||'" & bulanx & "'||'/'||'" & tahunx & "'||'/'||1 NOINVOICE,SYSDATE CREATEDATE,1 SUMBER, SUM(TOTAL) AMOUNT,'" & Session("CodeName") & "' FCENTRY, " & _ 
"       '" & Session("CodeName") & "' FCEDIT,0 FCIP,'" & pfcba & "' FCBA,SYSDATE LASTUPDATE,'" & invby & "' INVOICEBY, " & _ 
"       '" & tahunx & "' TAHUN,'" & bulanx & "' BULAN,'PDO UPAH PERIODE: ' || '" & bulanx & "' ||'/'||'" & tahunx & "' DESCRIPTION, "& hasil &" " & _ 
" from   VIEW_BKMPAYROLL " & _ 
" WHERE  FCBA='" & pfcba & "' AND month='" & bulanx & "' AND year='" & tahunx & "' " & _ 
" GROUP BY SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "',0,FCBA,SYSDATE,'" & invby & "' " & _ 
"    UNION ALL " & _ 
" Select FCBA||'/INV/PDO/'||'" & bulanx & "'||'/'||'" & tahunx & "'||'/'||2 NOINVOICE,SYSDATE CREATEDATE,2 SUMBER, SUM(GRANDTOTAL_AFTERTAX) AMOUNT,'" & Session("CodeName") & "' FCENTRY,  " & _ 
"       '" & Session("CodeName") & "' FCEDIT,0 FCIP,'" & pfcba & "' FCBA,SYSDATE LASTUPDATE,'" & invby & "' INVOICEBY,  " & _ 
"       '" & tahunx & "' TAHUN,'" & bulanx & "' BULAN,'PDO CONTRACTOR LOKAL PERIODE: ' || '" & bulanx & "' ||'/'||'" & tahunx & "' DESCRIPTION, "& hasil &" " & _
" From   CONTRACTORINVOICEHDR " & _ 
" WHERE  FCBA='" & pfcba & "' AND TRANSDATE BETWEEN TO_DATE('" & tglawalinv & "','MM/DD/YYYY') AND TO_DATE('" & tglakhirinv & "','MM/DD/YYYY') " & _ 
" GROUP BY SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "',0,FCBA,SYSDATE,'" & invby & "'"
' response.write sql2
connStr.Execute(sql2)


' ini qry untuk select company
q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & request("pfcba") & "' "
set rsCat = server.CreateObject("Adodb.recordset")
set rsCat = connStr.Execute(q)
if not rsCat.eof then
	fcbaname = rsCat("fcName")
	CompanyName=rsCat("FCCOMPANYNAME")
	CompanyCode=rsCat("fccompanycode")
	MgrName = rsCat("FCMANAGER")
	KTUName = rsCat("FCKTU")
	CropName=rsCat("FCCROP")
	FcbaType=rsCat("fcType")
end if
rsCat.close
set rsCat = Nothing
'Response.Write fcbaname
'Response.end


sql = "SELECT noinvoice,TO_char(createdate,'DD/MM/YYYY')tglinv,sumber,description,(NVL(amount,0)) amount " & _
		" FROM OPR_INV_PDO " & _
		" WHERE tahun = '" & tahunx & "'" & _ 
		" AND bulan = '" & bulanx & "'" & _ 
		" AND fcba = '" & pfcba & "' " 

		Set rsa = Server.CreateObject("ADODB.Recordset")
		Set rsa = connStr.Execute(sql)
if not rsa.eof then 
	txtnoinvoice = rsa("noinvoice")
	txttgl = rsa("tglinv")
end if
 %>

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
				<div class="judul"><strong>INVOICE PDO</strong></div>
					<table width="100%" >
					<!-- Left -->
					<td valign="top">
					<div class="point">
							<div class="name">
								PT
							</div>
							<div class="isi">
							: <%=CompanyName%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point">
							<div class="name">
							Kebun/Pabrik
							</div>
							<div class="isi">
							: <%=request("fcba")%> - <%=fcbaname%>
							</div>
						</div>
						<div class="clear"></div>
						
						<div class="clear"></div>
					<!-- right -->
					<td width="400px" valign="top">			
						<div class="point_right">
							<div class="name">
							Nomor 
							</div>
							<div class="isi">
							: <%=txtnoinvoice%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point_right">
							<div class="name">
							Tanggal 
							</div>
							<div class="isi">
							:  <%=txttgl%>
							</div>
						</div>
						<div class="clear"></div>
						
						<div class="point_right">
							<div class="name">
							Periode 
							</div>
							<div class="isi">
							:  <%=bulanx& " - " & tahunx%>
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

					<th style="font-size:10px">Keterangan</th>
					<th style="font-size:10px">Quantity</th>
					<th style="font-size:10px">Amount</th>					
					</tr>
			 </thead>
			 <tbody>
	

   <%
		
	  no=1
	  nQty=0
	  nAmount=0
	  While Not rsa.EOF

		 
		  %>
		   <tr>
					  <td align="center" style="<%=cmdstyleLeft%>"><%=no%>&nbsp;</td>
					  <td align="left"  style="width:200px;<%=cmdstyle%>" title="" valign="middle"><font face="Verdana"><%=rsa("DESCRIPTION")%></td>
					  <td align="center"  style="width:50px;<%=cmdstyle%> " title="" valign="middle">1 &nbsp;</td>
					  <td align="Right"  style="width:50px;<%=cmdstyle%> " title="" valign="middle"><%=FormatNumber(rsa("Amount"),0)%>&nbsp;</td>					  
		  </tr>
		  <%
		  nAmount = cdbl(nAmount) + cdbl(rsa("Amount"))
		  rsa.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  
	
  <%

   rsa.Close
   set rsa = Nothing
   
   
   %>
   	</tbody>
	</table>
		
   	
	 <table border="1" width="100%" id="product-table" >
	 <thead>

	  <tr>
	  <td colspan="2"  align="right"><font face="Verdana">TOTAL</td>
	  <td style="width:140px;<%=cmdstyle%>" align="right"><font face="Verdana"><%=FormatNumber(nAmount,0)%>&nbsp;</td>
	  </tr>				    
   	</tbody>
	</table>
	
	<br/>

</body>
</html>
 