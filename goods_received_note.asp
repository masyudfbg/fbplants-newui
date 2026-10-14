<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<%response.buffer=True%>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Goods Received Note</title>
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
#product-table th {font-size:10px; }
#product-table td {font-size:10px; padding:2px;padding-bottom: 0;
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
Dim oRS, oSQL, NoUrut, nTotalValue, grnno,cmdstyle,cmdstyleStock,cmdstyleTop,cmdstyleLeft

grnno = Request("grnno")

cmdstyle="font-size:10px ;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleStock="font-size:10px ;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="font-size:10px ;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="font-size:10px; border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
oSQL	=	" SELECT" & _
			"     FCBA" & _
			"     , SUPPLIERCODE" & _
			"     , SUPPLIER_API.GET_SUPP_NAME(SUPPLIERCODE, FCBA) SUPPLIERNAME" & _
			"     , STORE_API.get_fcname(STORECODE,FCBA) STORENAME "& _
			"     , POCODE" & _
			"     , STOCKORDERNO" & _
			"     , DELIVERYORDERCODE" & _
			"     , RECEIVENOTECODE" & _
			"     , TO_NUMBER(TO_CHAR(RNDATE,'DD')) RECEIVENOTE_DD" & _
			"     , TO_NUMBER(TO_CHAR(RNDATE,'MM')) RECEIVENOTE_MM" & _
			"     , TO_NUMBER(TO_CHAR(RNDATE,'YYYY')) RECEIVENOTE_YYYY" & _
			" FROM RECEIVENOTE" & _
			" WHERE FCBA = '" & Session("FCBA") & "'" & _
			"       AND RECEIVENOTECODE = '" & grnno & "'"

Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open oSQL, connstr, adLockOptimistic, adOpenStatic

	' Added By Sendy untuk validasi menu 11/17/2018
	Dim Menu : Menu = "N"
	Dim oRecordSetMenu, oSQLQueryMenu
	oSQLQueryMenu = " SELECT 'Y' istrue FROM users_tab " & _
				" WHERE fccode = '" & Session("CodeName") & "' AND MENU_GRN = 'TRUE'" 
	' response.write oSQLQueryMenu
	set oRecordSetMenu = server.CreateObject("Adodb.recordset")
	oRecordSetMenu.Open oSQLQueryMenu, connstr
	If Not oRecordSetMenu.Eof then
		Menu = oRecordSetMenu("istrue")
	Else
		Menu = "N"
	End if
	' End

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
		table.PAYSLIP td{padding:5px;}
		.page-break hr {display:block;border: 1px dashed red;}
		.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
		#note {margin-left:auto;padding:5px;}
		#info_note .judul { text-align:center; font-size:14px;}
		#info_note .left { float: left; width:35%;}
		#info_note .right {  float: right; width:35%;}
		#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:0px;}
		#info_note .point span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:0px;}
		#info_note .point .name { width:100px; float:left; margin-top: 0px;}
		#info_note .point .isi  { float:left; height:5px;  margin-left:10px;}

		#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:0px; text-align:left;}
		#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:0px;}
		#info_note .point_right .name { width:100px; float:left; margin-top: 0px;}
		#info_note .point_right .isi  { float:left; height:5px;  margin-left:10px;}
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
		@media print
		{
		td{padding:3px;}
		table.PAYSLIP td{padding:3px;}
		.page-break hr {border:none;display: none;}
		.PAYSLIP td {font-family:verdana,sans-serif; font-size:7px; padding:0px;}
		#note {margin-left:auto;padding:0px;}
		#info_note .judul { text-align:center; font-size:14px;}
		#info_note .left { float: left; width:35%;}
		#info_note .right {  float: right; width:35%;}
		#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:0px;}
		#info_note .point span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:0px;}
		#info_note .point .name { width:100px; float:left; margin-top: 0px;}
		#info_note .point .isi  { float:left; height:5px;  margin-left:10px;}

		#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:0px; text-align:left;}
		#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:0px;}
		#info_note .point_right .name { width:100px; float:left; margin-top: 0px;}
		#info_note .point_right .isi  { float:left; height:5px;  margin-left:10px;}
		.clear { clear:both; content:'';}
		table .long {width:400px;}
		table .small {width:40px;}
		table th span { font-style:italic; font-weight:normal;}
		.line { height:1px; background:#000; content:' ';margin: auto;}
		.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
		#product-table th { background-color:#fff; color:#000; font-size:10;}
		#product-table td { border-color:#000; color:#000;padding-bottom: 0;order-top:0px;border-bottom:0px;
    padding-top: 0;}
		#product-table td.none { border: none;}
		#footer_note .block { font-size:10px; float:left; width:19%; height:150px; background-color:#fff; margin-left:1%; text-align:center;}
		.tembusan{font-size:12px;}
		input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
		hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
		hr{ margin-top: 5px;}
		}
		</style>
		

	<header id="info_note" >
		<div class="judul"><strong>GOOD RECEIVE NOTE</strong></div>
			<table width="100%">
			<!-- Left -->
			<%IF NOT oRS.EOF THEN%>
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
							: <%=Session("FCBA")%> - <%=Session("BA")%>
							</div>
					</div>
				<div class="clear"></div>
					<div class="point">
							<div class="name">
							Gudang 
							</div>
							<div class="isi">
							: <%=oRS("STORENAME")%> 
							</div>
					</div>
				<div class="clear"></div>
					<div class="point">
						<div class="name">
						Nama Supplier
						</div>
						<div class="isi">
						: <%=oRS("SUPPLIERNAME")%>
						</div>
					</div>
			</td> 
			<!-- right -->
			<td width=40% valign="top">			
			    <div class="point_right">
					<div class="name">
					Tanggal
					</div>
					<div class="isi">
					: <%=oRS("RECEIVENOTE_DD")%>-<%=Bulan(oRS("RECEIVENOTE_MM"))%>-<%=oRS("RECEIVENOTE_YYYY")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Nomor GRN
					</div>
					<div class="isi">
					: <%=oRS("RECEIVENOTECODE")%> 
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Nomor PO
					</div>
					<div class="isi">
					: <%=oRS("POCODE")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Nomor STO 
					</div>
					<div class="isi">
					: <%=oRS("STOCKORDERNO")%> 
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Nomor DO 
					</div>
					<div class="isi">
					: <%=oRS("DELIVERYORDERCODE")%> 
					</div>
				</div>
				<div class="clear"></div>
			</td> 
		</table>
		

	</header>

<%End If%>
<%
oSQL	=	""
oRS.Close
set oRS	=	Nothing
%>

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
				font-size:10px;
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
        <th class="small" style="font-size:10px" style="border-bottom:1px #000000 solid;">No</th>
		<th style="font-size:10px">Stock Code</th>
		<th class="long" style="font-size:10px">Description</th>
        <th style="font-size:10px">Terima<div class="line"></div><span>QTY</span></th>
        <th style="font-size:10px">Satuan<div class="line"></div><span>UOM</span></th>
        <th style="font-size:10px">No.PR</th>
        <th style="font-size:10px">tgl.PR</th>
    </tr>
</thead >

<tbody>
<%

NoUrut	=	0
oSQL	=	" SELECT" & _
			"     RND.FCBA" & _
			"     , RN.POCODE" & _
			"     , RND.PURCHASELINENOKEY POLINEKEY" & _
			"     , RND.PURCHASEITEMCODE STOCKCODE" & _
			"     , STOCK_API.GET_STOCKNAME(RND.PURCHASEITEMCODE, RND.FCBA) STOCKNAME" & _
			"     , NVL(SUM(NVL(RND.QUANTITY, 0)), 0) QTY" & _
			"     , STOCK_API.GET_UOM(RND.PURCHASEITEMCODE, RND.FCBA) UOM" & _
			" 	  , SUM(NVL(RND.AMOUNTRECEIVED, 0))/NULLIF(SUM(NVL(RND.QUANTITY, 0)), 0) UNITCOST" & _
			"     , NVL(SUM(NVL(RND.AMOUNTRECEIVED, 0)), 0) TOTALVALUE" & _
			"     , RND.JOBCODE" & _
			"     , JOB_API.GET_JOBCODE_NAME(RND.JOBCODE, RND.FCBA) JOBNAME" & _
			"     , POD.PRCODE " & _
            "     , TO_CHAR(PR.PRDATE,'DD-MM-YYYY') PRDATE " & _
			" FROM RECEIVENOTEDETAIL RND" & _
			" INNER JOIN RECEIVENOTE RN ON (RN.FCBA = RND.FCBA AND RN.RECEIVENOTECODE = RND.RECEIVENOTECODE)" & _
			" LEFT JOIN PURCHASEORDERDETAILS POD ON (RN.FCBA=POD.FCBA AND RN.POCODE=POD.POCODE AND RND.PURCHASEITEMCODE=POD.ITEMCODE AND RND.PURCHASELINENOKEY =POD.LINENOKEY)" & _
			" LEFT JOIN PURCHASEREQUEST PR ON ( POD.FCBA=PR.FCBA AND POD.PRCODE=PR.PRCODE)" & _
			" WHERE RND.FCBA = '" & Session("FCBA") & "'" & _
			"       AND RND.RECEIVENOTECODE = '" & grnno & "'" & _
			
			" GROUP BY" & _
			"     RND.FCBA" & _
			"     , RN.POCODE" & _
			"     , RND.PURCHASELINENOKEY" & _
			"     , RND.PURCHASEITEMCODE" & _
			"     , RND.JOBCODE" & _
			"     , POD.PRCODE" & _ 
            "     , PR.PRDATE" 

			' "       AND POD.PRCODE IS NOT NULL " & _ 

Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open oSQL, connstr, adLockOptimistic, adOpenStatic
'response.write oSQL
%>
<%If oRS.BOF And oRS.EOF Then%>
    <tr>
        <td Colspan="14">NO DATA AVAILABLE</td>
    </tr>
<%Else
	While Not oRS.EOF
		NoUrut	=	NoUrut + 1%>
		<tr>
			<td align="center" style="<%=cmdstyleLeft%>"><%=NoUrut%></td>
			<td align="center"  style="<%=cmdstyleStock%>" title="ITEMCODE" valign="middle"><%=oRS("STOCKCODE")%></td>
			<td align="left"  style="<%=cmdstyle%>" title="stockname" valign="middle"><%=oRS("STOCKNAME")%></td>
			<td align="CENTER"  style="width:50px;<%=cmdstyle%> " title="approvedqty" valign="middle"><%=FormatNumber(oRS("QTY"), 2)%></td>
			<td align="center"  style="width:50px;<%=cmdstyle%>" title="unitofmeasurement" valign="middle"><%=oRS("UOM")%></td>
			<td align="center"  style="width:50px;<%=cmdstyle%>" title="pr" valign="middle"><%=oRS("PRCODE")%></td>
			<td align="center"  style="width:80px;<%=cmdstyle%>" title="prdate" valign="middle"><%=oRS("PRDATE")%></td>
		</tr>
	<%	oRS.MoveNext
	Wend%>
	
<%End If%>
<%
oSQL	=	""
oRS.Close
set oRS	=	Nothing
%>
</tbody>
</table>
<br/>
        <div id="footer_note">
    	<div class="block">Dibuat Oleh,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Krani
        </div>
		 <div class="block">Dikirim Oleh,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Transporter
        </div>
        <div class="block">Diterima Oleh,
        <br/><br/><br/><br/>
        <input class="isian" type="text" value="<%=Session("FCMANAGER")%>" readonly='True'>
        <hr/>
        Ka.Gudang
        </div>
		<div class="block">Diperiksa Oleh,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        KTU
        </div>
		<div class="block">Diketahui Oleh,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Manager/SAM
        </div>
        <div class="block"></div>
    </div>
    
    <div class="clear"></div>
</div>




</body>
</html>
