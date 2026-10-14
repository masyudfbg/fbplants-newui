<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
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
            new xTableHeaderFixed('xthf-green', window);
            new xTableHeaderFixed('xthf-blue', window);
          }, false
        );
        </script>
</head>

<body>
<style>
.note {margin:auto;padding:5px;}
#info_note .judul { text-align:center;}
#info_note .left { float: left; width:35%;}
#info_note .right {  float: right; width:35%;}
#info_note .point { font-size:12px; font-weight:bold; width:100%;  margin-top:5px;}
#info_note .point span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
#info_note .point .name { width:150px; float:left;}
#info_note .point .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}

#info_note .point_right { font-size:12px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
#info_note .point_right .name { width:100px; float:left;}
#info_note .point_right .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}
.clear { clear:both; content:'';}
table .long {width:400px;}
table .small {width:40px;}
table th span { font-style:italic; font-weight:normal;}
.line { height:1px; background:#FFFFFF; content:' ';}
.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
#product-table td.total { background-color:#F4F4A0; border-bottom:double; font-weight: bold; font-style: italic;}
#product-table td {border: 1px solid #2d2d2d;}
#product-table td.none { border: none;}
#footer_note .block { font-size:10px; float:left; width:19%; height:200px; background-color:#fff; margin-left:1%; text-align:center;}
input.isian[type="text"] {
    border: 0 none;
    box-shadow: none;
    color: #000;
    font-size: 10px;
    height: 20px;
    padding: 0;
    text-align: center;
    width: 100%;
}
hr {
    margin-top: 5px;
}
hr {
    border-color: #000000 -moz-use-text-color -moz-use-text-color;
    border-style: solid none none;
    border-width: 1px 0 0;
}
</style>
<%
Dim oRS, oSQL, NoUrut, nTotalValue, grnno

grnno = Request("grnno")

oSQL	=	" SELECT" & _
			"     FCBA" & _
			"     , SUPPLIERCODE" & _
			"     , SUPPLIER_API.GET_SUPP_NAME(SUPPLIERCODE, FCBA) SUPPLIERNAME" & _
			"     , POCODE" & _
			"     , RECEIVENOTECODE" & _
			"     , TO_NUMBER(TO_CHAR(RNDATE,'DD')) RECEIVENOTE_DD" & _
			"     , TO_NUMBER(TO_CHAR(RNDATE,'MM')) RECEIVENOTE_MM" & _
			"     , TO_NUMBER(TO_CHAR(RNDATE,'YYY')) RECEIVENOTE_YYYY" & _
			" FROM RECEIVENOTE" & _
			" WHERE FCBA = '" & Session("FCBA") & "'" & _
			"       AND RECEIVENOTECODE = '" & grnno & "'"

Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open oSQL, connstr, adLockOptimistic, adOpenStatic

%>
<div class="note" >
<%If oRS.BOF And oRS.EOF Then%>
	<header id="info_note" >
		<div class="judul"><strong>Goods Received Note</strong></div>
			<table width="100%">
			<!-- Left -->
			<td valign="top">
				<div class="point">
					<div class="name">
					Nama Supplier<span>Supplier's Name</span>
					</div>
					<div class="isi">
					: 
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Kode Kreditor<span>Creditor's Code</span>
					</div>
					<div class="isi">
					:
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Nomor DO Supplier <span>Supplier's DO Number</span>
					</div>
					<div class="isi">
					: 
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Nomor STO Kita<span>Our PO Number</span>
					</div>
					<div class="isi">
					: 
					</div>
				</div>
			</td> 
		 
			<!-- right -->
			<td width="400px" valign="top">			
				<div class="point_right">
					<div class="name">
					Nomor<span>Number</span>
					</div>
					<div class="isi">
					: 
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Tanggal<span>Date</span>
					</div>
					<div class="isi">
					: 
					</div>
				</div>
				<div class="clear"></div>
			</td>
		</table>
		

	</header>
<%Else%>
	<header id="info_note" >
		<div class="judul"><strong>Goods Received Note</strong></div>
			<table width="100%">
			<!-- Left -->
			<td valign="top">
				<div class="point">
					<div class="name">
					Nama Supplier<span>Supplier's Name</span>
					</div>
					<div class="isi">
					: <%=oRS("SUPPLIERNAME")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Kode Kreditor<span>Creditor's Code</span>
					</div>
					<div class="isi">
					:
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Nomor DO Supplier <span>Supplier's DO Number</span>
					</div>
					<div class="isi">
					: 
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Nomor STO Kita<span>Our PO Number</span>
					</div>
					<div class="isi">
					: <%=oRS("POCODE")%>
					</div>
				</div>
			</td> 
		 
			<!-- right -->
			<td width="400px" valign="top">			
				<div class="point_right">
					<div class="name">
					Nomor<span>Number</span>
					</div>
					<div class="isi">
					: <%=oRS("RECEIVENOTECODE")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Tanggal<span>Date</span>
					</div>
					<div class="isi">
					: <%=oRS("RECEIVENOTE_DD")%> - <%=Bulan(oRS("RECEIVENOTE_MM"))%> - <%=oRS("RECEIVENOTE_YYYY")%>
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
<br/><br/>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
    <tr>
        <th class="small">Nomor<div class="line"></div><span>Item</span></th>
		<th>STOCK CODE</th>
        <th class="long">DESCRIPTION</th>
        <th style="font-size:10px">No. kartu Stock<div class="line"></div><span>Stock Card No</span></th>
        <th>Terima<div class="line"></div><span>QTY</span></th>
        <th>Satuan<div class="line"></div><span>UOM</span></th>
        <th>Harga Satuan<div class="line"></div><span>Unit Cost</span></th>
        <th>Total Harga diluar PPN<div class="line"></div><span>Total Exct.Vat</span></th>
        <th>PPN<div class="line"></div><span>Vat Value</span></th>
        <th>Total Nilai<div class="line"></div><span>Total Value</span></th>
        <th>No.Perk<div class="line"></div><span>ACC Code</span></th>
        <th>Catatan<div class="line"></div><span>Remark</span></th>
        <th>No. PR</th>
        <th>tgl. PR</th>
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
			"     , SUM(NVL(RND.AMOUNTRECEIVED, 0))/SUM(NVL(RND.QUANTITY, 0)) UNITCOST" & _
			"     , NVL(SUM(NVL(RND.AMOUNTRECEIVED, 0)), 0) TOTALVALUE" & _
			"     , RND.JOBCODE" & _
			"     , JOB_API.GET_JOBCODE_NAME(RND.JOBCODE, RND.FCBA) JOBNAME" & _
			" FROM RECEIVENOTEDETAIL RND" & _
			" INNER JOIN RECEIVENOTE RN ON (RN.FCBA = RND.FCBA AND RN.RECEIVENOTECODE = RND.RECEIVENOTECODE)" & _
			" WHERE RND.FCBA = '" & Session("FCBA") & "'" & _
			"       AND RND.RECEIVENOTECODE = '" & grnno & "'" & _
			" GROUP BY" & _
			"     RND.FCBA" & _
			"     , RN.POCODE" & _
			"     , RND.PURCHASELINENOKEY" & _
			"     , RND.PURCHASEITEMCODE" & _
			"     , RND.JOBCODE"

Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open oSQL, connstr, adLockOptimistic, adOpenStatic
%>
<%If oRS.BOF And oRS.EOF Then%>
    <tr>
        <td Colspan="14">NO DATA AVAILABLE</td>
    </tr>
<%Else
	While Not oRS.EOF
		NoUrut	=	NoUrut + 1%>
		<tr>
			<td align="Right"><%=NoUrut%></td>
			<td align="Left"><%=oRS("STOCKCODE")%></td>
			<td align="Left"><%=oRS("STOCKNAME")%></td>
			<td></td>
			<td align="Right"><%=FormatNumber(oRS("QTY"), 2)%></td>
			<td align="Left"><%=oRS("UOM")%></td>
			<td align="Right"><%=FormatNumber(oRS("UNITCOST"), 0)%></td>
			<td></td>
			<td></td>
			<td align="Right"><%=FormatNumber(oRS("TOTALVALUE"), 0)%></td>
			<td align="Left"><%=oRS("JOBCODE")%></td>
			<td></td>
			<td></td>
			<td></td>
		</tr>
	<%	nTotalValue	=	nTotalValue + Cdbl(oRS("TOTALVALUE"))
		oRS.MoveNext
	Wend%>
    <tr>
        <td colspan="7" class="none"></td>
        <td></td>
        <td> Total</td>
        <td align="Right"><%=FormatNumber(nTotalValue, 0)%></td>
        <td colspan="4" class="none"></td>
        
    </tr>
     <tr>
        <td colspan="7" class="none"></td>
        <td class="total">Total IDR</td>
        <td class="total"> </td>
        <td class="total" align="Right"><%=FormatNumber(nTotalValue, 0)%></td>
        <td colspan="4" class="none"></td>
        
    </tr>
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
    	<div class="block" style="margin:0px;">
        	<table style="text-align:left; width:100%;">
            	<tr><td>Asli</td><td>- Accounting</td></tr>
                <tr><td>Copy 1</td><td>- Accounting</td></tr>
                <tr><td>Copy 2</td><td>- Pembelian</td></tr>
                <tr><td>Copy 3</td><td>- Pemesan</td></tr>
            </table>
        
        </div>
        <div class="block">Prepared
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <HR>
        <input class="isian" type="text" value="">
        </div>
        <div class="block">Passed
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <HR>
        <input class="isian" type="text" value="">
        </div>
        <div class="block">Approved
        <br/><br/><br/><br/>
        <input class="isian" type="text" value="">
        <HR>
        <input class="isian" type="text" value="">
        </div>
        <div class="block"></div>
    </div>
    
    
    <div class="clear"></div>
</div>




</body>
</html>
