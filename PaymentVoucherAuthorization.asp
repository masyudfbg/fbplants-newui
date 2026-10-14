<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Payment Voucher Approval</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<link rel="stylesheet" href="css/font/font-awesome.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

function harusangka(evt) {
    var e = evt || window.event;
    var key = e.keyCode || e.which;

    if (!e.shiftKey && !e.altKey && !e.ctrlKey &&
    // numbers   
    key >= 48 && key <= 57 ||
	e.ctrlKey && key == 97 ||
    // Backspace and Tab and Enter
    key == 8 || key == 9 || key == 13 ||
    // left and right arrows
    key == 37 || key == 39 ||
    // Del and Ins
	key == 46 ) {
        // input is VALID
		}
	else {
			// input is INVALID
			e.returnValue = false;
			if (e.preventDefault) e.preventDefault();
		}
}

</script>
<style>
.button {
    position: relative;
    overflow: visible;
    display: inline-block;
    padding: 0.5em 1em;
    border: 1px solid #d4d4d4;
    margin: 0;
    text-decoration: none;
    text-align: center;
    text-shadow: 1px 1px 0 #fff;
    font:11px/normal sans-serif;
    color: #333;
    white-space: nowrap;
    cursor: pointer;
    outline: none;
    background-color: #ececec;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#f4f4f4), to(#ececec));
    background-image: -moz-linear-gradient(#f4f4f4, #ececec);
    background-image: -ms-linear-gradient(#f4f4f4, #ececec);
    background-image: -o-linear-gradient(#f4f4f4, #ececec);
    background-image: linear-gradient(#f4f4f4, #ececec);
    -moz-background-clip: padding; /* for Firefox 3.6 */
    background-clip: padding-box;
    border-radius: 0.2em;
    /* IE hacks */
    zoom: 1;
    *display: inline;
}

.button:hover,
.button:focus,
.button:active,
.button.active {
    border-color: #3072b3;
    border-bottom-color: #2a65a0;
    text-decoration: none;
    text-shadow: -1px -1px 0 rgba(0,0,0,0.3);
    color: #fff;
    background-color: #3c8dde;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#599bdc), to(#3072b3));
    background-image: -moz-linear-gradient(#599bdc, #3072b3);
    background-image: -o-linear-gradient(#599bdc, #3072b3);
    background-image: linear-gradient(#599bdc, #3072b3);
}

.button:active,
.button.active {
    border-color: #2a65a0;
    border-bottom-color: #3884cd;
    background-color: #3072b3;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#3072b3), to(#599bdc));
    background-image: -moz-linear-gradient(#3072b3, #599bdc);
    background-image: -ms-linear-gradient(#3072b3, #599bdc);
    background-image: -o-linear-gradient(#3072b3, #599bdc);
    background-image: linear-gradient(#3072b3, #599bdc);
}

/* overrides extra padding on button elements in Firefox */
.button::-moz-focus-inner {
    padding: 0;
    border: 0;
}


/* =============================================================================
   Button icons
   ========================================================================== */

.button.icon:before {
    content: "";
    position: relative;
    top: 1px;
    float:left;
    width: 12px;
    height: 12px;
    margin: 0 0.75em 0 -0.25em;
    background: url(gh-icons.png) 0 99px no-repeat;
}

.button.arrowup.icon:before { background-position: 0 0; }
.button.arrowup.icon:hover:before,
.button.arrowup.icon:focus:before,
.button.arrowup.icon:active:before { background-position: -12px 0; }

.button.arrowdown.icon:before { background-position: 0 -12px; }
.button.arrowdown.icon:hover:before,
.button.arrowdown.icon:focus:before,
.button.arrowdown.icon:active:before { background-position: -12px -12px; }

.button.arrowleft.icon:before { background-position: 0 -24px; }
.button.arrowleft.icon:hover:before,
.button.arrowleft.icon:focus:before,
.button.arrowleft.icon:active:before { background-position: -12px -24px; }

.button.arrowright.icon:before { float:right; margin: 0 -0.25em 0 0.5em; background-position: 0 -36px; }
.button.arrowright.icon:hover:before,
.button.arrowright.icon:focus:before,
.button.arrowright.icon:active:before { background-position: -12px -36px; }

.button.approve.icon:before { background-position: 0 -48px; }
.button.approve.icon:hover:before,
.button.approve.icon:focus:before,
.button.approve.icon:active:before { background-position: -12px -48px; }

.button.add.icon:before { background-position: 0 -288px; }
.button.add.icon:hover:before,
.button.add.icon:focus:before,
.button.add.icon:active:before { background-position: -12px -288px; }

.button.remove.icon:before { background-position: 0 -60px; }
.button.remove.icon:hover:before,
.button.remove.icon:focus:before,
.button.remove.icon:active:before { background-position: -12px -60px; }

.button.log.icon:before { background-position: 0 -72px; }
.button.log.icon:hover:before,
.button.log.icon:focus:before,
.button.log.icon:active:before { background-position: -12px -72px; }

.button.calendar.icon:before { background-position: 0 -84px; }
.button.calendar.icon:hover:before,
.button.calendar.icon:focus:before,
.button.calendar.icon:active:before { background-position: -12px -84px; }

.button.chat.icon:before { background-position: 0 -96px; }
.button.chat.icon:hover:before,
.button.chat.icon:focus:before,
.button.chat.icon:active:before { background-position: -12px -96px; }

.button.clock.icon:before { background-position: 0 -108px; }
.button.clock.icon:hover:before,
.button.clock.icon:focus:before,
.button.clock.icon:active:before { background-position: -12px -108px; }

.button.settings.icon:before { background-position: 0 -120px; }
.button.settings.icon:hover:before,
.button.settings.icon:focus:before,
.button.settings.icon:active:before { background-position: -12px -120px; }

.button.comment.icon:before { background-position: 0 -132px; }
.button.comment.icon:hover:before,
.button.comment.icon:focus:before,
.button.comment.icon:active:before { background-position: -12px -132px; }

.button.fork.icon:before { background-position: 0 -144px; }
.button.fork.icon:hover:before,
.button.fork.icon:focus:before,
.button.fork.icon:active:before { background-position: -12px -144px; }

.button.like.icon:before { background-position: 0 -156px; }
.button.like.icon:hover:before,
.button.like.icon:focus:before,
.button.like.icon:active:before { background-position: -12px -156px; }

.button.favorite.icon:before { background-position: 0 -348px; }
.button.favorite.icon:hover:before,
.button.favorite.icon:focus:before,
.button.favorite.icon:active:before { background-position: -12px -348px; }

.button.home.icon:before { background-position: 0 -168px; }
.button.home.icon:hover:before,
.button.home.icon:focus:before,
.button.home.icon:active:before { background-position: -12px -168px; }

.button.key.icon:before { background-position: 0 -180px; }
.button.key.icon:hover:before,
.button.key.icon:focus:before,
.button.key.icon:active:before { background-position: -12px -180px; }

.button.lock.icon:before { background-position: 0 -192px; }
.button.lock.icon:hover:before,
.button.lock.icon:focus:before,
.button.lock.icon:active:before { background-position: -12px -192px; }

.button.unlock.icon:before { background-position: 0 -204px; }
.button.unlock.icon:hover:before,
.button.unlock.icon:focus:before,
.button.unlock.icon:active:before { background-position: -12px -204px; }

.button.loop.icon:before { background-position: 0 -216px; }
.button.loop.icon:hover:before,
.button.loop.icon:focus:before,
.button.loop.icon:active:before { background-position: -12px -216px; }

.button.search.icon:before { background-position: 0 -228px; }
.button.search.icon:hover:before,
.button.search.icon:focus:before,
.button.search.icon:active:before { background-position: -12px -228px; }

.button.mail.icon:before { background-position: 0 -240px; }
.button.mail.icon:hover:before,
.button.mail.icon:focus:before,
.button.mail.icon:active:before { background-position: -12px -240px; }

.button.move.icon:before { background-position: 0 -252px; }
.button.move.icon:hover:before,
.button.move.icon:focus:before,
.button.move.icon:active:before { background-position: -12px -252px; }

.button.edit.icon:before { background-position: 0 -264px; }
.button.edit.icon:hover:before,
.button.edit.icon:focus:before,
.button.edit.icon:active:before { background-position: -12px -264px; }

.button.pin.icon:before { background-position: 0 -276px; }
.button.pin.icon:hover:before,
.button.pin.icon:focus:before,
.button.pin.icon:active:before { background-position: -12px -276px; }

.button.reload.icon:before { background-position: 0 -300px; }
.button.reload.icon:hover:before,
.button.reload.icon:focus:before,
.button.reload.icon:active:before { background-position: -12px -300px; }

.button.rss.icon:before { background-position: 0 -312px; }
.button.rss.icon:hover:before,
.button.rss.icon:focus:before,
.button.rss.icon:active:before { background-position: -12px -312px; }

.button.tag.icon:before { background-position: 0 -324px; }
.button.tag.icon:hover:before,
.button.tag.icon:focus:before,
.button.tag.icon:active:before { background-position: -12px -324px; }

.button.trash.icon:before { background-position: 0 -336px; }
.button.trash.icon:hover:before,
.button.trash.icon:focus:before,
.button.trash.icon:active:before { background-position: -12px -336px; }

.button.user.icon:before { background-position: 0 -360px; }
.button.user.icon:hover:before,
.button.user.icon:focus:before,
.button.user.icon:active:before { background-position: -12px -360px; }


/* =============================================================================
   Button extensions
   ========================================================================== */

/* Primary button
   ========================================================================== */

.button.primary {
    font-weight: bold;
}

/* Danger button
   ========================================================================== */

.button.danger {
    color: #900;
}

.button.danger:hover,
.button.danger:focus,
.button.danger:active {
    border-color: #b53f3a;
    border-bottom-color: #a0302a;
    color: #fff;
    background-color: #dc5f59;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#dc5f59), to(#b33630));
    background-image: -moz-linear-gradient(#dc5f59, #b33630);
    background-image: -ms-linear-gradient(#dc5f59, #b33630);
    background-image: -o-linear-gradient(#dc5f59, #b33630);
    background-image: linear-gradient(#dc5f59, #b33630);
}

.button.danger:active,
.button.danger.active {
    border-color: #a0302a;
    border-bottom-color: #bf4843;
    background-color: #b33630;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#b33630), to(#dc5f59));
    background-image: -moz-linear-gradient(#b33630, #dc5f59);
    background-image: -ms-linear-gradient(#b33630, #dc5f59);
    background-image: -o-linear-gradient(#b33630, #dc5f59);
    background-image: linear-gradient(#b33630, #dc5f59);
}

/* Pill button
   ========================================================================== */

.button.pill {
    border-radius: 50em;
}

/* Disabled button
   ========================================================================== */

.button.disable {
    opacity: 0.5;
}

/* Big button
   ========================================================================== */

.button.big {
    font-size: 14px;
}

.button.big.icon:before {
    top: 0;
}


/* =============================================================================
   Button groups
   ========================================================================== */

/* Standard group
   ========================================================================== */

.button-group {
    display: inline-block;
    list-style: none;
    padding: 0;
    margin: 0;
    /* IE hacks */
    zoom: 1;
    *display: inline;
}

.button + .button,
.button + .button-group,
.button-group + .button,
.button-group + .button-group {
    margin-left: 5px;
}

.button-group li {
    float: left;
    padding: 0;
    margin: 0;
}

.button-group .button {
    float: left;
    margin-left: -1px;
}

.button-group > .button:not(:first-child):not(:last-child),
.button-group li:not(:first-child):not(:last-child) .button {
    border-radius: 0;
}

.button-group > .button:first-child,
.button-group li:first-child .button {
    margin-left: 0;
    border-top-right-radius: 0;
    border-bottom-right-radius: 0;
}

.button-group > .button:last-child,
.button-group li:last-child > .button {
    border-top-left-radius: 0;
    border-bottom-left-radius: 0;
}

/* Minor group
   ========================================================================== */

.button-group.minor-group .button {
    border: 1px solid #d4d4d4;
    text-shadow: none;
    background-image: none;
    background-color: #fff;
}

.button-group.minor-group .button:hover,
.button-group.minor-group .button:focus {
    background-color: #599bdc;
}

.button-group.minor-group .button:active,
.button-group.minor-group .button.active {
    background-color: #3072b3;
}

.button-group.minor-group .button.icon:before {
    opacity: 0.8;
}

/* =============================================================================
   Button container (mixing buttons and groups, e.g., nav bar)
   ========================================================================== */

.button-container .button,
.button-container .button-group {
    vertical-align: top;
}
</style>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %>>
<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
 <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 	<div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;"> 
 <%
'-------------awal-----------------------
 Dim cmdView,PrCode,PrDate
 Dim Description, Remarks, RequestorId
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax, tempMaxLineNo
 Dim ManagedId, PurchaseGroupId, RowStateId
 Dim LineId, ItemCode, RequestedQty, ApprovedQty, OrderedQty
 Dim UnitOfMeasurement, WantedDeliveryDate
 Dim strItemCodeInfo, rsstrItemCodeInfo
 Dim rsPurchase, rsPurchaseTemp
 Dim strQueryStatus, rsQueryStatus
 Dim strQueryAutoNumber
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

d_span = 14
 


 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 PRCODE = Request("PRCODE")
 
 '---- Data Header
 if Request("PRDATE")<>"" then
   PRDATE   =Request("PRDATE")
 else
   PRDATE   =textDate(date)
 end if
 Description = Request("Description")
 Remarks = Request("Remarks")
 RequestorId = Request("RequestorId")
 ManagedId = Request("ManagedId")
 PurchaseGroupId = Request("PurchaseGroupId")
 RowStateId = Request("RowStateId")
 
 
 
 '---- Data Details
 
 LINEID=Request("LineId")
 ITEMCODE=Request("ITEMCODE")
 REQUESTEDQTY=Request("REQUESTEDQTY")
 APPROVEDQTY=Request("APPROVEDQTY")
 IF APPROVEDQTY="" then
    APPROVEDQTY="0"
 end if
 IF APPROVEDQTY="0" then
     APPROVEDQTY=REQUESTEDQTY
 END IF

 ORDEREDQTY=Request("ORDEREDQTY")
 WANTEDDELIVERYDATE = Request ("WANTEDDELIVERYDATE")
 UnitOfMeasurement = Request("UnitOfMeasurement")
 REMARKS=request("REMARKS")
 
 'AMH Mar 24 - Update Payment di All Invoice table
Sub UpdateAmountPaymentInvoice(fccode,fcba)
    Dim jenisinv,txtnoinvoice,rsJenis
	qJenis = " select jenisinv,invoicecode from paymentvoucherdetail  a " & _ 
	         " inner join PaymentVoucher b on (a.VoucherCode=b.VoucherCode and a.fcba=b.fcba) " & _ 
			 " where a.jenisinv <> 0 and a.fcBA = '" & fcba & "'" & _
			 "       AND a.vouchercode = '" & fcCode & "'" 
	set rsJenis= Server.CreateObject("ADODB.Recordset")
	rsJenis.Open qJenis, connstr
	if not rsJenis.eof then
		jenisinv = rsJenis("jenisinv")
		txtnoinvoice = rsJenis("invoicecode")
	end if 
	rsJenis.Close
	set rsJenis = Nothing
	
	
	' update ke tabel invoice untuk kolom payment
	if jenisinv = "1" then
		sqlupdatepay =  " MERGE INTO INVOICE i " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )j ON  (i.invoice_by=j.fcba AND i.invoicecode=j.invoicecode) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET i.payment=j.pay "
		connStr.execute(sqlupdatepay)
	end if
						' update ke tabel invoice tbs
	if jenisinv = "2" then
		sqlupdatepay =  " MERGE INTO invoiceservicehdr i " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )j ON  (i.invoice_by=j.fcba AND i.invoiceno=j.invoicecode) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET i.payment=j.pay "
		connStr.execute(sqlupdatepay)
	end if

	' update ke tabel invoice kontraktor
	if jenisinv = "3" then
		sqlupdatepay = " MERGE INTO CONTRACTORINVOICEHDR i " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )j ON  (i.invoice_by=j.fcba AND i.invoicecode=j.invoicecode) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET i.payment=j.pay "
		connStr.execute(sqlupdatepay)
		'update ke tabel contractorinvoicedp
	
		sqlupdatepay = " MERGE INTO CONTRACTORINVOICEDP i " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )j ON  (i.invoice_by=j.fcba AND i.invoicecode=j.invoicecode) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET i.payment=j.pay "
		connStr.execute(sqlupdatepay)
	end if
	
	' update ke tabel invoice payment request
	if jenisinv = "4" then
		sqlupdatepay = " MERGE INTO INVOICEDPHDR i " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )j ON  (i.invoice_by=j.fcba AND i.invoicecode=j.invoicecode) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET i.payment=j.pay "
		
		connStr.execute(sqlupdatepay)
		sqlupdatepay =  " Merge INTO  purchaseorder a " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )b ON (a.suppliercode=b.suppliercode AND a.pocode=b.pocode AND a.fcba=b.fcba) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET a.DP=b.DP "		
		connStr.execute(sqlupdatepay)
	end if
	' update ke tabel invoice sto
	if jenisinv = "5" then
		sqlupdatepay = " MERGE INTO INVOICESTO i " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )j ON  (i.invoice_by=j.fcba AND i.invoicecode=j.invoicecode) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET i.payment=j.pay "
		connStr.execute(sqlupdatepay)
	end if
	
	if jenisinv = "6" then
		sqlupdatepay = " MERGE INTO invoiceservicehdr i " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )j ON  (i.invoice_by=j.fcba AND i.invoiceno=j.invoicecode) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET i.payment=j.pay "
	
	connStr.execute(sqlupdatepay)
	end if
	
	if jenisinv = "7" then
		sqlupdatepay = " MERGE INTO TRANSPORTERINVOICE i " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )j ON  (i.invoice_by=j.fcba AND i.invoicecode=j.invoicecode) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET i.payment=j.pay "
						
		connStr.execute(sqlupdatepay)
	end if

	if jenisinv = "8" then
		sqlupdatepay = " MERGE INTO salesinvoicehdr i " & _
						" USING " & _
						" ( " & _
						" SELECT d.fcba,jenisinv,d.invoicecode,Sum(amount) pay " & _
						" FROM paymentvoucherdetail d inner join paymentvoucher h on (d.VoucherCode=h.VoucherCode and d.fcba=h.fcba)" & _
						" WHERE jenisinv=" & jenisinv & "" & _
						" AND d.VoucherCode='" & fccode & "'" & _
						" AND d.fcba='" & fcba & "'" & _
						" AND h.rowstateid='Approved' " & _ 
						" GROUP BY d.fcba,jenisinv,invoicecode " & _
						" )j ON  (i.invoice_by=j.fcba AND i.invoiceno=j.invoicecode) " & _
						" WHEN Matched THEN " & _
						" UPDATE SET i.RECEIVEAMT=j.pay "
		connStr.execute(sqlupdatepay)
	end if

End Sub
 
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "ApprovalAction" or Request("txtJenis")= "RejectAction" Then

	if Request("txtJenis")= "ApprovalAction" then
	
		For n=1 to (cdbl(request("no"))-1)
			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
			
						q = " INSERT INTO paymentvoucher_hist ( vouchercode" & _
							"                                           , currentauthorizelevel" & _
							"                                           , fcentry" & _
							"                                           , fcedit" & _
							"                                           , fcip" & _
							"                                           , fcba" & _
							"                                           , lastupdate" & _
							"                                           , lasttime)"& _
							" SELECT vouchercode" & _
							"        , '" & request("authorizelevelcode_"&TRIM(n)) & "' currentauthorizelevel" & _
							"        , '" & Session("CodeName") & "' fcentry" & _
							"        , '" & Session("CodeName") & "' fcedit" & _
							"        , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
							"        , fcba" & _
							"        , SYSDATE lastupdate" & _
							"        , TO_CHAR(SYSDATE, 'hh24:mi') lasttime" & _
							" FROM paymentvoucher" & _
							" WHERE vouchercode = '" & request("vouchercode_"&TRIM(n)) & "'" & _
							"       AND fcba = '"& request("fcba_"&TRIM(n)) & "'"
						'response.write q
						connStr.execute(q)
						
				
				is_final = 0
				q = " SELECT COUNT(1) total FROM authorizelevel al" & _
					" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
					" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
					" WHERE at.document_type = 'PV'" & _
					"       AND DECODE(at.document_type, 'PR', at.document_classification, '0') = '0'" & _
					"       AND al.isactive = 'TRUE'" & _
					"       AND al.finallevel = 'TRUE'" & _
					"       AND al.fccode = '" & request("authorizelevelcode_"&TRIM(n)) & "'" & _
					"       AND aus.authorizetemplatecode = '" & request("authorizetemplatecode_"&TRIM(n)) & "'" & _
					"       AND aus.fcba = '" & request("fcba_"&TRIM(n)) & "'"
				' response.write q
				set rsPurchase = server.CreateObject("Adodb.recordset")
				rsPurchase.Open q, connstr
				IF not rsPurchase.eof then
					is_final = cdbl(to_nol(rsPurchase("total")))
				end if
				
				if is_final > 0 then
					set rsDone = server.CreateObject("Adodb.recordset")
					' qDone = " select PAYMENTVOUCHER_API.getApproved('" & request("vouchercode_"&TRIM(n)) & "','" & request("fcba_"&TRIM(n)) & "') tempDone from dual"
				   
					' rsDone.Open qDone, connstr
					' if Not rsDone.Eof then
					' 	doneApproved = rsDone ("tempDone")
					' End if
					' rsDone.Close : Set rsDone = nothing
					
					' IF doneApproved = "TRUE" THEN 
					'    strQueryStatus = " BEGIN PAYMENTVOUCHER_API.setApproved('" & request("vouchercode_"&TRIM(n)) & "','" & request("fcba_"&TRIM(n)) & "'); end;"
					'    connStr.execute(strQueryStatus)
					' else
					' 	err.number = 1001
					' 	err.description = "PO "&request("vouchercode_"&TRIM(n))&", business unit "&request("fcba_"&TRIM(n))&" has errors according to PR data!"
					' END IF
				
					strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("vouchercode_"&TRIM(n)) & "', 'Approve' ); END; "
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "UPDATE paymentvoucher SET rowstateid = 'Approved' WHERE fcba = '" & request("fcba_"&TRIM(n)) & "' AND vouchercode = '" & request("vouchercode_"&TRIM(n)) & "'"
					connStr.execute(strQueryStatus)
					
					Call UpdateAmountPaymentInvoice(request("vouchercode_"&TRIM(n)),request("fcba_"&TRIM(n)))
					
					if request("transtype_"&TRIM(n)) = "01-RECEIVE" then
						 strQueryStatus = "BEGIN FAC_JOURNAL_FAC_BNK_API.FAC_BNK_RCV_DOCNO('" & request("vouchercode_"&TRIM(n)) & "',TO_NUMBER(TO_CHAR(TO_DATE('" & request("date_"&TRIM(n))  & "','MM/DD/YYYY'),'YYYY')), " & _ 
										  "TO_NUMBER(TO_CHAR(TO_DATE('" & request("date_"&TRIM(n))  & "','MM/DD/YYYY'),'MM')) , '" & request("fcba_"&TRIM(n)) & "','System'); " & _ 
										  "END;"
						'response.write strQueryStatus
						connStr.execute(strQueryStatus)		
						
					end if
					
					if request("transtype_"&TRIM(n)) = "02-PAYMENT" then
						 strQueryStatus = "BEGIN FAC_JOURNAL_FAC_BNK_API.FAC_BNK_PYV_DOCNO('" & request("vouchercode_"&TRIM(n)) & "',TO_NUMBER(TO_CHAR(TO_DATE('" & request("date_"&TRIM(n))  & "','MM/DD/YYYY'),'YYYY')), " & _ 
										  "TO_NUMBER(TO_CHAR(TO_DATE('" & request("date_"&TRIM(n))  & "','MM/DD/YYYY'),'MM')) , '" & request("fcba_"&TRIM(n)) & "','System'); " & _ 
										  "END;"
						'response.write strQueryStatus
						connStr.execute(strQueryStatus)						
					end if
					
					strQueryStatus = "UPDATE paymentvoucher SET postingflg = '1' WHERE fcba = '" & request("fcba_"&TRIM(n)) & "' AND vouchercode = '" & request("vouchercode_"&TRIM(n)) & "'"
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "update BKMJOURNALHDR set postingflg=1 where fcba='" & request("fcba_"&TRIM(n)) & "' AND docno='" & request("vouchercode_"&TRIM(n)) & "'"
					connStr.execute(strQueryStatus)
	
					strQueryStatus = "BEGIN fac_closing_accounting_api.ProsesClosing(TO_NUMBER(TO_CHAR(TO_DATE('" & request("date_"&TRIM(n))  & "','MM/DD/YYYY'),'YYYY')), " & _ 
			                         "TO_NUMBER(TO_CHAR(TO_DATE('" & request("date_"&TRIM(n))  & "','MM/DD/YYYY'),'MM')) , '" & request("fcba_"&TRIM(n)) & "','System'); " & _ 
									 "END;"			
					'response.write strQueryStatus				 
					connStr.execute(strQueryStatus)
				
					strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("vouchercode_"&TRIM(n)) & "', 'Print' ); END; "
					connStr.execute(strQueryStatus)
					
					' strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("pocode_"&TRIM(n)) & "', 'Order' ); END; "
					' connStr.execute(strQueryStatus)
				end if
				
			end if
		next
		
	   cmdView	= "True"
	elseif Request("txtJenis")= "RejectAction" then
		For n=1 to (cdbl(request("no"))-1)
			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				strQueryStatus = "UPDATE paymentvoucher SET rowstateid = 'Rejected' WHERE fcba = '" & request("fcba_"&TRIM(n)) & "' AND vouchercode = '" & request("vouchercode_"&TRIM(n)) & "'"
				connStr.execute(strQueryStatus)
				
				Call UpdateAmountPaymentInvoice(request("vouchercode_"&TRIM(n)),request("fcba_"&TRIM(n)))
				
				' Perbaikan: dulu memakai variabel fcba/vouchercode yang tidak pernah diisi (kosong)
				strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("vouchercode_"&TRIM(n)) & "', 'Reject' ); END; "
				connStr.execute(strQueryStatus)
			end if
		next
		cmdView	= "True"
	else
		cmdView	= "True"
	end if								 
	' Response.Write "sql " & q
	
else
	cmdView	= request("cmdView")
End if
' response.write ManagedId
if err.number <> 0 then
    response.Write  "<div class='error'>" & err.description & "</div>"
End if
 '-------------ISI Payment Form
cmdView = "True"
 
if cmdView = "True" Then
	' q = "  SELECT hdr.fcba," & _
        ' "         hdr.vouchercode," & _
        ' "         hdr.datecreated," & _
        ' "         hdr.transtype," & _
        ' "         SUM (pvd.amount) amount," & _
        ' "         hdr.remarks," & _
        ' "         aus.authorizelevelcode," & _
        ' "         NVL (prh.currentauthorizelevel, 0) currentauthorizelevel," & _
        ' "         aus.authorizetemplatecode" & _
        ' "    FROM paymentvoucherdetail pvd" & _
        ' "         INNER JOIN paymentvoucher hdr" & _
        ' "            ON (pvd.vouchercode = hdr.vouchercode AND pvd.fcba = hdr.fcba)" & _
        ' "         LEFT JOIN authorizetemplate at ON at.document_type = 'PV'" & _
        ' "         LEFT JOIN" & _
        ' "         (  SELECT fcba," & _
        ' "                   vouchercode," & _
        ' "                   NVL (MAX (currentauthorizelevel), 0) currentauthorizelevel" & _
        ' "              FROM paymentvoucher_hist" & _
        ' "          GROUP BY fcba, vouchercode) prh" & _
        ' "            ON (prh.fcba = hdr.fcba AND prh.vouchercode = hdr.vouchercode)" & _
        ' "         INNER JOIN authorizestructure aus" & _
        ' "            ON (    aus.fcba = hdr.fcba" & _
        ' "                AND aus.authorizetemplatecode = at.fccode" & _
        ' "                AND aus.fcusercode = '"& Session("CodeName") &"'" & _
        ' "                AND authorizestructure_api.get_nextlevel (" & _
        ' "                       'PV'," & _
        ' "                       '0'," & _
        ' "                       hdr.fcba," & _
        ' "                       NVL (prh.currentauthorizelevel, 0)) =" & _
        ' "                       aus.authorizelevelcode)" & _
        ' "   WHERE hdr.rowstateid = 'AuthorizeOnProgress'" & _
		' " AND at.document_classification = hdr.document_classification" & _
        ' " GROUP BY hdr.fcba," & _
        ' "         hdr.vouchercode," & _
        ' "         hdr.datecreated," & _
        ' "         hdr.transtype," & _
        ' "         hdr.remarks," & _
        ' "         aus.authorizelevelcode," & _
        ' "         NVL (prh.currentauthorizelevel, 0)," & _
        ' "         aus.authorizetemplatecode" & _
        ' " ORDER BY hdr.fcba, hdr.vouchercode" 

		' q = " SELECT one.fcba," & _
			' "          one.vouchercode," & _
			' "          one.datecreated," & _
			' "          one.transtype," & _
			' "          one.amount," & _
			' "          one.remarks,        " & _
			' "          one.currentauthorizelevel ," & _
			' "          authorizestructure_api.get_nextlevelpv(" & _
			' "                        'PV'," & _
			' "                        one.document_classification," & _
			' "                        one.fcba," & _
			' "                        NVL (one.currentauthorizelevel, 0)," & _
			' "                        one.vouchercode," & _
			' "                        one.fcentry" & _
			' "                        ) authorizelevelcode," & _
			' "          one.document_classification," & _
			' "          one.fccode," & _
			' "          aus.authorizetemplatecode," & _
			' "          one.fcentry   " & _
			' " FROM" & _
			' "   (" & _
			' " SELECT hdr.fcba," & _
			' "          dtl.vouchercode," & _
			' "          hdr.datecreated," & _
			' "          hdr.transtype," & _
			' "          dtl.amount," & _
			' "          hdr.remarks," & _
			' "          NVL (prh.currentauthorizelevel, 0) currentauthorizelevel," & _
			' "          hdr.document_classification," & _
			' "          at.fccode,       " & _
			' "          prh.fcentry" & _
			' "     FROM (  SELECT fcba, vouchercode, fcentry, SUM (pvd.amount) amount" & _
			' "             FROM paymentvoucherdetail pvd" & _
			' "         GROUP BY fcba, vouchercode, fcentry) dtl" & _
			' "          LEFT JOIN paymentvoucher hdr" & _
			' "             ON (dtl.vouchercode = hdr.vouchercode AND dtl.fcba = hdr.fcba)" & _
			' "          LEFT JOIN authorizetemplate at ON at.document_type = 'PV' AND at.document_classification = hdr.document_classification" & _
			' "             AND dtl.amount BETWEEN at.amountrangefrom AND at.amountrangeuntil" & _
			' "          LEFT JOIN" & _
			' "          (  SELECT fcba," & _
			' "                    vouchercode," & _
			' "                    FCENTRY," & _
			' "                    NVL (MAX (currentauthorizelevel), 0) currentauthorizelevel" & _
			' "               FROM paymentvoucher_hist" & _
			' "           GROUP BY fcba, vouchercode, FCENTRY) prh" & _
			' "             ON (prh.fcba = hdr.fcba AND prh.vouchercode = hdr.vouchercode)" & _
			' "    WHERE     hdr.rowstateid = 'AuthorizeOnProgress'" & _
			' "          AND at.document_classification = hdr.document_classification" & _
			' "           AND dtl.amount BETWEEN at.amountrangefrom AND at.amountrangeuntil" & _
			' " ORDER BY hdr.fcba, hdr.vouchercode)one" & _
			' " INNER JOIN authorizestructure aus" & _
			' "             ON (    aus.fcba = one.fcba" & _
			' "                 AND aus.authorizetemplatecode = one.fccode" & _
			' "                 AND aus.fcusercode = '"& Session("CodeName") &"')" 
		
		' Rev Query Select Added By Gibran
			q = " select distinct " & _
                " stage1.businessunit fcba," & _
                " stage1.vouchercode," & _
                " to_char(stage1.voucherdate,'MM/DD/YYYY') datecreated," & _
                " stage1.fccode authorizetemplatecode," & _
                " stage1.document_classification," & _
                " stage1.amount," & _
                " stage1.currentauthorizelevel," & _
                " stage1.authorizelevelname," & _
                " stage1.lastauthorizer," & _
                " stage1.lastactiondate," & _
                " stage1.authorizelevelcode," & _
                " aust.fcusercode," & _
                " stage1.accountingyear," & _
                " stage1.accountingperiod," & _
                " stage1.rowstateid," & _
                " stage1.bankcode," & _
                " stage1.bankname," & _
                " stage1.transtype," & _
                " stage1.remarks" & _
                " from (" & _
                "     select " & _
                "     abc.businessunit," & _
                "     abc.vouchercode," & _
                "     abc.voucherdate," & _
                "     at.fccode," & _
                "     abc.document_classification," & _
                "     abc.amount," & _
                "     abc.currentauthorizelevel," & _
                "     abc.authorizelevelname," & _
                "     abc.lastauthorizer," & _
                "     abc.lastactiondate," & _
                "     authorizestructure_api.get_nextlevelpv (" & _
                "         'PV'," & _
                "         abc.document_classification," & _
                "         abc.businessunit," & _
                "         nvl (abc.currentauthorizelevel, 0)," & _
                "         abc.vouchercode," & _
                "         abc.lastauthorizer" & _
                "     ) as authorizelevelcode," & _
                "     abc.rowstateid," & _
                "     abc.accountingyear," & _
                "     abc.accountingperiod," & _
                "     abc.bankcode," & _
                "     abc.bankname," & _
                "     abc.transtype," & _
                "     abc.remarks" & _
                "     from (  " & _
                "         select " & _
                "         paymentvoucher.fcba as businessunit," & _
                "         paymentvoucher.vouchercode as vouchercode," & _
                "         paymentvoucher.datecreated as voucherdate," & _
                "         paymentvoucher.document_classification," & _
                "         sum(paymentvoucherdetail.amount) as amount," & _
                "         authorizelevel.fccode as currentauthorizelevel," & _
                "         authorizelevel.fcname as authorizelevelname," & _
                "         paymentvoucher_hist.fcentry as lastauthorizer," & _
                "         paymentvoucher_hist.lastupdate as lastactiondate," & _
                "         paymentvoucher.rowstateid," & _
                "         paymentvoucher.remarks," & _
                "         paymentvoucher.transtype," & _
                "         paymentvoucher.bankcode," & _
                "         bank_api.get_bankcode_name (" & _
                "             paymentvoucher.bankcode," & _
                "             paymentvoucher.fcba" & _
                "         ) bankname," & _
                "         paymentvoucher.accountingperiod," & _
                "         paymentvoucher.accountingyear" & _
                "         from paymentvoucher_hist" & _
                "         inner join paymentvoucherdetail" & _
                "         on (" & _
                "         paymentvoucher_hist.fcba = paymentvoucherdetail.fcba and " & _
                "         paymentvoucher_hist.vouchercode = paymentvoucherdetail.vouchercode)" & _
                "         inner join paymentvoucher" & _
                "         on (    " & _
                "         paymentvoucherdetail.fcba = paymentvoucher.fcba and " & _
                "         paymentvoucherdetail.vouchercode = paymentvoucher.vouchercode)" & _
                "         inner join authorizelevel" & _
                "         on paymentvoucher_hist.currentauthorizelevel = authorizelevel.fccode" & _
                "         inner join (" & _
                "             select " & _
                "             paymentvoucher_hist.vouchercode," & _
                "             max (paymentvoucher_hist.currentauthorizelevel) alevel" & _
                "             from paymentvoucher_hist" & _
                "             group by paymentvoucher_hist.vouchercode" & _
                "         ) aa" & _
                "         on (    " & _
                "         paymentvoucher_hist.currentauthorizelevel = aa.alevel and " & _
                "         paymentvoucher_hist.vouchercode = aa.vouchercode)" & _
                "         group by " & _
                "         paymentvoucher.fcba," & _
                "         paymentvoucher.vouchercode," & _
                "         paymentvoucher.datecreated," & _
                "         paymentvoucher.document_classification," & _
                "         authorizelevel.fccode," & _
                "         authorizelevel.fcname," & _
                "         paymentvoucher.accountingperiod," & _
                "         paymentvoucher.accountingyear," & _
                "         paymentvoucher_hist.fcentry," & _
                "         paymentvoucher_hist.lastupdate," & _
                "         paymentvoucher.rowstateid," & _
                "         paymentvoucher.bankcode," & _
                "         bank_api.get_bankcode_name (" & _
                "             paymentvoucher.bankcode," & _
                "             paymentvoucher.fcba" & _
                "         )," & _
                "         paymentvoucher.remarks," & _
                "         paymentvoucher.transtype" & _
                "     ) abc" & _
                "     left join authorizetemplate at" & _
                "     on abc.document_classification = at.document_classification" & _
                "     left join authorizestructure aus" & _
                "     on (    " & _
                "     at.fccode = aus.authorizetemplatecode and " & _
                "     abc.currentauthorizelevel = aus.authorizelevelcode and " & _
                "     abc.businessunit = aus.fcba and " & _
                "     aus.fcusercode = abc.lastauthorizer)" & _
                "     where abc.amount between at.amountrangefrom and at.amountrangeuntil" & _
                " ) stage1" & _
                " inner join authorizestructure aust" & _
                " on (    " & _
                " stage1.businessunit = aust.fcba and " & _
                " stage1.authorizelevelcode = aust.authorizelevelcode and " & _
                " stage1.fccode = aust.authorizetemplatecode and " & _
                " aust.fcusercode = '"& Session("CodeName") &"')" & _
				" and stage1.businessunit = '" & Session("fcba") & "'" & _
                " and stage1.rowstateid = 'AuthorizeOnProgress'" & _
                " order by " & _
                " stage1.vouchercode" 
		' End Added
	' response.write q
	Set rsPurchaseRequest = Server.CreateObject("ADODB.Recordset")
	rsPurchaseRequest.Open q, connstr
%>
<form action="PaymentVoucherAuthorization.Asp" method="post" name="eForm" >
   <table class="data" style="width:100%;">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="center">No</TD>
	<TD align="center">Kebun/Pabrik</TD>
	<TD align="center">Voucher Code</TD>
    <TD align="center">Accounting Year</TD>
    <TD align="center">Accounting Period</TD>
    <TD align="center">Bank Code</TD>
    <TD align="center">Bank Name</TD>
	<TD align="center">Date</TD>
	<TD align="center">Transaction Type</TD>
	<TD align="center">Paid For</TD>
	<TD align="center">Amount</TD>
	<TD align="center"><input type='checkbox' name='select-all' id='select-all'/></TD>	
	</tr> 
	
   <%
   if Not rsPurchaseRequest.eof then
      nCountofField=rsPurchaseRequest.fields.count-1
	  no=1
	  While Not rsPurchaseRequest.EOF
		if no  mod 2 = 1 then
			cColor = "#EEEEEE"
		else
			cColor = "#E5E5E5"
		end if%>
		<tr bgcolor="<%=cColor%>">
			<td align="center" valign="top"><%=NO%></td>
			
			<td align="center" title="<%=rsPurchaseRequest("fcba").Name%>" valign="top"><%=rsPurchaseRequest("fcba")%></td>
			<td align="center" title="<%=rsPurchaseRequest("vouchercode").Name%>" valign="top">
				<a href="PaymentVoucherApproval.asp?fcba=<%=rsPurchaseRequest("fcba")%>&vouchercode=<%=rsPurchaseRequest("vouchercode")%>&authorizelevelcode=<%=rsPurchaseRequest("authorizelevelcode")%>&authorizetemplatecode=<%=rsPurchaseRequest("authorizetemplatecode")%>" target="_blank">
				<%=rsPurchaseRequest("vouchercode")%>
				</a>
			</td>
			<td align="center" title="<%=rsPurchaseRequest("accountingyear").Name%>" valign="top"><%=rsPurchaseRequest("accountingyear")%></td>
            <td align="center" title="<%=rsPurchaseRequest("accountingperiod").Name%>" valign="top"><%=monthname(rsPurchaseRequest("accountingperiod"))%></td>
            <td align="center" title="<%=rsPurchaseRequest("bankcode").Name%>" valign="top"><%=rsPurchaseRequest("bankcode")%></td>
            <td align="center" title="<%=rsPurchaseRequest("bankname").Name%>" valign="top"><%=rsPurchaseRequest("bankname")%></td>
            <td align="center" title="<%=rsPurchaseRequest("datecreated").Name%>" valign="top"><%=rsPurchaseRequest("datecreated")%></td>
			<td align="center" title="<%=rsPurchaseRequest("transtype").Name%>" valign="top"><%=rsPurchaseRequest("transtype")%></td>
			<td align="left" title="<%=rsPurchaseRequest("remarks").Name%>" valign="top"><%=rsPurchaseRequest("remarks")%></td>
			<td align="right" title="<%=rsPurchaseRequest("amount").Name%>" valign="top"><%=FormatNumber(rsPurchaseRequest("amount"),0)%></td>
			<td valign="middle" align="center">
				<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);'>
				<input type="hidden" name="fcba_<%=NO%>" value="<%=rsPurchaseRequest("fcba")%>">
				<input type="hidden" name="vouchercode_<%=NO%>" value="<%=rsPurchaseRequest("vouchercode")%>">
				<input type="hidden" name="authorizelevelcode_<%=NO%>" value="<%=rsPurchaseRequest("authorizelevelcode")%>">
				<input type="hidden" name="authorizetemplatecode_<%=NO%>" value="<%=rsPurchaseRequest("authorizetemplatecode")%>">
				<input type="hidden" name="transtype_<%=NO%>" value="<%=rsPurchaseRequest("transtype")%>">
				<input type="hidden" name="year_<%=NO%>" value="<%=rsPurchaseRequest("accountingyear")%>">
				<input type="hidden" name="period_<%=NO%>" value="<%=rsPurchaseRequest("accountingperiod")%>">
				<input type="hidden" name="date_<%=NO%>" value="<%=rsPurchaseRequest("datecreated")%>">
				
								
				
			</td>
		  </tr>
		  <%
		  rsPurchaseRequest.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write"<tr><td colspan='"&d_span&"' align='left'>Data is empty</td></tr>"
   end if 
   rsPurchaseRequest.Close : set rsPurchaseRequest = Nothing
%>
		<tr>
			<td colspan="<%=d_span%>" align="left">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Hidden" Value="<%=no%>" Name="no">
				<button type="submit" class="button" name="txtJenis" value="ApprovalAction" onClick="return confirm('Anda yakin untuk melakukan persetujuan terhadap Payment Voucher yang sudah dipilih?')"><i class="fa fa-check"></i> Approve</button>
				<button type="submit" class="button danger" name="txtJenis" value="RejectAction" onClick="return confirm('Anda yakin untuk mereject Payment Voucher yang sudah dipilih?');"><i class="fa fa-times"></i> Reject</button>
			</td>
		</tr>
   </table>
   </form>
<%End if%>
</div></div>
  <!--#include file="bottom.jsp"-->
<script type="text/javascript">
	function handleClick(cb){
		if(cb.checked == 1){
			cb.value = 1;
		} else {
			cb.value = 0;
		}
	}
	

</script>
<script type="text/javascript">	
		$('#select-all').click( function () {
			$('.item').prop('checked', this.checked);
			$('.item').val(this.checked ? 1 : 0);
		})
	</script>
</body>
</html>
