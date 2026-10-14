<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Transport Task Approval</title>
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
</head>
<body>
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
 Dim errnumber : errnumber = 0 
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
  JobCode=Request("JobCode")
 LocationType="CB"
 LocationCode="BS"
 Reference = Request("Reference")
 ITEMCODE=Request("ITEMCODE")
 STORECODE = Request("STORECODE")
 LineNo=Request("LineNo")
 QuantityTransfered=Request("QuantityTransfered")
 QuantityReceived=Request("QuantityReceived")
 Price = Request("Price")
 AMOUNT =Request("AMOUNT")
 Linenohist = Request("linenohist")
 FcbaDistination = Request("FcbaDistination")
 document_classification = Request("document_classification")
 comments = request ("comments")
 orderno = request ("orderno")
 
 
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
 
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "ApprovalAction" or Request("txtJenis")= "RejectAction" Then

	if Request("txtJenis")= "ApprovalAction" then
	
		For n=1 to (cdbl(request("no"))-1)
			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
			
				q = " INSERT INTO transporttask_hist ( orderno" & _
					"                                           , currentauthorizelevel" & _
					"                                           , fcentry" & _
					"                                           , fcedit" & _
					"                                           , fcip" & _
					"                                           , fcba" & _
					"                                           , lastupdate" & _
					"                                           , lasttime)"& _
					" SELECT orderno" & _
					"        , '" & request("authorizelevelcode_"&TRIM(n)) & "' currentauthorizelevel" & _
					"        , '" & Session("CodeName") & "' fcentry" & _
					"        , '" & Session("CodeName") & "' fcedit" & _
					"        , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
					"        , fcba" & _
					"        , SYSDATE lastupdate" & _
					"        , TO_CHAR(SYSDATE, 'hh24:mi') lasttime" & _
					" FROM transporttask_tab" & _
					" WHERE orderno = '" & request("orderno_"&TRIM(n)) & "'" & _
					"       AND fcba = '"& request("fcba_"&TRIM(n)) & "'"
				' response.write q
				connStr.execute(q)
						
				' Insert STock On Hand dan Stock Trans Hist
				 qLineMax = "select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
							"where " & _
							"fcba = '" & request("fcba_"&TRIM(n)) & "'" 
			   
			   ' response.write qLineMax
				set rsqLineMax = Server.CreateObject("ADODB.Recordset")
				rsqLineMax.Open qLineMax, connstr
				if not rsqLineMax.eof then
					tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
				else
					tempMaxLine = 0
				end if 
		
				set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
				   q = 	" SELECT td.ORDERNO," & _
						"        tt.orderdate," & _
						"        td.LINENO," & _
						"        td.LINENOHIST," & _
						"        td.STOCKCODE," & _
						"        stock_api.get_stockName (Stockcode, td.fcba) StockName," & _
						"        td.JOBCODE," & _
						"        td.LOCATIONTYPE," & _
						"        td.LOCATIONCODE," & _
						"        td.QUANTITYTRANSFERED," & _
						"		( -1 * td.QUANTITYTRANSFERED) QUANTITYTRANSFERED_INPUT, " & _
						"        td.STORECODE," & _
						"        td.FCBADISTINATION," & _
						"        td.QUANTITYRECEIVED," & _
						"        td.Price," & _
						"        td.AMOUNT," & _
						"        td.REMARKS," & _
						"        td.FCENTRY," & _
						"        td.FCEDIT," & _
						"        td.FCIP," & _
						"        td.FCBA" & _
						"   FROM TransportTaskDtl_tab td" & _
						"        INNER JOIN transporttask_tab tt" & _
						"           ON (tt.orderno = td.orderno AND tt.fcba = td.fcba)" & _
						"  WHERE td.OrderNo = '" & request("orderno_"&TRIM(n)) & "' AND td.fcBA = '" & request("fcba_"&TRIM(n)) & "'"

				   ' response.Write q 
					Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
					rsPurchaseOrder.Open q, connstr
					if Not rsPurchaseOrder.Eof then
						Do While Not rsPurchaseOrder.eof
						QUANTITYTRANSFERED = rsPurchaseOrder("QUANTITYTRANSFERED")  
						' Add by NAF to handle Inventory Stock On Hand
						' Insert the History of Received Inventory Transaction
						
						'qLessZero = "select islesszero from stock " & _    
						'    "where " & _
						'    "fccode = '" & rsPurchaseOrder("STOCKCODE") & "' and " & _
						'    "fcba = '" & rsPurchaseOrder("fcba") & "'" 
						' response.write qLessZero
						'set rsLessZero = Server.CreateObject("ADODB.Recordset")
						'rsLessZero.Open qLessZero, connstr
						'if not rsLessZero.eof then
						'    tempislesszero = rsLessZero("islesszero")
						'else
						'    tempislesszero = "FALSE"
						'end if 
							
						' We have to Update the Quantity by Adding them with the New one.
						
						
						'qQuantityOnHand = "select quantity from stockonhand " & _
						'				  "where " & _
						'				  "storecode = '" & rsPurchaseOrder("STORECODE") & "' and " & _
						'				  "stockcode  = '" & rsPurchaseOrder("STOCKCODE") & "' and " & _ 
						'				  "fcba = '" & rsPurchaseOrder("fcba") &"' "
			
						' response.write  qQuantityOnHand     
						
						'set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
						'rsQuantityOnHand.Open qQuantityOnHand, connstr
						'if not rsQuantityOnHand.eof then
						'    tempQuantityOnHand = cdbl(to_nol(rsQuantityOnHand("quantity")))
						'else
						'    tempQuantityOnHand = 0
						'end if      

						
						'if (cdbl(to_nol(tempQuantityOnHand)) - cdbl(to_nol(QuantityTransfered)) >= 0) or tempislesszero = "TRUE"  then 
						
							tempMaxLine = tempMaxLine + 1 
							
							q =  " Insert into StockTransHist " & _
								 "(LINENO, FCBA, DOCUMENTNO, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, " & _
								 " SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY,  " & _ 
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & tempMaxLine & ",'" & rsPurchaseOrder("FCBA") & "','" & rsPurchaseOrder("ORDERNO") & "','" & rsPurchaseOrder("STOCKCODE") & "'" & _ 
								 " ,to_date('" & rsPurchaseOrder("ORDERDATE") & "','mm/dd/yyyy') " & ",'ST_TransportTask','" & rsPurchaseOrder("STORECODE") & "'," & _
								 " 'ST','" & rsPurchaseOrder("STORECODE") & "'," & _
								 " '" & rsPurchaseOrder("LOCATIONTYPE") & "','" & rsPurchaseOrder("LOCATIONCODE") & "'" & _
								 " ," & rsPurchaseOrder("QUANTITYTRANSFERED_INPUT") & ", sysdate, to_char(sysdate,'HH24:MI')) " 
							' response.write q 
							connStr.execute(q)
							  
							 q = 	" SELECT lineno, documentno, fcba, stockcode " & _
									"   FROM stocktranshist " & _  
									"  WHERE documentno = '" & rsPurchaseOrder("ORDERNO") & "' AND fcBA = '" & rsPurchaseOrder("fcba") & "'"  
							   ' response.Write q 
							   Set rsPurchaseOrders = Server.CreateObject("ADODB.Recordset")
							   rsPurchaseOrders.Open q, connstr
									if Not rsPurchaseOrders.Eof then
										Do While Not rsPurchaseOrders.eof
											q = " update TransportTaskDtl_tab " & _
												" set LINENOHIST = " & rsPurchaseOrders("lineno") & " " & _
												" , LASTUPDATE = sysdate " & _
												" , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
												" where " & _
												" ORDERNO = '" & rsPurchaseOrders("documentno") & "' and " & _
												" fcba = '" & rsPurchaseOrders("fcBA") &"'" & _
												"  AND stockcode = '"& rsPurchaseOrders("STOCKCODE") &"'" 
													 
												connStr.execute(q)
											rsPurchaseOrders.MoveNext
										loop 
									end if  
							' q=  "Insert into  TransportTaskDtl_tab (ORDERNO, STOCKCODE,LINENO,LINENOHIST, LOCATIONCODE,LOCATIONTYPE, QUANTITYTRANSFERED,STORECODE,JOBCODE,FCBADISTINATION,QUANTITYRECEIVED,PRICE,AMOUNT, " & _
							'   "REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
							'   "values ('" & OrderNo & "','" & ITEMCODE & "','" & tempMaxLineNo & "'," & tempMaxLine &",'"  &  LOCATIONCODE & "','" & LOCATIONTYPE & "'" & _
							'   "," & QuantityTransfered & ",'" & STORECODE & "','" & Jobcode & "','" & FCBADISTINATION & "',0," &  realprice & "," &  realAmount & ",'" & detail_remarks & "','" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
							'   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
							'   'response.Write(q)
							' connStr.execute(q)

							'qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
							'                     "where " & _
							'                     "storecode = '" & rsPurchaseOrder("STORECODE") & "' and " & _                                              
							'                     "stockcode = '" & rsPurchaseOrder("STOCKCODE") & "' and " & _
							'                     "fcba = '" &  rsPurchaseOrder("fcBA") & "'"

							'set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
							'rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
							'if not rsTotalRecordOnHand.eof then
							'    tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
							'else
							'    tempTotalRecordOnHand = 0
							'end if 
									
							'if tempTotalRecordOnHand > 0 then
							
							'    q = "update stockonhand " & _
							'        "set quantity = quantity - " & QuantityTransfered & " " & _
							'        ", LASTUPDATE = sysdate " & _
							'        ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							'        "where " & _
							'        "storecode = '" & rsPurchaseOrder("STORECODE") & "' and " & _
							'        "stockcode  = '" & rsPurchaseOrder("STOCKCODE") & "' and " & _ 
							'        "fcba = '" & rsPurchaseOrder("fcBA") &"' "
									 
							'    connStr.execute(q)
							'else
								' Just Insert the New Record with the following/received Quantity
							'    q = " insert into stockonhand (storecode" & _
							'        "                          , stockcode" & _
							'        "                          , fcba" & _
							'        "                          , quantity" & _
							'        "                          , lastupdate" & _
							'        "                          , lasttime )" & _
							'        " values ('" & rsPurchaseOrder("STORECODE") & "'" & _
							'        "         , '" & rsPurchaseOrder("STOCKCODE") & "'" & _
							'        "         , '" & rsPurchaseOrder("fcBA") & "'" & _
							'        "         , '0'" & _
							'        "         , sysdate" & _
							'        "         , To_Char(sysdate,'HH24:MI') ) "
									' response.write q
							'    connStr.execute(q)
								
							'    q = " update stockonhand " & _
							'        " set quantity = quantity - " & QuantityTransfered & " " & _
							'        "     , LASTUPDATE = sysdate " & _
							'        "     , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							'        " where storecode = '" & rsPurchaseOrder("STORECODE") & "'" & _
							'        "       and stockcode  = '" & rsPurchaseOrder("STOCKCODE") & "'" & _ 
							'        "       and fcba = '" & rsPurchaseOrder("fcBA") & "'"
									 
							'    connStr.execute(q)
							'end if
						'else
						'    errnumber = 1001
						'    err.number = 1001
						'    err.description = "Nilai Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
						'end if
						rsPurchaseOrder.MoveNext
						loop 
					End if

        ' End Insert
				is_final = 0
				q = " SELECT COUNT(1) total FROM authorizelevel al" & _
					" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
					" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
					" WHERE at.document_type = 'TT'" & _
					"       AND DECODE(at.document_type, 'PR', at.document_classification, '0') = '0'" & _
					"       AND al.isactive = 'TRUE'" & _
					"       AND al.finallevel = 'TRUE'" & _
					"       AND al.fccode = '" & request("authorizelevelcode_"&TRIM(n)) & "'" & _
					"       AND aus.authorizetemplatecode = '" & request("authorizetemplatecode_"&TRIM(n)) & "'" & _
					"       AND aus.fcba = '" & request("fcba_"&TRIM(n)) & "'" & _
					"       AND aus.fcusercode = '"& session("CodeName") &"'"
					
				set rsPurchase = server.CreateObject("Adodb.recordset")
				rsPurchase.Open q, connstr
				IF not rsPurchase.eof then
					is_final = cdbl(to_nol(rsPurchase("total")))
				end if
				
				if is_final = 1  then
					' set rsDone = server.CreateObject("Adodb.recordset")
					' qDone = " select TRANSPORTTASK_API.getApproved('" & request("orderno_"&TRIM(n)) & "','" & request("fcba_"&TRIM(n)) & "') tempDone from dual"
				   
					' rsDone.Open qDone, connstr
					' if Not rsDone.Eof then
					' 	doneApproved = rsDone ("tempDone")
					' End if
					' rsDone.Close : Set rsDone = nothing
					
					' IF doneApproved = "TRUE" THEN 
					'    strQueryStatus = " begin TRANSPORTTASK_API.setApproved('" & request("orderno_"&TRIM(n)) & "','" & request("fcba_"&TRIM(n)) & "'); end;"
					'    connStr.execute(strQueryStatus)
					' else
					' 	err.number = 1001
					' 	err.description = "Order No "&request("orderno_"&TRIM(n))&", business unit "&request("fcba_"&TRIM(n))&" has errors according to Transport Task data!"
					' END IF
				
					strQueryStatus = "BEGIN TRANSPORTTASK_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("orderno_"&TRIM(n)) & "', 'Approve' ); END; "
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "UPDATE transporttask_tab SET rowstateid = 'Approved' WHERE fcba = '" & request("fcba_"&TRIM(n)) & "' AND orderno = '" & request("orderno_"&TRIM(n)) & "'"
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "BEGIN TRANSPORTTASK_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("orderno_"&TRIM(n)) & "', 'Print' ); END; "
					connStr.execute(strQueryStatus)
					
					' strQueryStatus = "BEGIN TRANSPORTTASK_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("orderno_"&TRIM(n)) & "', 'Order' ); END; "
					' connStr.execute(strQueryStatus)
				end if
				
			end if
		next
		
	   cmdView	= "True"
	elseif Request("txtJenis")= "RejectAction" then
		For n=1 to (cdbl(request("no"))-1)
			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				strQueryStatus = "UPDATE transporttask_tab SET rowstateid = 'Rejected' WHERE fcba = '" & request("fcba_"&TRIM(n)) & "' AND orderno = '" & request("orderno_"&TRIM(n)) & "'"
				connStr.execute(strQueryStatus)
				
				strQueryStatus = "BEGIN TRANSPORTTASK_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("orderno_"&TRIM(n)) & "', 'Reject' ); END; "
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
	q = " SELECT hdr.fcba, hdr.orderno, hdr.orderdate, hdr.remarks, aus.authorizelevelcode, NVL(prh.currentauthorizelevel,0) currentauthorizelevel,aus.authorizetemplatecode" & _
	    " FROM transporttask_tab hdr" & _
	    " LEFT JOIN authorizetemplate at ON at.document_type = 'TT' " & _
	    " LEFT JOIN ( SELECT fcba, orderno, NVL(MAX(currentauthorizelevel),0) currentauthorizelevel from transporttask_hist group by fcba, orderno) prh ON (prh.fcba = hdr.fcba AND prh.orderno = hdr.orderno)" & _
		" INNER JOIN authorizestructure aus ON (aus.fcba = hdr.fcba AND aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & Session("CodeName") & "' AND authorizestructure_api.get_nextlevel('TT', '0', hdr.fcba, NVL(prh.currentauthorizelevel,0)) = aus.authorizelevelcode )" & _
	    " WHERE hdr.rowstateid = 'AuthorizeOnProgress'" & _
		" AND at.document_classification = hdr.document_classification" & _
		" ORDER BY hdr.fcba, hdr.orderno"
	'response.write q
	Set rsPurchaseRequest = Server.CreateObject("ADODB.Recordset")
	rsPurchaseRequest.Open q, connstr
%>
<form action="transporttaskauthorization.Asp" method="post" name="eForm" >
   <table class="data" style="width:100%;">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="center">No</TD>
	<TD align="center">Kebun/Pabrik</TD>
	<TD align="center">Order No</TD>
	<TD align="center">Order Date</TD>
	<TD align="center">Remarks</TD>
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
			<td align="right" valign="top"><%=NO%></td>
			
			<td align="left" title="<%=rsPurchaseRequest("fcba").Name%>" valign="top"><%=rsPurchaseRequest("fcba")%></td>
			<td align="left" title="<%=rsPurchaseRequest("orderno").Name%>" valign="top">
				<a href="TransporttaskApproval.asp?fcba=<%=rsPurchaseRequest("fcba")%>&orderno=<%=rsPurchaseRequest("orderno")%>&authorizelevelcode=<%=rsPurchaseRequest("authorizelevelcode")%>&authorizetemplatecode=<%=rsPurchaseRequest("authorizetemplatecode")%>" target="_blank">
				<%=rsPurchaseRequest("orderno")%>
				</a>
			</td>
			<td align="left" title="<%=rsPurchaseRequest("orderdate").Name%>" valign="top"><%=rsPurchaseRequest("orderdate")%></td>
			<td align="left" title="<%=rsPurchaseRequest("remarks").Name%>" valign="top"><%=rsPurchaseRequest("remarks")%></td>
			<td valign="middle" align="center">
				<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);'>
				<input type="hidden" name="fcba_<%=NO%>" value="<%=rsPurchaseRequest("fcba")%>">
				<input type="hidden" name="orderno_<%=NO%>" value="<%=rsPurchaseRequest("orderno")%>">
				<input type="hidden" name="authorizelevelcode_<%=NO%>" value="<%=rsPurchaseRequest("authorizelevelcode")%>">
				<input type="hidden" name="authorizetemplatecode_<%=NO%>" value="<%=rsPurchaseRequest("authorizetemplatecode")%>">
			</td>
		  </tr>
		  <%
		  rsPurchaseRequest.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      no=0
      response.Write"<tr><td colspan='"&d_span&"' align='left'>Data is empty</td></tr>"
   end if 
   rsPurchaseRequest.Close : set rsPurchaseRequest = Nothing
%>
		<tr>
			<td colspan="<%=d_span%>" align="left">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Hidden" Value="<%=no%>" Name="no">
				<button type="submit" class="button" name="txtJenis" value="ApprovalAction" onClick="return confirm('Anda yakin untuk melakukan persetujuan terhadap Transport Task yang sudah dipilih?')"><i class="fa fa-check"></i> Approve</button>
				<button type="submit" class="button danger" name="txtJenis" value="RejectAction" onClick="return confirm('Anda yakin untuk mereject Transport Task yang sudah dipilih?');"><i class="fa fa-times"></i> Reject</button>
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
