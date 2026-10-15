<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<!--#include file="email/po_mail_queue.inc"-->
<html>
<head>
<title>ePlaS - Purchase Order Approval</title>
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
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-table.css + css/fbp-approval.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-table.css?v=20261010c" type="text/css" />
<link rel="StyleSheet" href="css/fbp-approval.css?v=20261007b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<script type="text/javascript" src="js/fbp-table.js?v=20261009c"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-approval"<% end if %>>
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

d_span = 8
 


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
 
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "ApprovalAction" or Request("txtJenis")= "RejectAction" Then

	if Request("txtJenis")= "ApprovalAction" then
	
		For n=1 to (cdbl(request("no"))-1)
			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
			
						q = " INSERT INTO purchaseorderdetails_hist ( pocode" & _
							"                                           , currentauthorizelevel" & _
							"                                           , fcentry" & _
							"                                           , fcedit" & _
							"                                           , fcip" & _
							"                                           , fcba" & _
							"                                           , lastupdate" & _
							"                                           , lasttime)"& _
							" SELECT pocode" & _
							"        , '" & request("authorizelevelcode_"&TRIM(n)) & "' currentauthorizelevel" & _
							"        , '" & Session("CodeName") & "' fcentry" & _
							"        , '" & Session("CodeName") & "' fcedit" & _
							"        , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
							"        , fcba" & _
							"        , SYSDATE lastupdate" & _
							"        , TO_CHAR(SYSDATE, 'hh24:mi') lasttime" & _
							" FROM purchaseorder" & _
							" WHERE pocode = '" & request("pocode_"&TRIM(n)) & "'" & _
							"       AND fcba = '"& request("fcba_"&TRIM(n)) & "'"
						'response.write q
						connStr.execute(q)
						
						q = " INSERT INTO po_authorization_comments ( pocode" & _
							"                                         , currentauthorizelevel" & _
							"                                         , comments" & _
							"                                         , fcentry" & _
							"                                         , fcedit" & _
							"                                         , fcip" & _
							"                                         , fcba" & _
							"                                         , lastupdate" & _
							"                                         , lasttime)"& _
							" VALUES ( '" & request("pocode_"&TRIM(n)) & "'" & _
							"          , '" & request("authorizelevelcode_"&TRIM(n)) & "'" & _
							"          , '" & comments & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & request("fcba_"&TRIM(n)) & "'" & _
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
						'response.write q
						connStr.execute(q)
						
				
				is_final = 0
				q = " SELECT COUNT(1) total FROM authorizelevel al" & _
					" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
					" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
					" WHERE at.document_type = 'PO'" & _
					"       AND DECODE(at.document_type, 'PR', at.document_classification, '0') = '0'" & _
					"       AND al.isactive = 'TRUE'" & _
					"       AND al.finallevel = 'TRUE'" & _
					"       AND al.fccode = '" & request("authorizelevelcode_"&TRIM(n)) & "'" & _
					"       AND aus.authorizetemplatecode = '" & request("authorizetemplatecode_"&TRIM(n)) & "'" & _
					"       AND aus.fcba = '" & request("fcba_"&TRIM(n)) & "'" & _ 
				    "       AND aus.fcusercode = '" & Session("CodeName") & "'"
				
				set rsPurchase = server.CreateObject("Adodb.recordset")
				rsPurchase.Open q, connstr
				IF not rsPurchase.eof then
					is_final = cdbl(to_nol(rsPurchase("total")))
				end if
				
				if is_final = 1 then
					set rsDone = server.CreateObject("Adodb.recordset")
					qDone = " select purchaseorder_api.getApproved('" & request("pocode_"&TRIM(n)) & "','" & request("fcba_"&TRIM(n)) & "') tempDone from dual"
				   
					rsDone.Open qDone, connstr
					if Not rsDone.Eof then
						doneApproved = rsDone ("tempDone")
					End if
					rsDone.Close : Set rsDone = nothing
					
					' IF doneApproved = "TRUE" THEN 
					   strQueryStatus = " begin purchaseorder_api.setApproved('" & request("pocode_"&TRIM(n)) & "','" & request("fcba_"&TRIM(n)) & "'); end;"
					   connStr.execute(strQueryStatus)
					' else
					' 	err.number = 1001
					' 	err.description = "PO "&request("pocode_"&TRIM(n))&", business unit "&request("fcba_"&TRIM(n))&" has errors according to PR data!"
					' END IF
				
					strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("pocode_"&TRIM(n)) & "', 'Approve' ); END; "
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "UPDATE purchaseorderdetails SET rowstateid = 'Ordered' WHERE fcba = '" & request("fcba_"&TRIM(n)) & "' AND pocode = '" & request("pocode_"&TRIM(n)) & "'"
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("pocode_"&TRIM(n)) & "', 'Print' ); END; "
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("pocode_"&TRIM(n)) & "', 'Order' ); END; "
					connStr.execute(strQueryStatus)
					
					sqlupdateqtypr = " select prcode, prlineid, quantity, fcba from purchaseorderdetails" & _
									 " where pocode = '" & request("pocode_"&TRIM(n)) & "'" & _
									 "       and fcba = '" & request("fcba_"&TRIM(n)) & "'" & _
									 "       and prcode is not null" & _
									 "       and prlineid is not null"
									 
					Set rsupdateqtypr = Server.CreateObject("ADODB.Recordset")
					rsupdateqtypr.Open sqlupdateqtypr, connstr, adLockOptimistic, adOpenStatic
					
					if not rsupdateqtypr.EOF then
						prlineid = rsupdateqtypr ("prlineid")
						quantity = rsupdateqtypr ("quantity")
						prcode = rsupdateqtypr ("prcode")
						fcba = rsupdateqtypr ("fcba")
					end if
								
					rsupdateqtypr.Close
					set rsupdateqtypr = Nothing 
					
					q = " update purchaserequestdetails" & _
						" set orderedqty = '" & quantity & "'" & _
						" where  prcode = '" & prcode & "'" & _
						"        and lineid = '" & prlineid & "'" & _
						"        and fcba = '" & fcba & "'"
					' Response.write q
					connStr.execute(q)
					
					' Approver terakhir: antrikan email PO ke supplier (PDF + QR), lihat email/po_mail_queue.inc
					If Err.Number = 0 Then Call PoMail_Queue(request("fcba_"&TRIM(n)), request("pocode_"&TRIM(n)))
				end if
				
			end if
		next
		
		' Kirim email PO yang baru diantrikan (cadangan: Task Scheduler email/run_po_outbox.vbs)
		Call PoMail_Kick()
		
		
	   cmdView	= "True"
	elseif Request("txtJenis")= "RejectAction" then
		For n=1 to (cdbl(request("no"))-1)
			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				strQueryStatus = "UPDATE purchaseorderdetails SET rowstateid = 'Rejected' WHERE fcba = '" & request("fcba_"&TRIM(n)) & "' AND pocode = '" & request("pocode_"&TRIM(n)) & "'"
				connStr.execute(strQueryStatus)
				
				strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("pocode_"&TRIM(n)) & "', 'Reject' ); END; "
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
	q = " SELECT hdr.fcba, hdr.pocode, hdr.podate, hdr.purchaseby, hdr.grandtotal, hdr.remarks, aus.authorizelevelcode, NVL(hist.currentauthorizelevel,0) currentauthorizelevel, aus.authorizetemplatecode" & _
	    " FROM purchaseorder hdr" & _
	    " LEFT JOIN authorizetemplate at ON at.document_type = 'PO'" & _
	    " LEFT JOIN ( SELECT fcba, pocode, NVL(MAX(currentauthorizelevel),0) currentauthorizelevel FROM purchaseorderdetails_hist group by fcba, pocode) hist ON (hist.fcba = hdr.fcba AND hist.pocode = hdr.pocode)" & _
		" INNER JOIN authorizestructure aus ON (aus.fcba = hdr.fcba AND aus.fcba = hist.fcba AND aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & Session("CodeName") & "' AND authorizestructure_api.get_nextlevelnonstart('PO', '0', hdr.fcba, NVL(hist.currentauthorizelevel,0), hdr.pocode, hdr.fcentry) = aus.authorizelevelcode )" & _
	    " WHERE hdr.rowstateid = 'AuthorizeOnProgress' AND hdr.grandtotal BETWEEN at.amountrangefrom AND at.amountrangeuntil" & _
	    " AND authorizestructure_api.UserEntry(hdr.pocode, hdr.fcba, aus.authorizetemplatecode, hdr.fcentry) = 1" & _
		" AND at.document_classification = hdr.document_classification" & _
		" ORDER BY hdr.fcba, hdr.pocode"
	' response.write q
	Set rsPurchaseRequest = Server.CreateObject("ADODB.Recordset")
	rsPurchaseRequest.Open q, connstr
%>
<% if Request("oldui") <> "1" then %>
<div class="pra-head">
	<div>
		<div class="pra-eyebrow">Procurement</div>
		<h2>Persetujuan Purchase Order</h2>
		<p>PO berstatus <b>AuthorizeOnProgress</b> yang menunggu persetujuan Anda. Klik nomor PO untuk melihat detail, centang PO, lalu pilih Approve atau Reject.</p>
	</div>
</div>
<% end if %>
<form action="PurchaseOrderAuthorization.Asp" method="post" name="eForm" >
   <div class="pra-card">
   <table class="data pra-table" id="praTable" style="width:100%;">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="center">No</TD>
	<TD align="center">Kebun/Pabrik</TD>
	<TD align="center">PO Code</TD>
	<TD align="center">PO Date</TD>
	<TD align="center">Purchase By</TD>
	<TD align="center">Remarks</TD>
	<TD align="center">Amount</TD>
	<TD align="center" data-nosort title="Pilih semua PO yang tampil di halaman ini"><input type='checkbox' name='select-all' id='select-all' aria-label="Pilih semua PO yang tampil"/></TD>
	</tr>
	</thead>
	<tbody>
	
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
			<td align="left" title="<%=rsPurchaseRequest("pocode").Name%>" valign="top">
				<a href="PurchaseOrderDetailsApproval.asp?fcba=<%=rsPurchaseRequest("fcba")%>&fcCode=<%=rsPurchaseRequest("pocode")%>&authorizelevelcode=<%=rsPurchaseRequest("authorizelevelcode")%>&authorizetemplatecode=<%=rsPurchaseRequest("authorizetemplatecode")%>" target="_blank">
				<%=rsPurchaseRequest("pocode")%>
				</a>
			</td>
			<td align="left" title="<%=rsPurchaseRequest("podate").Name%>" valign="top"><%=rsPurchaseRequest("podate")%></td>
			<td align="left" title="<%=rsPurchaseRequest("purchaseby").Name%>" valign="top"><%=rsPurchaseRequest("purchaseby")%></td>
			<td align="left" title="<%=rsPurchaseRequest("remarks").Name%>" valign="top"><%=rsPurchaseRequest("remarks")%></td>
			<td align="right" title="<%=rsPurchaseRequest("grandtotal").Name%>" valign="top"><%=FormatNumber(rsPurchaseRequest("grandtotal"),2)%></td>
			<td valign="middle" align="center">
				<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);' aria-label="Pilih PO <%=rsPurchaseRequest("pocode")%>">
				<input type="hidden" name="fcba_<%=NO%>" value="<%=rsPurchaseRequest("fcba")%>">
				<input type="hidden" name="pocode_<%=NO%>" value="<%=rsPurchaseRequest("pocode")%>">
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
      response.Write"<tr><td colspan='"&d_span&"' align='left' class='pra-empty'>Data is empty</td></tr>"
   end if 
   rsPurchaseRequest.Close : set rsPurchaseRequest = Nothing
%>
	</tbody>
   </table>
   </div>
   <div class="pra-actions">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Hidden" Value="<%=no%>" Name="no">
		<span class="pra-selected" id="praSelected" aria-live="polite"></span>
		<button type="submit" class="button" name="txtJenis" value="ApprovalAction" onClick="return confirm('Anda yakin untuk melakukan persetujuan terhadap PO yang sudah dipilih?')"><i class="fa fa-check"></i> Approve</button>
		<button type="submit" class="button danger" name="txtJenis" value="RejectAction" onClick="return confirm('Anda yakin untuk mereject PO yang sudah dipilih?');"><i class="fa fa-times"></i> Reject</button>
   </div>
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
		if (window.praSync) praSync();
	}
	

</script>
<script type="text/javascript">	
		// "Pilih semua" hanya untuk baris yang sedang tampil (hasil cari / halaman aktif),
		// supaya PO yang tersembunyi tidak ikut ter-approve tanpa terlihat.
		$('#select-all').click( function () {
			var rows = $('#praTable tbody tr:visible .item');
			rows.prop('checked', this.checked);
			rows.val(this.checked ? 1 : 0);
			if (window.praSync) praSync();
		})
	</script>
<script type="text/javascript">
	// Tampilan: jumlah PO terpilih, tombol aktif bila ada pilihan, tabel cari/urut/halaman
	(function () {
		var table = document.getElementById('praTable');
		var info = document.getElementById('praSelected');
		var buttons = document.querySelectorAll('.pra-actions button[type="submit"]');
		if (!table) return;
		window.praSync = function () {
			var items = table.querySelectorAll('tbody .item');
			var n = 0, visible = 0, visibleChecked = 0;
			for (var i = 0; i < items.length; i++) {
				var shown = items[i].offsetParent !== null;
				if (items[i].checked) n++;
				if (shown) { visible++; if (items[i].checked) visibleChecked++; }
				var tr = items[i].closest ? items[i].closest('tr') : null;
				if (tr) tr.classList.toggle('pra-checked', items[i].checked);
			}
			if (info) info.innerHTML = items.length ? ('<b>' + n + '</b> dari ' + items.length + ' PO dipilih') : 'Tidak ada PO yang menunggu persetujuan';
			for (var b = 0; b < buttons.length; b++) buttons[b].disabled = (n === 0);
			var all = document.getElementById('select-all');
			if (all) { all.checked = visible > 0 && visibleChecked === visible; all.indeterminate = visibleChecked > 0 && visibleChecked < visible; }
		};
		if (window.FbpTable) FbpTable.init(table, { pageSize: 10, onChange: function () { praSync(); } });
		praSync();
	})();
</script>
</body>
</html>
