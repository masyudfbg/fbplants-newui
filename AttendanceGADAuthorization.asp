<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Attendance & Activity Approval</title>
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
 '--- add Aam Feb 21
	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)


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
                q = " INSERT INTO attendance_gad_hist ( fddate" & _
                    "                                           , gangcode" & _
                    "                                           , currentauthorizelevel" & _
                    "                                           , fcentry" & _
                    "                                           , fcedit" & _
                    "                                           , fcip" & _
                    "                                           , fcba" & _
                    "                                           , lastupdate" & _
                    "                                           , lasttime)"& _
                    " SELECT fddate, gangcode" & _
                    "        , '" & request("authorizelevelcode_"&TRIM(n)) & "' currentauthorizelevel" & _
                    "        , '" & Session("CodeName") & "' fcentry" & _
                    "        , '" & Session("CodeName") & "' fcedit" & _
                    "        , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
                    "        , fcba" & _
                    "        , SYSDATE lastupdate" & _
                    "        , TO_CHAR(SYSDATE, 'hh24:mi') lasttime" & _
                    " FROM attendance_gad_temp" & _
                    " WHERE fddate = to_date('" & request("fddate_"&TRIM(n)) & "','mm/dd/yyyy')" & _
                    "       AND gangcode = '"& request("gangcode_"&TRIM(n)) & "'" & _
                    "       AND fcba = '"& Session("fcba") & "'"
                'response.write q & "<br>"
                connStr.execute(q)

                is_final = 0
                q = " SELECT COUNT(1) total FROM authorizelevel al" & _
                    " INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
                    " INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
                    " WHERE at.document_type = 'AA'" & _
                    "       AND DECODE(at.document_type, 'PR', at.document_classification, '0') = '0'" & _
                    "       AND al.isactive = 'TRUE'" & _
                    "       AND al.finallevel = 'TRUE'" & _
                    "       AND al.fccode = '" & request("authorizelevelcode_"&TRIM(n)) & "'" & _
                    "       AND aus.authorizetemplatecode = '" & request("authorizetemplatecode_"&TRIM(n)) & "'" & _
                    "       AND aus.fcba = '" & Session("fcba") & "'"
                ' response.write q
                set rsPurchase = server.CreateObject("Adodb.recordset")
                rsPurchase.Open q, connstr
                IF not rsPurchase.eof then
                    is_final = cdbl(to_nol(rsPurchase("total")))
                end if
                
				'response.write is_final 
                if is_final > 0 then

                                q = " INSERT INTO attendance_gad(GANGCODE" & _
                                    "                             , FDDATE" & _
                                    "                             , SUPERVISION_1" & _
                                    "                             , SUPERVISION_2" & _
                                    "                             , SUPERVISION_3" & _
                                    "                             , SUPERVISION_4" & _
                                    "                             , EMPLOYEECODE" & _
                                    "                             , ATTENDANCE" & _
                                    "                             , JOBCODE" & _
                                    "                             , LOCATIONTYPE" & _
                                    "                             , LOCATIONCODE" & _
                                    "                             , CHARGEJOB" & _
                                    "                             , CHARGETYPE" & _
                                    "                             , CHARGECODE" & _
                                    "                             , MANDAYS" & _
                                    "                             , OTHRS" & _
                                    "                             , RATE" & _
                                    "                             , UNIT" & _
                                    "                             , OUTPUT" & _
                                    "                             , REFERENCE" & _
                                    "                             , REMARKS" & _
                                    "                             , BUCKET" & _
                                    "                             , FCENTRY" & _
                                    "                             , FCEDIT" & _
                                    "                             , FCIP" & _
                                    "                             , FCBA" & _
                                    "                             , LASTUPDATE" & _
                                    "                             , LASTTIME" & _
                                    "                             , linenokey" & _
                                    "                             , RowState" & _
                                    "                             , document_classification" & _
                                    "                             , TYPE_OVERTIME" & _
                                    "                             , OVERTIME_HOURS" & _
                                    "                             , KG_BRONDOLAN" & _
                                    "                             , SPBNO" & _
                                    "                             , KG_JANJANG" & _
									"                             , ABW" & _
									"                             , BUNCH_WEIGHT" & _
									"                             )" & _
                                    "                       SELECT GANGCODE" & _
                                    "                             , FDDATE" & _
                                    "                             , SUPERVISION_1" & _
                                    "                             , SUPERVISION_2" & _
                                    "                             , SUPERVISION_3" & _
                                    "                             , SUPERVISION_4" & _
                                    "                             , EMPLOYEECODE" & _
                                    "                             , ATTENDANCE" & _
                                    "                             , JOBCODE" & _
                                    "                             , LOCATIONTYPE" & _
                                    "                             , LOCATIONCODE" & _
                                    "                             , CHARGEJOB" & _
                                    "                             , CHARGETYPE" & _
                                    "                             , CHARGECODE" & _
                                    "                             , MANDAYS" & _
                                    "                             , OTHRS" & _
                                    "                             , RATE" & _
                                    "                             , UNIT" & _
                                    "                             , OUTPUT" & _
                                    "                             , REFERENCE" & _
                                    "                             , REMARKS" & _
                                    "                             , BUCKET" & _
                                    "                             , fcentry" & _
                                    "                             , '" & Session("CodeName") & "' fcedit" & _
                                    "                             , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
                                    "                             , FCBA" & _
                                    "                             , SYSDATE lastupdate" & _
                                    "                             , TO_CHAR(SYSDATE, 'hh24:mi') lasttime" & _
                                    "                             , linenokey" & _
                                    "                             , RowState" & _
                                    "                             , document_classification" & _
                                    "                             , TYPE_OVERTIME" & _
                                    "                             , OVERTIME_HOURS" & _
                                    "                             , KG_BRONDOLAN" & _
                                    "                             , SPBNO" & _
                                    "                             , KG_JANJANG" & _
									"                             , ABW" & _
									"                             , BUNCH_WEIGHT " & _
                                    "                   FROM attendance_gad_temp" & _
                                    "                   WHERE fddate = to_date('" & request("fddate_"&TRIM(n)) & "','mm/dd/yyyy')" & _
                                    "                   AND gangcode = '"& request("gangcode_"&TRIM(n)) & "'" & _
                                    "                   AND fcba = '"& Session("fcba") & "'"
                               
                                connStr.execute(q)
   
                    strQueryStatus = "BEGIN ATTENDANCEGAD_API.NextState('" & Session("fcba") & "',to_date('" & request("fddate_"&TRIM(n)) & "','mm/dd/yyyy'),'"& request("gangcode_"&TRIM(n)) &"','Approve' ); END; "
                    connStr.execute(strQueryStatus)
                    
                    strQueryStatus = " UPDATE attendance_gad SET rowstate = 'Approved' WHERE fcba = '" & Session("fcba") & "'" & _
                                 " AND fddate = to_date('" & request("fddate_"&TRIM(n)) & "','mm/dd/yyyy') AND gangcode = '"& request("gangcode_"&TRIM(n)) &"'"
                    connStr.execute(strQueryStatus)

                    strQueryStatus = " UPDATE attendance_gad_temp SET rowstate = 'Approved' WHERE fcba = '" & Session("fcba") & "'" & _
                                 " AND fddate = to_date('" & request("fddate_"&TRIM(n)) & "','mm/dd/yyyy') AND gangcode = '"& request("gangcode_"&TRIM(n)) &"'"
                    connStr.execute(strQueryStatus)
                End If
            End if
	   next
		
	   cmdView	= "True"
	elseif Request("txtJenis")= "RejectAction" then
		For n=1 to (cdbl(request("no"))-1)
			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
			    strQueryStatus = " UPDATE attendance_gad SET rowstate = 'Reject' WHERE fcba = '" & Session("fcba") & "'" & _
                                 " AND fddate = to_date('" & request("fddate_"&TRIM(n)) & "') AND gangcode = '"& request("gangcode_"&TRIM(n)) &"'"
                connStr.execute(strQueryStatus)

                strQueryStatus = " UPDATE attendance_gad_temp SET rowstate = 'Reject' WHERE fcba = '" & Session("fcba") & "'" & _
                                 " AND fddate = to_date('" & request("fddate_"&TRIM(n)) & "') AND gangcode = '"& request("gangcode_"&TRIM(n)) &"'"
                connStr.execute(strQueryStatus)
				
				strQueryStatus = "BEGIN ATTENDANCEGAD_API.NextState('" & Session("fcba") & "',to_date('" & request("fddate_"&TRIM(n)) & "','mm/dd/yyyy'),'"& request("gangcode_"&TRIM(n)) &"','Reject' ); END; "
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
	
		' Rev Query Select Added By Gibran
			q = "   SELECT stage1.businessunit fcba," & _
				"          stage1.gangcode," & _
				"          stage1.fddate," & _
				"          stage1.fccode authorizetemplatecode," & _
				"          stage1.document_classification," & _
				"          stage1.currentauthorizelevel," & _
				"          stage1.authorizelevelname," & _
				"          stage1.lastauthorizer," & _
				"          stage1.lastactiondate," & _
				"          stage1.authorizelevelcode," & _
				"          aust.fcusercode," & _
				"          stage1.rowstate" & _
				"     FROM (SELECT abc.businessunit," & _
				"                  abc.gangcode," & _
				"                  abc.fddate," & _
				"                  at.fccode," & _
				"                  abc.document_classification," & _
				"                  abc.currentauthorizelevel," & _
				"                  abc.authorizelevelname," & _
				"                  abc.lastauthorizer," & _
				"                  abc.lastactiondate," & _
				"                  NVL ( Min(aus.authorizelevelcode),0) authorizelevelcode ," & _
				"                  rowstate" & _
				"             FROM (  SELECT temp.fcba AS businessunit," & _
				"                            temp.gangcode AS gangcode," & _
				"                            temp.fddate AS fddate," & _
				"                            temp.document_classification," & _
				"                            authorizelevel.fccode AS currentauthorizelevel," & _
				"                            authorizelevel.fcname AS authorizelevelname," & _
				"                            Temp2.fcentry AS lastauthorizer," & _
				"                            Temp2.lastupdate AS lastactiondate," & _
				"                            temp.rowstate" & _
				"                       FROM " & _
				"                            (SELECT DISTINCT " & _
                "                                    attendance_gad_hist.fcba," & _
                "                                    attendance_gad_hist.fddate," & _
                "                                    attendance_gad_hist.gangcode," & _
                "                                    attendance_gad_hist.fcentry," & _
                "                                    attendance_gad_hist.lastupdate ," & _
                "                                    attendance_gad_hist.currentauthorizelevel" & _
                "                             FROM attendance_gad_hist " & _
                "                             WHERE attendance_gad_hist.fcba = '"& session("fcba") &"'" & _
				"                                    AND attendance_gad_hist.fddate >= TO_DATE ('"& v_txtdatefrom &"', 'mm/dd/yyyy') " & _
				"                  	                 AND attendance_gad_hist.fddate <= TO_DATE ('"& v_txtdateuntil &"', 'mm/dd/yyyy') " & _
                "                             )Temp2 " & _
				"                            INNER JOIN" & _
				"                            (SELECT attendance_gad_temp.fcba," & _
				"                                    attendance_gad_temp.gangcode," & _
				"                                    attendance_gad_temp.fddate," & _
				"                                    attendance_gad_temp.rowstate," & _
				"                                    attendance_gad_temp.document_classification" & _
				"                              FROM attendance_gad_temp" & _
				"                              WHERE attendance_gad_temp.fcba = '"& session("fcba") &"'" & _
				"                                    AND attendance_gad_temp.fddate >= TO_DATE ('"& v_txtdatefrom &"', 'mm/dd/yyyy') " & _
				"                  					 AND attendance_gad_temp.fddate <= TO_DATE ('"& v_txtdateuntil &"', 'mm/dd/yyyy') "  & _
				"                                    AND  rowstate <>   'Approved' "  & _
				"                            ) temp" & _
				"                               ON (    Temp2.fcba = temp.fcba" & _
				"                                   AND Temp2.gangcode = " & _
				"                                          temp.gangcode" & _
				"                                   AND Temp2.fddate = temp.fddate)" & _
				"                            INNER JOIN authorizelevel" & _
				"                               ON Temp2.currentauthorizelevel =" & _
				"                                     authorizelevel.fccode" & _
				"                            INNER JOIN" & _
				"                            (  SELECT attendance_gad_hist.fcba," & _
				"                                      attendance_gad_hist.gangcode," & _
				"                                      attendance_gad_hist.fddate," & _
				"                                      MAX (" & _
				"                                         attendance_gad_hist.currentauthorizelevel)" & _
				"                                         alevel" & _
				"                               FROM attendance_gad_hist" & _
				"                               WHERE attendance_gad_hist.fcba = '"& session("fcba") &"'" & _
				"                                    AND attendance_gad_hist.fddate >= TO_DATE ('"& v_txtdatefrom &"', 'mm/dd/yyyy') " & _
				"                  					 AND attendance_gad_hist.fddate <= TO_DATE ('"& v_txtdateuntil &"', 'mm/dd/yyyy') "  & _
				"                               GROUP BY attendance_gad_hist.fcba," & _
				"                                      attendance_gad_hist.gangcode," & _
				"                                      attendance_gad_hist.fddate) aa" & _
				"                               ON (    Temp2.fcba = aa.fcba" & _
				"                                   AND Temp2.currentauthorizelevel =" & _
				"                                          aa.alevel" & _
				"                                   AND Temp2.fddate = aa.fddate" & _
				"                                   AND Temp2.gangcode = aa.gangcode)" & _
				"                            GROUP BY temp.fcba," & _
				"                            temp.gangcode," & _
				"                            temp.fddate," & _
				"                            temp.document_classification," & _
				"                            authorizelevel.fccode," & _
				"                            authorizelevel.fcname," & _
				"                            Temp2.fcentry," & _
				"                            Temp2.lastupdate," & _
				"                            temp.rowstate) abc" & _
				"                  LEFT JOIN authorizetemplate at" & _
				"                     ON abc.document_classification = at.document_classification" & _
				"                  LEFT JOIN authorizestructure aus" & _
				"                     ON (    at.fccode = aus.authorizetemplatecode" & _
				"                         AND abc.businessunit = aus.fcba" & _
				"                          AND aus.authorizelevelcode > abc.currentauthorizelevel" & _
				"                        ) " & _
				"                         GROUP BY abc.businessunit," & _
				"                                          abc.gangcode," & _
				"                                          abc.fddate," & _
				"                                          at.fccode," & _
				"                                          abc.document_classification," & _
				"                                          abc.currentauthorizelevel," & _
				"                                          abc.authorizelevelname," & _
				"                                          abc.lastauthorizer," & _
				"                                          abc.lastactiondate," & _
                "                                          rowstate" & _
				"                        ) stage1" & _
				"          INNER JOIN authorizestructure aust" & _
				"             ON (    stage1.businessunit = aust.fcba" & _
				"                 AND stage1.authorizelevelcode = aust.authorizelevelcode" & _
				"                 AND stage1.fccode = aust.authorizetemplatecode )" & _
				"    WHERE     stage1.rowstate = 'AuthorizeOnProgress'" & _
				"          AND stage1.businessunit = '"& session("fcba") &"'" & _
				"          AND aust.fcusercode = '"& session("CodeName") &"'" & _
				" GROUP BY stage1.businessunit," & _
				"          stage1.gangcode," & _
				"          stage1.fddate," & _
				"          stage1.fccode," & _
				"          stage1.document_classification," & _
				"          stage1.currentauthorizelevel," & _
				"          stage1.authorizelevelname," & _
				"          stage1.lastauthorizer," & _
				"          stage1.lastactiondate," & _
				"          stage1.authorizelevelcode," & _
				"          aust.fcusercode," & _
				"          stage1.rowstate" & _
				" ORDER BY stage1.FDDATE, stage1.GANGCODE" 
		' End Added
	' response.write q
	Set rsPurchaseRequest = Server.CreateObject("ADODB.Recordset")
	rsPurchaseRequest.Open q, connstr
%>
<form action="AttendanceGADAuthorization.asp" method="post" name="eForm" >
   <table class="data" style="width:100%;">
	<tr>
	    <td><font face="Verdana" size="1">From</font></td>
	    <td>
			<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
	    </td>
	    <td><font face="Verdana" size="1">To</font></td>
	    <td>
	        <input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
	    </td>
		<td>
			<input type="submit" value="View" />
		</td>
	</tr>

	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="center">No</TD>
	<TD align="center">Kebun/Pabrik</TD>
	<TD align="center">Date</TD>
    <TD align="center">Gang Code</TD> 
    <TD align="center">Document Classification</TD> 
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
            <td align="center" title="<%=rsPurchaseRequest("fddate").Name%>" valign="top"><%=rsPurchaseRequest("fddate")%></td>
			<td align="center" title="<%=rsPurchaseRequest("gangcode").Name%>" valign="top">
				<a href="AttendanceGADApproval.asp?fcba=<%=rsPurchaseRequest("fcba")%>&fddate=<%=rsPurchaseRequest("fddate")%>&gangcode=<%=rsPurchaseRequest("gangcode")%>&authorizelevelcode=<%=rsPurchaseRequest("authorizelevelcode")%>&authorizetemplatecode=<%=rsPurchaseRequest("authorizetemplatecode")%>" target="_blank">
				<%=rsPurchaseRequest("gangcode")%>
				</a>
			</td>
            <td align="center" title="<%=rsPurchaseRequest("document_classification").Name%>" valign="top"><%=rsPurchaseRequest("document_classification")%></td>
			<td valign="middle" align="center">
				<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);'>
				<input type="hidden" name="fcba_<%=NO%>" value="<%=rsPurchaseRequest("fcba")%>">
				<input type="hidden" name="gangcode_<%=NO%>" value="<%=rsPurchaseRequest("gangcode")%>">
                <input type="hidden" name="fddate_<%=NO%>" value="<%=rsPurchaseRequest("fddate")%>">
                <input type="hidden" name="document_classification_<%=NO%>" value="<%=rsPurchaseRequest("document_classification")%>">
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
      response.Write"<tr><td colspan='"&d_span&"' align='left'>Data is empty</td></tr>"
   end if 
   rsPurchaseRequest.Close : set rsPurchaseRequest = Nothing
%>
		<tr>
			<td colspan="<%=d_span%>" align="left">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Hidden" Value="<%=no%>" Name="no">
				<button type="submit" class="button" name="txtJenis" value="ApprovalAction" onClick="return confirm('Anda yakin untuk melakukan persetujuan terhadap Attendance & Activity yang sudah dipilih?')"><i class="fa fa-check"></i> Approve</button>
				<button type="submit" class="button danger" name="txtJenis" value="RejectAction" onClick="return confirm('Anda yakin untuk mereject Attendance & Activity yang sudah dipilih?');"><i class="fa fa-times"></i> Reject</button>
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
