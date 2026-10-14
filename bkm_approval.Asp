<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<html>
<head>
<title>FBPlants - BKM Approval</title>
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
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014a"></script>
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
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,rsbkm,statusapproved
 DIM bkmno,txtfcba,txtafdeling,txtstatus,tipebkm,rsDocument,ScriptName
 ScriptName	=	Request.ServerVariables("SCRIPT_NAME")


d_span = 14
 


 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")
 statusapproved =Request("statusapproved")
 txtfcba =Request("businessunit")
 txtafdeling =Request("sectionname")
 txtstatus =Request("statusapproved")
 'response.write  txtstatus

 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 bkmno = Request("bkmno")
 
 '---- Data Header
 if Request("PRDATE")<>"" then
   PRDATE   =Request("PRDATE")
 else
   PRDATE   =textDate(date)
 end if



 '-------------ISI Payment Form
cmdView = "True"
if cmdView = "True" Then
    dim bkmfrom,bkmto,txtfcba1,txtafd,txtstatus1,filter,document_clasification
    bkmfrom     = Request("bkmdate_From")
    bkmto       = Request("bkmdate_To")
    txtfcba1    = Request("businessunit")
    txtafd      = Request("sectionname")
    txtstatus1  = Request("statusapproved")
    

    'query dulu ke tabel authorizestructure untuk mendapatkan ID document_classification yang 504 (WIP)
    q = "SELECT distinct at.document_classification " & _
    " FROM authorizetemplate at " & _
    " INNER JOIN  authorizestructure aus ON (aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & Session("CodeName") & "') AND at.document_classification in ('504','505')"
    Set rsDocument = Server.CreateObject("ADODB.Recordset")
	rsDocument.Open q, connstr
    IF not rsDocument.EOF Then
        document_clasification = rsDocument("document_classification")
    end If
   ' Response.Write q
   ' Response.Write document_clasification

    if (document_clasification <> "504" and document_clasification <> "505")   then
        if trim(txtafd) = "" then 
            filter = " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & Session("CodeName") & "' ) "
        else
            filter = " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & Session("CodeName") & "' AND sectionname='" & txtafd & "') "
        end if
    end if

	q = " SELECT hdr.fcba, hdr.bkmno,hdr.rkbno, to_char(hdr.bkmdate,'dd/mm/yyyy')bkmdate, hdr.document_classification, hdr.sectioncode, hdr.gangcode,  hdr.document_classification," & _
		" hdr.VEHICLECODE,hdr.fcentry,hdr.fcedit,aus.authorizelevelcode,hdr.rowstate, hdr.contractorcode " & _
	    " FROM bkmhdr hdr" & _
	    " LEFT JOIN authorizetemplate at ON (at.document_type = 'AA' AND at.document_classification = hdr.document_classification)" & _
		" INNER JOIN authorizestructure aus ON (aus.fcba = hdr.fcba AND aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & Session("CodeName") & "')" & _
	    " WHERE hdr.bkmdate between to_date('" & bkmfrom & "','mm/dd/yyyy') AND to_date('" & bkmto & "','mm/dd/yyyy') AND hdr.rowstate = '" & txtstatus1 & "'" & _
		" AND at.document_classification = hdr.document_classification AND hdr.fcba='" & txtfcba & "' AND authorizelevelcode = 900" & _
        filter & _ 
		" ORDER BY hdr.bkmdate,hdr.fcba, hdr.bkmno"
	'response.write q
	Set rsbkm = Server.CreateObject("ADODB.Recordset")
	rsbkm.Open q, connstr
%>
<form action="bkm_approval.Asp" method="post" name="eForm" >
   <table class="data" style="width:100%;">
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
        <TD align="center">No</TD>
        <TD align="center">Kebun/Pabrik</TD>
        <TD align="center">BKM No</TD>
        <TD align="center">BKM RKB</TD>
        <TD align="center">BKM Date</TD>
        <TD align="center">BKM Type</TD>
        <TD align="center">Section</TD>
        <TD align="center">Gang</TD>
        <TD align="center">Vehcode</TD>
        <TD align="center">Contractor</TD>
        <TD align="center">Status</TD>
        <TD align="center">Action</TD>
	</tr> 
	
   <%
   dim tipebkmname
   if Not rsbkm.eof then
      nCountofField=rsbkm.fields.count-1
	  no=1
	  While Not rsbkm.EOF
        tipebkm = rsbkm("document_classification")
        if tipebkm = "501" then 
            tipebkmname="BKM Perawatan"
        end if
        if tipebkm = "502" then 
            tipebkmname="BKM Traksi"
        end if
        if tipebkm = "503" then 
            tipebkmname="BKM Panen"
        end if
        if tipebkm = "504" then 
            tipebkmname="BKM WIP"
        end if
		if tipebkm = "505" then 
            tipebkmname="BKM Adjusment"
        end if
		
		if no  mod 2 = 1 then
			cColor = "#EEEEEE"
		else
			cColor = "#E5E5E5"
		end if%>
		<tr bgcolor="<%=cColor%>">

			<td align="center" valign="top"><%=NO%></td>
			
			<td align="center" title="<%=rsbkm("fcba").Name%>" valign="top"><%=rsbkm("fcba")%></td>
			<td align="center" title="<%=rsbkm("bkmno").Name%>" valign="top">
            <%
                if tipebkm = "501" then %>
                    <a href="bkm_operasional_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                    <%=rsbkm("bkmno")%>
                    </a> 
                <%elseif tipebkm = "502" then%>
                    <a href="bkm_traksi_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                    <%=rsbkm("bkmno")%>
                    </a> 
                 <%elseif tipebkm = "503" then%>
                    <a href="bkm_panen_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                    <%=rsbkm("bkmno")%>
                 <%elseif tipebkm = "504" then%>
                    <a href="bkm_wip_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                    <%=rsbkm("bkmno")%>
                    </a> 
				 <%elseif tipebkm = "505" then%>
                    <a href="bkm_adjusment_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                    <%=rsbkm("bkmno")%>
                    </a> 
                <%end if%>
        
			</td>
			<td align="left" title="<%=rsbkm("rkbno").Name%>" valign="top"><%=rsbkm("rkbno")%></td>
			<td align="center" title="<%=rsbkm("bkmdate").Name%>" valign="top"><%=rsbkm("bkmdate")%></td>
			<td align="left" title="<%=rsbkm("document_classification").Name%>" valign="top"><%=tipebkmname%></td>
			<td align="left" title="<%=rsbkm("sectioncode").Name%>" valign="top"><%=rsbkm("sectioncode")%></td>
			<td align="left" title="<%=rsbkm("gangcode").Name%>" valign="top"><%=rsbkm("gangcode")%></td>			
			<td align="left" title="<%=rsbkm("VEHICLECODE").Name%>" valign="top"><%=rsbkm("VEHICLECODE")%></td>			
			<td align="left" title="<%=rsbkm("contractorcode").Name%>" valign="top"><%=rsbkm("contractorcode")%></td>			
			<td align="left" title="<%=rsbkm("rowstate").Name%>" valign="top"><%=rsbkm("rowstate")%></td>			
			<td align="center">
                <%
                if tipebkm = "501" then %>
                    <a href="bkm_operasional_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                    <img src="IMAGES/view.png" alt="View BKM="0">
                    </a> 
                <%elseif tipebkm = "502" then%>
                    <a href="bkm_traksi_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                     <img src="IMAGES/view.png" alt="View BKM="0">
                    </a> 
                 <%elseif tipebkm = "503" then%>
                    <a href="bkm_panen_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                     <img src="IMAGES/view.png" alt="View BKM="0">
                 <%elseif tipebkm = "504" then%>
                    <a href="bkm_wip_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                    <img src="IMAGES/view.png" alt="View BKM="0">
                    </a> 
				 <%elseif tipebkm = "505" then%>
                    <a href="bkm_adjusment_view.asp?bkmno=<%=rsbkm("bkmno")%>" target="_blank">
                     <img src="IMAGES/view.png" alt="View BKM="0">
                    </a> 
                <%end if%>
                
            </td>
			
		  </tr>
		  <%
		  rsbkm.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write"<tr><td colspan='"&d_span&"' align='left'>Data is empty</td></tr>"
   end if 
   rsbkm.Close : set rsbkm = Nothing
%>
		<tr>
			<td colspan="<%=d_span%>" align="left">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Hidden" Value="<%=no%>" Name="no">
				<button type="submit" class="button danger" name="tutup" value="RejectAction" onClick="window.close()"><i class="fa fa-times"></i> Close</button>
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
