<%Option Explicit%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<%
' Belum login: langsung ke halaman login (logout.asp), jangan tampilkan dashboard dengan header "Please login first..."
If Session("USERNAME") & "" = "" Then Response.Redirect "logout.asp"
%>
<html>
<head>
<meta name="robots" content="noindex, nofollow">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="icon" href="/IMAGES/logofbg.jpeg" type="image/jpeg"/>
<title>Dashboard &middot; FB-Plants</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<script type="text/javascript" src="dtree.js"></script>
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
<style>
/* Khusus dashboard (kerangka halaman ada di css/fbp-shell.css) */
body.dash .judul_head { display: none; }
body.dash .dash-welcome {
	display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 16px;
	padding: 22px 26px;
	margin-bottom: 24px;
	border-radius: 16px;
	color: #fff;
	background: linear-gradient(120deg, var(--brand) 0%, var(--brand-dark) 100%);
	position: relative; overflow: hidden;
}
body.dash .dash-welcome::after {
	content: ""; position: absolute; right: -40px; top: -60px;
	width: 220px; height: 220px; border-radius: 40px;
	background: rgba(255, 255, 255, .06);
	transform: rotate(45deg);
}
body.dash .dash-welcome .eyebrow { font-size: 12px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; color: var(--accent); }
body.dash .dash-welcome h1 { margin: 4px 0 0; font-family: inherit; font-size: 22px; font-weight: 700; color: #fff; }
body.dash .dash-welcome .meta { display: flex; flex-wrap: wrap; gap: 8px; position: relative; z-index: 1; }
body.dash .dash-welcome .chip {
	display: inline-flex; align-items: center; gap: 8px;
	padding: 7px 12px;
	font-size: 13px;
	background: rgba(255, 255, 255, .12);
	border: 1px solid rgba(255, 255, 255, .18);
	border-radius: 999px;
}
body.dash .dash-welcome .chip .fa { color: var(--accent); }
/* Chip BA = tombol dropdown ganti BA (js/fbp-baswitch.js, ba_switch.asp) */
body.dash .dash-welcome button.ba-switch { font: inherit; font-size: 13px; color: #fff; cursor: pointer; transition: background .15s, border-color .15s; }
body.dash .dash-welcome button.ba-switch:hover, body.dash .dash-welcome button.ba-switch[aria-expanded="true"] { background: rgba(255, 255, 255, .22); border-color: rgba(255, 255, 255, .38); }
body.dash .dash-welcome button.ba-switch:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
body.dash .dash-welcome button.ba-switch .fa-caret-down { margin-left: 2px; color: #fff; opacity: .8; }
.ba-panel { position: fixed; z-index: 2000; display: none; flex-direction: column; overflow: hidden; background: #fff; color: #17262a;
	border: 1px solid #dfe6e4; border-radius: 14px; box-shadow: 0 18px 48px rgba(23, 38, 42, .22); font-size: 13px; }
.ba-panel.open { display: flex; }
.ba-panel.busy { opacity: .7; pointer-events: none; }
.ba-panel-head { padding: 14px 16px 8px; }
.ba-panel-head b { display: block; font-size: 14px; }
.ba-panel-head span { display: block; margin-top: 2px; font-size: 12px; color: #5d6d69; }
.ba-panel-search { margin: 4px 12px 8px; height: 36px; padding: 4px 12px; font: inherit; font-size: 13px; border: 1px solid #d5dfdc; border-radius: 9px; outline: 0; }
.ba-panel-search:focus { border-color: #1B7863; box-shadow: 0 0 0 3px rgba(27, 120, 99, .18); }
.ba-panel-list { flex: 1 1 auto; min-height: 60px; overflow-y: auto; padding: 0 6px 8px; border-top: 1px solid #edf1f0; }
.ba-group { position: sticky; top: 0; z-index: 1; display: flex; gap: 8px; align-items: baseline; padding: 10px 10px 4px; background: #fff; font-size: 11px; letter-spacing: .3px; color: #5d6d69; text-transform: uppercase; }
.ba-group b { color: #12584a; }
.ba-group span { overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
.ba-item { display: flex; align-items: center; gap: 10px; padding: 8px 10px; border-radius: 8px; cursor: pointer; }
/* hasil filter: display:flex di atas menimpa atribut hidden, jadi item yang tidak cocok dulu tetap tampil */
.ba-item[hidden], .ba-group[hidden] { display: none; }
.ba-item:hover, .ba-item.active { background: #eef6f3; }
.ba-item.current { cursor: default; background: #e2f4ec; }
.ba-code { flex: 0 0 auto; min-width: 44px; padding: 1px 7px; font-size: 11.5px; font-weight: 700; text-align: center; color: #12584a; background: #e7f3ef; border-radius: 6px; }
.ba-name { flex: 1 1 auto; min-width: 0; }
.ba-tag { flex: 0 0 auto; padding: 1px 8px; font-size: 11px; font-weight: 600; color: #fff; background: #1B7863; border-radius: 999px; }
.ba-panel-msg { padding: 18px 12px; text-align: center; color: #5d6d69; }
</style>
</head>
<body<% if Session("fcBa") <> "" then %> class="dash"<% end if %>>
<%if Session("USERNAME")="" then%>
	      <!--#include file="toplogin.jsp"-->
<%else%>
	     <!--#include file="top.jsp"-->
<%end if%>


 <% if Session("fcBa") <> "" then %>
    <div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
      <!--#include file="menu.asp"-->
    </div>
    <div class="judul_head"  valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
    <div class="judul">
	    FB-Plants ( First Borneo - Plantation System)
	  </div>
    </div>

 <% end if %>
    <div style="overflow-y: scroll;">
      <div class="box_both" style="
        <% if Session("fcBa") <> "" then %>
        position:fixed;
        <% else%>
        position:static;
        margin:auto;
        <% end if %>
        " >

        <% if Session("fcBa") <> "" then %>
        <div class="dash-welcome">
          <div style="position: relative; z-index: 1;">
            <div class="eyebrow">FB-Plants &middot; First Borneo Plantation System</div>
            <h1>Selamat datang, <%=Server.HTMLEncode(Session("FULLNAME") & "")%></h1>
          </div>
          <div class="meta">
            <%
            ' token untuk POST ganti BA (ba_switch.asp)
            If Session("ba_tok") & "" = "" Then
              Randomize
              Session("ba_tok") = Hex(Int(Rnd * 2147483647)) & Hex(Int(Rnd * 2147483647)) & Hex(Int(Timer * 1000))
            End If
            %>
            <button type="button" class="chip ba-switch" aria-haspopup="dialog" aria-expanded="false" title="Ganti BA tanpa logout" data-tok="<%=Server.HTMLEncode(Session("ba_tok") & "")%>"><i class="fa fa-briefcase" aria-hidden="true"></i> BA: <%=Server.HTMLEncode(Session("BA") & "")%> (<%=Server.HTMLEncode(Session("fcBA") & "")%>) <i class="fa fa-caret-down" aria-hidden="true"></i></button>
            <span class="chip"><i class="fa fa-calendar" aria-hidden="true"></i> <%=FormatDateTime(Cdate(Now),1)%></span>
          </div>
        </div>
        <% end if %>

        <%
          dim userlogin
          userlogin = Session("CodeName")
          'fcBA  = Session("fcBA")
        ' IF userlogin="DIV_HEAD_FAC" or userlogin="IT_HO1" or userlogin="IT_HO2" or userlogin="DIV_BRA" then %>
            <div style="display: block">
              <!--#include file="dashboard_acct_closing.asp"-->
            </div>
            <div>
              <!--#include file="slideMonitoringDataInput.asp"-->
            </div>
        <% ' end if
        %>
      </div>
    </div>
  <!--#include file="bottom.jsp"-->
<% if Session("fcBa") <> "" then %>
<script type="text/javascript" src="js/fbp-baswitch.js?v=20261010a"></script>
<% end if %>
</body>
</html>
