<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<html>
	<head>
		
		<title>ePlaS - Closing Harian Produksi</title>
		<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>

 <script>
$(function() {
$( ".datepicker" ).datepicker();
});

</script>

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

function getData()
{
	var tglskrg = $("#txttgl").val();
	window.location.href = "lhp_closing_produksi.Asp?txtJenis=BUKA&tglskrg=" + tglskrg;
}
</script>


<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261015a"></script>
<% end if %>
</head>
<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiriisi{
		position: relative;
		left: 0px;
	}
	.lebarkolomsatu{
		width: 10px;
	}
	.lebarkolomdua{
		width: 10px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}
</style>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
<!--#include file="top.jsp"-->

<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
<div class="judul">
	   Module : Closing Harian Produksi
</div>
<div class="box_both" style="overflow:scroll;">

<%

Dim sqlsimpan,sqlskrg,sqlberikut,rsskrg,rsberikut, txtJenis,tglsekarangtxt,tglberikuttxt,tglskrgtxt

txtJenis 	= request("txtJenis")
tglskrgtxt 	= request("tglskrg")

'Response.Write txtJenis


sqlskrg = "select PRODUCTIONMILLDATE from configmst"
Set rsskrg = Server.CreateObject("ADODB.Recordset")
rsskrg.CursorType = 3
rsskrg.CursorLocation = 3
rsskrg.Open sqlskrg, connstr
if not rsskrg.eof then
	tglsekarangtxt 		= rsskrg("PRODUCTIONMILLDATE")
end if
rsskrg.Close
set rsskrg = Nothing

sqlberikut = "select PRODUCTIONMILLDATE+1 Berikut from configmst"
Set rsberikut = Server.CreateObject("ADODB.Recordset")
rsberikut.CursorType = 3
rsberikut.CursorLocation = 3
rsberikut.Open sqlberikut, connstr
if not rsberikut.eof then
	tglberikuttxt 		= rsberikut("Berikut")
end if
rsberikut.Close
set rsberikut = Nothing

if Request("txtJenis")="TUTUP" then
	sqlsimpan = "update CONFIGMST set PRODUCTIONMILLDATE=to_date('" & tglsekarangtxt & "','mm/dd/yyyy')+1"
	connStr.execute(sqlsimpan) 

	sqlskrg = "select PRODUCTIONMILLDATE,PRODUCTIONMILLDATE+1 NEXTDATE from configmst"
	Set rsskrg = Server.CreateObject("ADODB.Recordset")
	rsskrg.CursorType = 3
	rsskrg.CursorLocation = 3
	rsskrg.Open sqlskrg, connstr
	if not rsskrg.eof then
		tglsekarangtxt 		= rsskrg("PRODUCTIONMILLDATE")
		tglberikuttxt 		= rsskrg("NEXTDATE")
	end if
	rsskrg.Close
	set rsskrg = Nothing

end if
if Request("txtJenis")="BUKA" then
	sqlsimpan = "update CONFIGMST set PRODUCTIONMILLDATE=to_date('" & tglskrgtxt & "','mm/dd/yyyy')-1"
	connStr.execute(sqlsimpan)

	sqlskrg = "select PRODUCTIONMILLDATE,PRODUCTIONMILLDATE+1 NEXTDATE from configmst"
	Set rsskrg = Server.CreateObject("ADODB.Recordset")
	rsskrg.CursorType = 3
	rsskrg.CursorLocation = 3
	rsskrg.Open sqlskrg, connstr
	if not rsskrg.eof then
		tglsekarangtxt 		= rsskrg("PRODUCTIONMILLDATE")
		tglberikuttxt 		= rsskrg("NEXTDATE")
	end if
	rsskrg.Close
	set rsskrg = Nothing

	
	
end if
%>


<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-lock" aria-hidden="true"></i>Header Closing Produksi (LHP)</div><% end if %>
<form name="frmEntry" method="post" action="lhp_closing_produksi.asp">
	
	<table class="data" border="0" width="400px">
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="2" >CLOSING PRODUKSI HARIAN</td>
		</tr>
			
		<tr>
			<td align="left" class="lebarkolomsatu">Tgl Sekarang</td>
			<td align="left" class="lebarkolomdua" colspan="3">
				<input type="text" class="" maxlength="15" name="txttgl" readonly id="txttgl" size="15" value="<%=tglsekarangtxt%>" title="Tanggal Losses">
			</td>
		</tr>
		
		<tr>
			<td align="left" class="lebarkolomsatu">Tgl Berikut</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="" maxlength="15" name="tglberikut" readonly id="tglberikut" size="15" value="<%=tglberikuttxt%>" >
			</td>
		</tr>
		
		<tr>
			<td align="center" colspan="2">
				<input type="hidden" name="txtJenis" value="TUTUP">
		        <input Type="Submit"  name="Add" value="TUTUP">
		        <input Type="button"  id="buka" value="BUKA" onclick="getData()">
		        
			</td>
		</tr>
	</table>
</form>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
