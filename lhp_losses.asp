<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<html>
	<head>
		<title>ePlaS - Laporan Harian Produksi</title>
		<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
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


function getTotal()
{
	var emptybunch = $("#emptybunch").val();
	var usb = $("#usb").val();
	var pcf = $("#pcf").val();
	var npc = $("#npc").val();
	var finalef = $("#finalef").val();
	var cyclonefiber = $("#cyclonefiber").val();
	var ltds1 = $("#ltds1").val();
	var ltds2 = $("#ltds2").val();
	var claybath = $("#claybath").val();
	var fruitloss = $("#fruitloss").val();
	var totaloil = parseFloat(emptybunch)+ parseFloat(usb)+ parseFloat(pcf)+ parseFloat(npc)+ parseFloat(finalef);
	var totalpk = parseFloat(cyclonefiber)+ parseFloat(ltds1)+ parseFloat(ltds2)+ parseFloat(claybath)+ parseFloat(fruitloss);
	var totaloildes = totaloil.toFixed(2);
	var totalpkdes = totalpk.toFixed(2);
	$("#totaloil").val(totaloildes);
	$("#totalpk").val(totalpkdes);
}
</script>
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
<body>
<!--#include file="top.jsp"-->

<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
<div class="judul">
	   Module : Losses
</div>
<div class="box_both" style="overflow:scroll;">

<%
Dim tgltxt,emptybunchtxt,cyclonefibertxt,usbtxt,ltds1txt
Dim pcftxt,ltds2txt,npctxt,claybathtxt,finaleftxt,fruitlosstxt
Dim totaloiltxt,totalpktxt,sludgectxt
Dim sqlsimpan,sqldelete,txtJenis

txtJenis 	= request("txtJenis")
tgltxt 		= request("txttgl")
emptybunchtxt 		= request("emptybunch")
cyclonefibertxt 	= request("cyclonefiber")
usbtxt 				= request("usb")
ltds1txt 			= request("ltds1")
pcftxt 				= request("pcf")
ltds2txt 			= request("ltds2")
npctxt 				= request("npc")
claybathtxt 		= request("claybath")
finaleftxt 			= request("finalef")
fruitlosstxt 		= request("fruitloss")
sludgectxt 			= request("sludgec")
totaloiltxt 		= request("totaloil")
totalpktxt 			= request("totalpk")

if Request("txtJenis")= "New" then

	'delete dulu sesuai tgl input'
	sqldelete = "delete from productionlosses where proddate=to_date('" & tgltxt & "','mm/dd/yyyy')"
	connStr.execute(sqldelete)

	'lalu insert'
	sqlsimpan = "insert into productionlosses(PRODDATE,EMPTYBUNCHOIL,USBOIL,PCFOIL,NPCOIL,FINALEFFLUENTOIL,SLUDGECENTRIFUGEOIL,CYCLONEFIBERPK,LTDSIPK,LTDSIIPK,CLAYBATHPK,FLINFBPK,TOTALOIL,TOTALPK) " & _
	" values(to_date('" & tgltxt & "','mm/dd/yyyy'),'" & emptybunchtxt & "','" & cyclonefibertxt & "', " & _
	" '" & usbtxt & "','" & ltds1txt & "','" & pcftxt & "','" & ltds2txt & "', " & _
	" '" & npctxt & "','" & claybathtxt & "','" & finaleftxt & "','" & fruitlosstxt & "', " & _
	" '" & sludgectxt & "','" & totaloiltxt & "','" & totalpktxt & "')"
	connStr.execute(sqlsimpan)
end if

%>


<form name="frmEntry" method="post" action="lhp_losses.asp">
	
	<table class="data" border="0" width="800px">
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="4" >LOSSES HARIAN</td>
		</tr>
			
		<tr>
			<td align="left" class="lebarkolomsatu">Tanggal</td>
			<td align="left" class="lebarkolomdua" colspan="3">
				<input type="text" class="datepicker" maxlength="15" name="txttgl" readonly id="txttgl" size="15" value="<%=Date%>" title="Tanggal Losses">
			</td>
		</tr>
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="2" >OIL LOSS</td>
			<td valign="top" align="center" colspan="2" >KERNEL LOSS</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">1. Empty Bunch</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="emptybunch" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="emptybunch" size="15" value="0" >
			</td>
			<td align="left" class="lebarkolomsatu">1. Cyclone Fiber</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="cyclonefiber" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="cyclonefiber" size="15" value="0">
			</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">2. USB</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="usb" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="usb" size="15" value="0" >
			</td>
			<td align="left" class="lebarkolomsatu">2. LTDS I</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="ltds1" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="ltds1" size="15" value="0">
			</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">3. Press Cake Fibre</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="pcf" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="pcf" size="15" value="0" >
			</td>
			<td align="left" class="lebarkolomsatu">    LTDS II</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="ltds2" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="ltds2" size="15" value="0">
			</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">4. Nut Press Cake</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="npc" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="npc" size="15" value="0" >
			</td>
			<td align="left" class="lebarkolomsatu">3. Claybath</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="claybath" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="claybath" size="15" value="0">
			</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">5. Final Effluent</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="finalef" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="finalef" size="15" value="0" >
			</td>
			<td align="left" class="lebarkolomsatu">4. Fruit Loss in EB</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="fruitloss" onkeyup="getTotal()" onkeypress="return harusangka(event);" id="fruitloss" size="15" value="0">
			</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu"><strong>Total</strong></td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="totaloil" readonly onkeypress="return harusangka(event);" id="totaloil" size="15" value="0" >
			</td>
			<td align="left" class="lebarkolomsatu"><strong>Total</strong></td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="totalpk" readonly onkeypress="return harusangka(event);" id="totalpk" size="15" value="0">
			</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">6. Sludge Centrifuge</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" maxlength="15" name="sludgec" onkeypress="return harusangka(event);" id="sludgec" size="15" value="0" >
			</td>
		</tr>
		<tr>
			<td align="center" colspan="4">
				<input type="hidden" name="txtJenis" value="New">
		        <input Type="Submit"  name="Add" value="Simpan">
			</td>
		</tr>
	</table>
</div>
</div>
</form>
</div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
