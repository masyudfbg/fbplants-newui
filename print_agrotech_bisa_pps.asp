<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
validatesession

txtyear = request("txtyear")
txtmonth = request("txtmonth")
txtindeks = request("txtindeks")
txtdivision = request("txtdivision")
txtfcba = request("txtfcba")

q = " SELECT fcname" & _
    " FROM setup_bisa_pps_tab ch" & _
    " WHERE fccode = '" & txtindeks & "'"

'response.Write q	
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open q, connstr
if not oRecordset.eof then
	fcname_indeks = oRecordset("fcname")
End if
oRecordset.close
set oRecordset = Nothing

%>
<html>
	<head>
		<title>BISA <%=fcname_indeks%></title>
		<link rel="stylesheet" href="/plugins/paper-css/paper.min.css">
		<style>
			@page { size: A4 }
			table {
				border-collapse: collapse;
				width:100%;
				font: 11px arial, sans-serif;
			}

			th, td {
				border: 1px solid #111;
			}
			
			th, td {
				padding: 7px 5px;
			}
			td.sideonly {
				border-top: none;
				border-bottom: none;
			}
			td.noborder {border:none;}
			article{font: 12px arial, sans-serif;}
			#wrapper {
				float: right;
				margin-bottom:10px;
			}
			
			input[type="button"]{
				background: #fff;
				-webkit-border-radius: 4;
				-moz-border-radius: 4;
				border-radius: 4px;
				font-family: Arial;
				color: #000;
				font-size: 12px;
				padding: 10px 20px 10px 20px;
				text-decoration: none;
				cursor:pointer;
				display:block;
				border-style: solid;
				border-width: 1px;
			}
			input.tableToExcel[type="button"] {
				background: #5bb75b;
				background-image: -webkit-linear-gradient(top, #5bb75b, #5bb75b);
				background-image: -moz-linear-gradient(top, #5bb75b, #5bb75b);
				background-image: -ms-linear-gradient(top, #5bb75b, #5bb75b);
				background-image: -o-linear-gradient(top, #5bb75b, #5bb75b);
				background-image: linear-gradient(to bottom, #5bb75b, #5bb75b);
				-webkit-border-radius: 4;
				-moz-border-radius: 4;
				border-radius: 4px;
				font-family: Arial;
				color: #ffffff;
				font-size: 12px;
				padding: 7px 10px;
				text-decoration: none;
			}			
			input.tableToExcel[type="button"]:hover {
				background: #4da34d;
				background-image: -webkit-linear-gradient(top, #4da34d, #4da34d);
				background-image: -moz-linear-gradient(top, #4da34d, #4da34d);
				background-image: -ms-linear-gradient(top, #4da34d, #4da34d);
				background-image: -o-linear-gradient(top, #4da34d, #4da34d);
				background-image: linear-gradient(to bottom, #4da34d, #4da34d);
				text-decoration: none;
			}
			.clear {clear:both;}
			@media print {
				#wrapper,.clear{display:none;}
			}
		</style>
	</head>
	<body>
		<section class="sheet padding-10mm">
		<div id="wrapper"><!--#include file="print_excel.asp"--></div>
		<div class="clear"></div>
		<article id="product-table">
<%

q = " SELECT fccompanyname, fcmanager" & _
    " FROM businessunit" & _
    " WHERE fccode = '" & txtfcba & "'"

'response.Write q	
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open q, connstr
if not oRecordset.eof then
	fccompanyname = oRecordset("fccompanyname")
	fcmanager = oRecordset("fcmanager")
End if
oRecordset.close
set oRecordset = Nothing

q = " SELECT data.fccode, master.fcname, SUM(data.conformation) conformation, master.isentry FROM (" & _
	    " SELECT fccode, 0 conformation" & _
		" FROM (" & _
		"     SELECT DISTINCT fccode, fcname" & _
		"     FROM setup_bisa_pps_tab ch" & _
		"     WHERE SUBSTR(parent_fccode,1,2) = '" & txtindeks & "'" & _
		"     CONNECT BY PRIOR parent_fccode = fccode" & _
		"     START WITH levelid <> '1'" & _
		" ) UNION ALL" & _
		" SELECT fccode, conformation FROM trans_bisa_pps_dtl" & _
		" WHERE fcba = '" & txtfcba & "'" & _
		"       AND accountingyear = '" & txtyear & "'" & _
		"       AND accountingperiod = '" & txtmonth & "'" & _
		"       AND division = '" & txtdivision & "'" & _
		"       AND conformation > 0" & _
		" ) data LEFT JOIN setup_bisa_pps_tab master ON master.fccode = data.fccode" & _
		" WHERE SUBSTR(master.parent_fccode,1,2) = '" & txtindeks & "'" & _
		" GROUP BY data.fccode, master.fcname, master.isentry" & _
		" ORDER BY data.fccode"

'response.Write q	
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open q, connstr

arr_rdata_eof = true
if not oRecordset.eof then
	arr_rdata = oRecordset.getrows()
	arr_rdata_eof = false
End if
oRecordset.close
set oRecordset = Nothing

if not arr_rdata_eof then
	count_data_input = 0
	standar_data_input = 0
	for yy=0 to ubound(arr_rdata,2)
		if cdbl(to_nol(arr_rdata(3,yy))) = 1 then
			count_data_input = count_data_input +1
			
			if cdbl(to_nol(arr_rdata(2,yy))) = 1 then
				standar_data_input = standar_data_input +1
			end if
		end if
	next
end if

p_persen = 0
if count_data_input = 0 then
	p_persen = 0
else
	p_persen = Round(standar_data_input/count_data_input*100,0)
end if

if p_persen > 95 then
	style_color = "background-color:#0074D9;"
elseif p_persen >= 80 then
	style_color = "background-color:#01FF70;"
elseif p_persen >= 70 then
	style_color = "background-color:#FABF8F;"
elseif p_persen >= 60 then
	style_color = "background-color:#FFDC00;"
else
	style_color = "background-color:#FF4136;"
end if



response.write "<div style='padding:7px 0;'></div>"
response.write "<table>"
response.write "	<tbody>"
response.write "		<tr>"
response.write "			<td align='center' colspan='5' style='border:none;'><b><i></i></b></td>"
response.write "			<td align='center' colspan='2' rowspan='3' style='border:1px solid #000;font-size:3em;font-weight:bold;"&style_color&"'>"&p_persen&" %</td>"
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td align='center' colspan='5' style='border:none;font-size:2em;font-weight:bold;'>INDEKS PENILAIAN BISA "&UCASE(fcname_indeks)&"</td>"
response.write "		</tr>"
response.write "		<tr>"
response.write "			<td align='center' colspan='5' style='border:none;font-size:12px;font-weight:bold;'>"&Bulan(txtmonth) & " " & txtyear&"</td>"
response.write "		</tr>"
response.write "	</tbody>"
response.write "</table>"

response.write "<div style='padding:7px 0;'></div>"
response.write "<table>"
response.write "	<tbody>"
if cstr("" & fccompanyname) <> "" then
	response.write "		<tr><td colspan='7' style='border:none;'><b>PT. "& fccompanyname&"</b></td></tr>"
else
	response.write "		<tr><td colspan='7' style='border:none;'><b></b></td></tr>"
end if
response.write "		<tr><td colspan='7' style='border:none;'><b>Kebun/ Afdeling : "&txtfcba&" . "&txtdivision&"</b></td></tr>"
response.write "	</tbody>"
response.write "</table>"






if not arr_rdata_eof then
	no = 0
	total_standar = 0
	total_tidakstandar = 0
	response.write "<table>"
	response.write "	<thead>"
	response.write "		<tr style='background-color:#00A300;color:#FFFFFF;'>"
	response.write "			<th style='min-width:7mm;max-width:7mm;'>No</th>"
	response.write "			<th colspan='2' style='min-width:40mm;max-width:40mm;'>Item Kegiatan</th>"
	response.write "			<th style='min-width:10mm;max-width:10mm;'>Standar</th>"
	response.write "			<th style='min-width:10mm;max-width:10mm;'>Tidak Standar</th>"
	response.write "			<th colspan='2' style='min-width:20mm;max-width:20mm;'>Keterangan Kondisi</th>"
	response.write "		</tr>"
	response.write "	</thead>"
	response.write "	<tbody>"
	for yy=0 to ubound(arr_rdata,2)
		if cdbl(to_nol(arr_rdata(3,yy))) = 1 then
			no = no +1
			response.write "		<tr>"
			response.write "			<td align='center' class='sideonly'>"&no&"</td>"
			response.write "			<td colspan='2' class='sideonly'>"&arr_rdata(1,yy)&"</td>"
			if cdbl(to_nol(arr_rdata(2,yy))) = 1 then
				total_standar = total_standar +1
				response.write "			<td align='center' class='sideonly'>"& "v" &"</td>"
			else
				response.write "			<td class='sideonly'></td>"
			end if
			if cdbl(to_nol(arr_rdata(2,yy))) = 0 then
				total_tidakstandar = total_tidakstandar +1
				response.write "			<td align='center' class='sideonly'>"& "v" &"</td>"
			else
				response.write "			<td class='sideonly'></td>"
			end if
			response.write "			<td colspan='2' class='sideonly'></td>"
			response.write "		</tr>"
		else
			response.write "<tr style='background-color:#FF851B;color:#FFFFFF;' ><td class='sideonly' colspan='7'><u>"&arr_rdata(1,yy)&"</u></td></tr>"
		end if
		
	next
	response.write "<tr>"
	response.write "	<td colspan='3' align='center'><b>Total</b></td>"
	response.write "	<td align='center'><b>"&FormatNumber(total_standar,0)&"</b></td>"
	response.write "	<td align='center'><b>"&FormatNumber(total_tidakstandar,0)&"</b></td>"
	response.write "	<td colspan='2' align='center'></td>"
	response.write "</tr>"
end if

response.write "<div style='padding:7px 0;'></div>"
response.write "<table>"
response.write "	<tbody>"
response.write "		<tr>"
response.write "			<td width='24%' align='center' colspan='2' rowspan='3' style='border:none;padding:0;padding-top:5px;'>"
								call legenda
response.write "            </td>"
'response.write "			<td width='24%' align='center' colspan='2' style='border:none;'><b><i></i></b></td>"
response.write "			<td width='24%' align='center' colspan='2' style='border:none;'>Dibuat Oleh,</td>"
response.write "			<td width='24%' align='center' style='border:none;'><b><i></i></b></td>"
response.write "			<td width='24%' align='center' colspan='2' style='border:none;'>Diketahui Oleh,</td>"
response.write "		</tr>"
response.write "		<tr>"
'response.write "			<td width='24%' colspan='2' style='border:none;'>&nbsp;</td>"
response.write "			<td width='24%' colspan='2' style='border:none;'>&nbsp;</td>"
response.write "			<td width='24%' style='border:none;'>&nbsp;</td>"
response.write "			<td width='24%' colspan='2' align='center' style='border:none;padding:80px 0 0 0;'>"&fcmanager&"</td>"
response.write "		</tr>"
response.write "		<tr>"
'response.write "			<td width='24%' align='center' colspan='2' style='border:none;'><b><i></i></b></td>"
response.write "			<td width='24%' align='center' colspan='2' style='border:none;'><div style='border-top: 1px solid #000;width:70%;padding-top:5px;'><b>Asst "&txtdivision&"</b></div></td>"
response.write "			<td width='24%' align='center' style='border:none;'><b><i></i></b></td>"
response.write "			<td width='24%' align='center' colspan='2' style='border:none;padding:0;'><div style='border-top: 1px solid #000;width:70%;padding-top:5px;'><b>Estate Manager</b></div></td>"
response.write "		</tr>"
response.write "	</tbody>"
response.write "</table>"
%>
		</article>
		</section>
		
		
		<style>
		.stick { display:none;}
		.xfTBarF { height:15px; line-height: 15px; font-size:10px; padding-left:20px; background-color:#008001;}
		.xfContainer { border:1px #008001 solid; box-shadow:1px 1px 1px rgba(0,0,0,0.5); background:#FFFFFF;}
		.fenster-content { padding:5px; padding-top:10px;}
		#indicator_color { border:1px #999999 solid ; padding:2px; font-family:Arial, Helvetica, sans-serif; }
		#indicator_color table { border-collapse:collapse; border-spacing:0px;}
		#indicator_color table, #indicator_color tr, #indicator_color td { border: none; line-height:15px; }
		#indicator_color th { font-size:12px; background-color:#222222; border-left:1px #666666 dashed; color:#FFFFFF;}
		#indicator_color .2nd { background-color:#f1f1f1;}
		#indicator_color td { font-size:11px; }
		.merah { width:15px; height:15px; background-color:#FF0000; float:left; margin-right:5px; margin-top:7px;
				  border-radius:3px;
				  -moz-border-radius:3px;
				  -webkit-border-radius:3px;
				  -o-border-radius:3px;
				  box-shadow:1px 1px 0px #ccc;
				  -moz-box-shadow:1px 1px 0px #ccc;
				  -webkit-box-shadow:1px 1px 0px #ccc;
				  -o-box-shadow:1px 1px 0px #ccc;}
		.kuning { width:15px; height:15px; background-color:#FFFF00; float:left; margin-right:5px; margin-top:7px;
				   border-radius:3px;
				  -moz-border-radius:3px;
				  -webkit-border-radius:3px;
				  -o-border-radius:3px;
				  box-shadow:1px 1px 0px #ccc;
				  -moz-box-shadow:1px 1px 0px #ccc;
				  -webkit-box-shadow:1px 1px 0px #ccc;
				  -o-box-shadow:1px 1px 0px #ccc;}
		.hijau { width:15px; height:15px; background-color:#008001; float:left; margin-right:5px; margin-top:7px;
				  border-radius:3px;
				  -moz-border-radius:3px;
				  -webkit-border-radius:3px;
				  -o-border-radius:3px;
				  box-shadow:1px 1px 0px #ccc;
				  -moz-box-shadow:1px 1px 0px #ccc;
				  -webkit-box-shadow:1px 1px 0px #ccc;
				  -o-box-shadow:1px 1px 0px #ccc;}
		#indicator_color .span_color { float:left;}
		#indicator_color .point_table {}
		#indicator_color .MsoNormal { margin-bottom:3px; padding:0px 5px;}
		.symbol_1 { font-size:18px; color:#008001; margin-right:5px; }
		</style>
		
<%Sub legenda()%>

<div id="indicator_color" style="margin-bottom:5px;">
	<table cellpadding="0" border="1" style="border:none;padding:0;margin:0;">
			<tr>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#0074D9;color:#0074D9;padding-left:10px;">&nbsp;</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;width:20%;">
					<p class="MsoNormal">> 90</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Excellent</p>
				</td>
			</tr>
			
			<tr>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#01FF70;padding-left:10px;">&nbsp;</p>
				</td>				
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">80 - 95</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Good</p>
				</td>
			</tr>
			
			<tr>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#FABF8F;padding-left:10px;">&nbsp;</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">70 - 79</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Satisfactory</p>
				</td>
			</tr>
			<tr>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#FFDC00;padding-left:10px;">&nbsp;</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">60 - 69</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Fair</p>
				</td>
			</tr>
			<tr>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#FF4136;padding-left:10px;">&nbsp;</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">< 60</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Poor</p>
				</td>
			</tr>
	</table>
</div>
<%End Sub%>
	</body>
</html>