<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Vehicle Operation</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
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
</script>
<script language="javascript" type="text/javascript">



                function showHint(str)
					{
				var hoursVehicle;
				var vehiclecode_;	
				
				vehiclecode_ = document.getElementById("combo_vehiclecode").value;
				
					if (str=="")
					  {
					  document.getElementById("hours").innerHTML="";
					  document.getElementById("unit").innerHTML="";
					  document.getElementById("output").innerHTML="";
					  
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  hoursVehicle=new XMLHttpRequest();
					  unitVehicle=new XMLHttpRequest();
					  outputVehicle=new XMLHttpRequest();

					  }
					else
					  {// code for IE6, IE5
					  hoursVehicle=new ActiveXObject("Microsoft.XMLHTTP");
					  unitVehicle=new ActiveXObject("Microsoft.XMLHTTP");
					  outputVehicle=new ActiveXObject("Microsoft.XMLHTTP");

					  }
					  
						hoursVehicle.onreadystatechange=function()
						  {
						  if (hoursVehicle.readyState==4 && hoursVehicle.status==200)
							{
							document.getElementById("hours").innerHTML=hoursVehicle.responseText;
							}
						  }
						hoursVehicle.open("GET","ajax/Vehicle_ann.asp?ajaxUoM=hours&FcCode="+ vehiclecode_ +"&q="+str,true);
						hoursVehicle.send();
						
						unitVehicle.onreadystatechange=function()
						  {
						  if (unitVehicle.readyState==4 && unitVehicle.status==200)
							{
							document.getElementById("unit").innerHTML=unitVehicle.responseText;
							}
						  }
						unitVehicle.open("GET","ajax/Vehicle_ann.asp?ajaxUoM=unit&FcCode="+ vehiclecode_ +"&q="+str,true);
						unitVehicle.send();
						
						outputVehicle.onreadystatechange=function()
						  {
						  if (outputVehicle.readyState==4 && outputVehicle.status==200)
							{
							document.getElementById("output").innerHTML=outputVehicle.responseText;
							}
						  }
						outputVehicle.open("GET","ajax/Vehicle_ann.asp?ajaxUoM=output&FcCode="+ vehiclecode_ +"&q="+str,true);
						outputVehicle.send();
	
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
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <!-------------awal------------->
<%
Cek_IsDailyClosing = ""
 
vehicledate = request("vehicledate")
vehiclecode = request("vehiclecode")
vehiclestatus = request("vehiclestatus")
unit = request("unit")
fuel = request("fuel")
harihujan = request("harihujan")
curahhujan = request("curahhujan")
remarks = request("remarks")
lineno = request("lineno")

if "" & vehicledate = "" then vehicledate = textDate(date)


if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fdDate)
		If Cek_IsDailyClosing = "FALSE" then
			q = " DELETE FROM vehicleoperation WHERE fcba = '" & Session("fcBA") & "' AND lineno = '" & lineno & "'"

			' response.write q
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & vehicledate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), vehicledate)
		If Cek_IsDailyClosing = "FALSE" then
		
			lineno = 0
			qMaxLineNo = " SELECT NVL(MAX(lineno),0) maxLineNo " & _
						 " FROM vehicleoperation " & _
						 " WHERE fcba = '" & Session("fcBA") & "' "
							
			Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
			rsqMaxLineNo.CursorType = 3
			rsqMaxLineNo.CursorLocation = 3
			rsqMaxLineNo.Open qMaxLineNo, connstr

			if not rsqMaxLineNo.eof then
				lineno = cdbl(rsqMaxLineNo("maxLineNo"))
			end if
			rsqMaxLineNo.close
			set rsqMaxLineNo = Nothing
			
			has_record = 0
			qMaxLineNo = " SELECT COUNT(1) totalrecord " & _
						 " FROM vehicleoperation " & _
						 " WHERE vehicledate = TO_DATE('" & vehicledate & "','mm/dd/yyyy')" & _
						 "       AND fcba = '" & Session("fcBA") & "'" & _
						 "       AND vehiclecode = '" & vehiclecode & "'"
							
			Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
			rsqMaxLineNo.CursorType = 3
			rsqMaxLineNo.CursorLocation = 3
			rsqMaxLineNo.Open qMaxLineNo, connstr

			if not rsqMaxLineNo.eof then
				has_record = cdbl(rsqMaxLineNo("totalrecord"))
			end if
			rsqMaxLineNo.close
			set rsqMaxLineNo = Nothing
			
			if has_record = 0  then
				lineno = lineno +1
				q = " INSERT INTO vehicleoperation ( lineno" & _
					"                                , vehicledate" & _
					"                                , vehiclecode" & _
					"                                , vehiclestatus" & _
					"                                , unit" & _
					"                                , fuel" & _
					"                                , harihujan" & _
					"                                , curahhujan" & _
					"                                , remarks" & _
					"                                , fcentry" & _
					"                                , fcedit" & _
					"                                , fcip" & _
					"                                , fcba" & _
					"                                , lastupdate" & _
					"                                , lasttime )" & _
					" VALUES ( '" & lineno & "'" & _
					"          , TO_DATE('" & vehicledate & "','mm/dd/yyyy')" & _
					"          , '" & vehiclecode & "'" & _
					"          , '" & vehiclestatus & "'" & _
					"          , '" & unit & "'" & _
					"          , '" & fuel & "'" & _
					"          , '" & harihujan & "'" & _
					"          , '" & curahhujan & "'" & _
					"          , '" & remarks & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					"          , '" & Session("fcBA") & "'" & _
					"          , SYSDATE" & _
					"          , TO_CHAR(SYSDATE,'HH24:MI') )"
						
				'response.Write(q)
				connStr.execute(q)
			else
				err.number = 1001
				err.description = "Date kendaraan " & vehiclecode & " Tanggal " & vehicledate & " sudah di entry, Silahkan cek kembali inputan anda"
			End If
		else
			err.number = 1001
			err.description = "Tanggal """ & vehicledate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), vehicledate)
		If Cek_IsDailyClosing = "FALSE" then
			q = " UPDATE vehicleoperation" & _
                " SET vehiclecode = '" & vehiclecode & "'" & _
                "     , vehiclestatus = '" & vehiclestatus & "'" & _					
                "     , unit = '" & unit & "'" & _
                "     , fuel = '" & fuel & "'" & _
                "     , harihujan = '" & harihujan & "'" & _
                "     , curahhujan = '" & curahhujan & "'" & _
                "     , remarks = '" & remarks & "'" & _
                "     , fcedit = '" & Session("CodeName") & "'" & _
                "     , lastupdate = SYSDATE " & _
                "     , lasttime = TO_CHAR(SYSDATE, 'HH24:MI') "& _
                " WHERE fcba = '" & Session("fcBA") & "' AND lineno = '" & lineno & "'"
					
			' response.write q
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & vehicledate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

if request("txtJenis") = "Edit" then
	oSQLQuery = " SELECT *" & _
				" FROM vehicleoperation" & _
				" WHERE fcba = '" & Session("fcBA") & "'" & _
				"       AND lineno = '" & lineno & "'"
				
	set oRecordSet = server.CreateObject("Adodb.recordset")
	oRecordSet.Open oSQLQuery, connstr
	If Not oRecordSet.Eof then
		vehiclecode = oRecordSet("vehiclecode")
		vehicledate = oRecordSet("vehicledate")
		vehiclestatus = oRecordSet("vehiclestatus")
		unit = oRecordSet("unit")
		fuel = oRecordSet("fuel")
		harihujan = oRecordSet("harihujan")
		curahhujan = oRecordSet("curahhujan")
		remarks = oRecordSet("remarks")
	End if
	'response.Write oSQLQuery
	oRecordSet.Close
	Set oRecordSet = Nothing
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & vehicledate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & vehicledate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Vehicle')"

set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
'response.Write oSQLQuery
oRecordSet.Close
Set oRecordSet = Nothing

%>
 <!-------------ISI Payment Form------------->

 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-road" aria-hidden="true"></i>Header Vehicle Operation</div><% end if %>
 <form name="frmEntry" method="post" action="VehicleOperation.asp">
	<table class="data" style="width:70%;">
		<tr>
			<td align="left" style="width:20%;">Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="vehicledate" size="10" value="<%=vehicledate%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>	
		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;
			</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" and "" & vehicledate <> "" Then	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q = " SELECT vo.lineno" & _
		"        , vo.vehicledate" & _
		"        , vo.vehiclecode" & _
		"        , vo.vehiclestatus" & _
		"        , pr.fcname vehiclestatusname" & _
		"        , vo.unit" & _
		"        , vo.fuel" & _
		"        , vo.harihujan" & _
		"        , vo.curahhujan" & _
		"        , vo.remarks" & _
		" FROM vehicleoperation vo" & _
		" LEFT JOIN parameter pr ON (pr.fcba = vo.fcba AND pr.fccode = vo.vehiclestatus AND pr.parcode = 'VO001')" & _
		" WHERE vo.vehicledate = TO_DATE('" & vehicledate & "', 'MM/DD/YYYY')" & _
        "       AND vo.fcba = '" & session("fcBA") & "'" & _
        " ORDER BY vo.vehiclecode"
	'response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.Open q, connstr
%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="VehicleOperation.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" style="width:70%;">
		<!--Yang tidak bisa di edit adalah PK Table-->
		<tr>
			<td  align="left" style="width:20%;">Vehicle Code</td>
			<td td  align="left">
				<select id="vehiclecode" name="vehiclecode" <%if request("txtJenis") = "Edit" then%> disabled <%end if%>>
					<%cmdCallCombo "Vehicle",vehiclecode,"",""%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Vehicle Status</td>
			<td align="left">
				<select id="vehiclestatus" name="vehiclestatus">
				<%
				q = " SELECT fccode, fcname FROM parameter where fcba = '" &  Session("fcBA") & "' AND parcode = 'VO001'"
				Set rsComboA = Server.CreateObject("ADODB.Recordset")
				
				rsComboA.Open q, connstr
				if not rsComboA.eof then
					Do While Not rsComboA.eof
				%>
					<option value="<%=rsComboA("fccode")%>" <% If cstr("" & vehiclestatus) = cstr(rsComboA("fccode")) Then %> Selected <%End if%>><%=rsComboA("fcname")%></option> 
				<% 
						rsComboA.movenext
					loop
				End if
				rsComboA.Close : set rsComboA = Nothing : q = ""
				%>
				</select> <font color="#FF0000">*</font>
			</td>
		</tr>
		<!--
		<tr>
			<td align="Left">Hm/Km/Jam</td>
			<td align="Left">
				<input type="text" size="15" name="UNIT" value="<%=UNIT%>">
			</td>
		</tr>
		<tr>
			<td align="Left">Fuel (Ltr)</td>
			<td align="Left">
				<input type="text" size="15" name="fuel" value="<%=fuel%>">
			</td>
		</tr>
		<tr>
			<td align="Left">Hari Hujan (Hari)</td>
			<td align="Left">
				<input type="text" size="15" name="harihujan" value="<%=harihujan%>">
			</td>
		</tr>
		<tr>
			<td align="Left">Curah Hujan (mm)</td>
			<td align="Left">
				<input type="text" size="15" name="curahhujan" value="<%=curahhujan%>">
			</td>
		</tr>-->
		<tr>
			<td align="Left">Remarks</td>
			<td align="Left">
				<textarea id="remarks" name="remarks" cols="50" rows="3"><%=remarks%></textarea>
			</td>
		</tr>
	    <TR><td colspan="2" align="Left">
	      <input type="hidden" name="vehicledate" value="<%=vehicledate%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <!-- Validasi Period Control In Here -->
		  <%IF Request("txtJenis")="Edit" then%>
			<input type="hidden" name="vehiclecode" value="<%=vehiclecode%>">
			<input type="hidden" name="lineno" value="<%=lineno%>">
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update" onclick="return validasiform();">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add" onclick="return validasiform();">
		  <%End if%>
		   <!-- End Validasi Period Control In Here -->
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	<TD align="left">No</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">Date</TD>
	<TD align="left">Vehicle Code</TD>
	<TD align="left">Vehicle Status</TD>
	<!--<TD align="left">Hm/Km/Jam</TD>
	<TD align="left">Fuel (Ltr)</TD>
	<TD align="left">Hari Hujan (Hari)</TD>
	<TD align="left">Curah Hujan (mm)</TD>-->
	<TD align="left">Remarks</TD>
	</tr>
	</thead>
	<tbody>
   
   <%
	if Not rsVH.eof then
		no=1
		While Not rsVH.EOF
			if no  mod 2 = 1 then
				cColor = "#EEEEEE"
			else
				cColor = "#E5E5E5"
			end if%>
			<tr bgcolor="<%=cColor%>">
				<td align="left" valign="top"><%=NO%></td>
				 <!-- Validasi Period Control In Here -->
					<td>
						<a href="VehicleOperation.asp?lineno=<%=rsVH("lineno")%>&vehicledate=<%=rsVH("vehicledate")%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0">
						</a>
					</td>
					<td>
						<a href="VehicleOperation.asp?lineno=<%=rsVH("lineno")%>&vehicledate=<%=rsVH("vehicledate")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');">
							<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0">
						</a>
					</td>
				 <!-- End Validasi Period Control In Here -->
				<td align="left" title="<%=rsVH("vehicledate").Name%>" valign="top"><%=rsVH("vehicledate")%></td>
				<td align="left" title="<%=rsVH("vehiclecode").Name%>" valign="top"><%=rsVH("vehiclecode")%></td>				
				<td align="left" title="<%=rsVH("vehiclestatusname").Name%>" valign="top"><%=rsVH("vehiclestatusname")%></td>
				<!--<td align="left" title="<%=rsVH("unit").Name%>" valign="top"><%=rsVH("unit")%></td>
				<td align="left" title="<%=rsVH("fuel").Name%>" valign="top"><%=rsVH("fuel")%></td>
				<td align="left" title="<%=rsVH("harihujan").Name%>" valign="top"><%=rsVH("harihujan")%></td>
				<td align="left" title="<%=rsVH("curahhujan").Name%>" valign="top"><%=rsVH("curahhujan")%></td>-->
				<td align="left" title="<%=rsVH("remarks").Name%>" valign="top"><%=rsVH("remarks")%></td>
			</tr>
		  <%
		  rsVH.MoveNext
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </tbody>
   </TABLE>
   </div>
<% if Request("oldui") <> "1" then %>
<script type="text/javascript">
// Status baris jadi badge (tampilan saja)
(function () {
	var cells = document.querySelectorAll('table.fbp-doc-lines tbody td[title]');
	for (var i = 0; i < cells.length; i++) {
		if (cells[i].getAttribute('title').toUpperCase() !== 'ROWSTATEID') continue;
		var s = (cells[i].textContent || '').replace(/\s+/g, ' ').trim();
		if (!s) continue;
		cells[i].parentNode.setAttribute('data-state', s);
		cells[i].innerHTML = '';
		var b = document.createElement('span');
		b.className = 'fbp-state';
		b.textContent = s;
		cells[i].appendChild(b);
	}
})();
</script>
<% end if %>
   </form>
<%end if%>
  <!-------------akhir------------->
</div></div>
  <!--#include file="bottom.jsp"-->
<script type="text/javascript">
    function validasiform() {
	
		var vc = document.getElementById('vehiclecode');
		var vs = document.getElementById('vehiclestatus');
		var rm = document.getElementById('remarks');
        if (vc.value === '') {
            alert("kode kendaraan harus diisi");
            return false;
        }
		
		if (vs.value === '3' || vs.value == '4'){
			if (rm.value === '') {
				alert("kolom keterangan harus diisi");
				return false;
			}
		}
        
    }
</script>
</body>
</html>
