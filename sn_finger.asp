<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Serial Number Fingerprint</title>
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

<script>
	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
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
   <div class="judul_head"  valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;"> 
 <!-------------awal------------->
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
DIM PeriodDate,	EMPCODE,FDDATE,FIELDCODE,UNDER_RIPE,OVER_RIPE,ABNORMAL,LONG_STALK,EATEN_BY_RAT,UNHARVEST_FFB,UNCOLLECT_LF_CIRCLE,UNCOLLECT_LF_PIECE
DIM UNARRANGE_FFB,UNPRUNE_FROND,QE_1,QE_2,QE_3,QE_4,QE_5,QE_6,QE_7,QE_8,QE_9,QE_10
Dim cmdView

Dim Header_EmpCode
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

txtyear =request("txtyear")
serial_number =request("serial_number")
search_sn =request("search_sn")
locationname=request("locationname")
location=request("location")
activation=request("activation")
fccode=request("fccode")

function getlocationname(location)
	result = ""
	q = " Select fcname from parameter where parcode = 'FP_LOCATION' and fcBA = '" & Session("fcBA") & "' and fccode = '"& location &"'"
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.Open q, connstr

	if not rsVH.eof then
		result = rsVH("fcname")
	end if

	getlocationname = result
end function

Function CheckPrimary(serial_number, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	" select count(1) total" & _
	                " from master_machine_fingerprint " & _
					" where serial_number ='" & serial_number & "'" & _
					" and fcba = '" & initialFcba & "'"

	 ' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckPrimary = initialAvailable
End Function

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		
			locationname = getlocationname(location)
			q = " Delete from master_machine_fingerprint " & _
				" where fcBA = '" & Session("fcBA") & "'" & _
				"       and serial_number = '" & serial_number & "'"
				' Response.Write q
				connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")="New" then	
	
		If Len(Trim(serial_number)) = 0 Then
			Err.Number = 1001
			Err.Description = "Serial Number Harus diisi, tidak boleh kosong"
		Else
			Dim PrimaryCode : PrimaryCode = 0
			PrimaryCode = CheckPrimary(serial_number, session("fcBA"))
			If PrimaryCode = 1 Then
				Err.Number = 1001
				Err.Description = "Serial Number """ & serial_number & """ sudah ada, mohon masukan Serial Number yang lain."
			Else
				If Len(Trim(location)) = 0 Then
					Err.Number = 1001
					Err.Description = "Location Harus diisi, tidak boleh kosong"
				Else

					locationname = getlocationname(location)

					q = " INSERT INTO master_machine_fingerprint" & _
						" (serial_number, location, locationname, activation, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) VALUES (" & _
						" 	'" & serial_number & "'" & _
						" , '" & location & "'" & _
						" , '" & locationname & "'" & _
						" , '" & activation & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						" , '" & Session("fcba") & "'" & _
						" , sysdate" & _
						" , to_char(sysdate, 'hh24:mi')" & _
						" )"
						' Response.Write q
						connStr.execute(q)
				End If
			End If
		End if
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PeriodDate)
		If Cek_IsDailyClosing = "FALSE" then
			
					locationname = getlocationname(location)
					q =	" Update master_machine_fingerprint" & _
						" set location = '" & location & "'" & _
						"     , activation = '" & activation & "'" & _
						"     , locationname = '" & locationname & "'" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" where fcBA = '" & Session("fcBA") & "'" & _
						"       and serial_number = '" & serial_number & "'"
		 			' response.write q
					connStr.execute(q)
					
		else
			err.number = 1001
			err.description = "Tanggal """ & PeriodDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis


txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   response.Write  "<div class = ""error"">" & err.description & "</div>"
End if
%>

<!-------------ISI Payment Form------------->
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-tablet" aria-hidden="true"></i>Header Serial Number Mesin Finger</div><% end if %>
 <form name="frmEntry" method="post" action="sn_finger.asp">
	<table class="data" width="100%">
		
		<tr>
			<td align="left" style="width:20%;">Search</td>
			<td align="left">
				<input type="text" size="15" name="fccode" value="<%=fccode%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
				
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" Then
	Dim rsVH,intRecordsPerPage,cColor
	set rsVH = server.CreateObject("Adodb.recordset")
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
		q	= " SELECT * FROM master_machine_fingerprint " & _
			  " WHERE fcba = '" & session("fcBA") & "'" 

		if ""&fccode <> "" then q = q & "AND serial_number = '" & fccode & "'" 

	 'response.Write q

	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="sn_finger.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
   		<%IF Request("txtJenis")="Edit" then%>
			<tr>
				<td align="left" style="width:20%;">Serial Number</td>
				<td align="left">
					<input type="text" size="15" name="serial_number" style="background:#eee;" value="<%=serial_number%>" autocomplete="off" maxlength="32" readonly="true">
					<font color="#FF0000">*</font>
				</td>
			</tr>
		<%Else%>
			<tr>
				<td align="left" style="width:20%;">Serial Number</td>
				<td align="left">
					<input type="text" size="15" name="serial_number" value="<%=serial_number%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
					<font color="#FF0000">*</font>
				</td>
			</tr>
		<%End If%>
			<tr> 
	            <td align="left"><font face="Verdana" size="1">Location</font></td>
					<td align="left"> 
						<select name="location" title="LOCATION NAME Character a s/d z or 1 s/d 9">
							<option value=""></option> 
								<% 	Dim rsComboV
								q = "SELECT * FROM PARAMETER WHERE PARCODE = 'FP_LOCATION' and FCBA = '"& Session("Fcba") &"' ORDER BY fccode"
								Set rsComboV = Server.CreateObject("ADODB.Recordset")
								rsComboV.CursorType = 3
								rsComboV.CursorLocation = 3
								rsComboV.Open q, connstr
								if not rsComboV.eof then%>
								<%
								Do While Not rsComboV.eof
								%>
								<%if cstr(""&location)=cstr(""&rsComboV("FCCODE")) then%>
									<option value="<%=rsComboV("FCCODE")%>" selected ><%=rsComboV("FCNAME")%></option>
									<%ELSE%>
									<option value="<%=rsComboV("FCCODE")%>"><%=rsComboV("FCNAME")%></option>
									<%end if%> 
								<%rsComboV.movenext
								loop
								End if
								rsComboV.Close
								set rsComboV = Nothing
								q=""
								%>
						</select> <font color="#FF0000">*</font>
					</td>
         	</tr>
         	<tr>
			    <td align="left"><font face="Verdana" size="1">Activation</font></td>
	            	<td align="left">
	            		<select name="ACTIVATION" title="ACTIVATION Character a s/d z or 1 s/d 9">
			                <Option value="TRUE">Yes</Option>
			                <Option value="FALSE">No</Option>
	              		</select>
              		</td>
          	</tr>
	    <tr>
				<td align="left" colspan="2">
					<input type="hidden" name="fccode" value="<%=fccode%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtJenis" value="Edit">
						<input Type="Submit"  name="Add" value="Update">
					<%Else%>
						<input type="hidden" name="txtJenis" value="New">
						<input Type="Submit"  name="Add" value="Add">
					<%End if%>

					<input Type="Reset"  name="Reset" value="Reset">
					<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				</td>
			</tr>
   </table>

  <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
  <div class="fbp-doc-scroll">
  <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="1" >&nbsp;</td>
		<td align="left">SERIAL NUMBER</td>
		<td align="left">LOCATION</td>
		<td align="left">LOCATION NAME</td>
		<td align="left">ACTIVATION</td>
	</tr>
	</thead>
	<tbody>
   
   <%
	if Not rsVH.eof then
		nCountofField = rsVH.fields.count - 1
		no = 1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=NO%></td>
						
			<td valign="middle" width="50" >
				<a href="sn_finger.asp?fccode=<%=fccode%>&serial_number=<%=rsvh("serial_number")%>&location=<%=rsvh("location")%>&locationname=<%=rsvh("locationname")%>&activation=<%=rsvh("activation")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>

				<a href="sn_finger.asp?fccode=<%=fccode%>&serial_number=<%=rsvh("serial_number")%>&location=<%=rsvh("location")%>&locationname=<%=rsvh("locationname")%>&activation=<%=rsvh("activation")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtyear")%>" border="0" ></a>
			
			</td>
			    <td align="left" title="<%=rsVH("serial_number").Name%>" valign="top"><%=rsVH("serial_number")%></td>
			    <td align="left" title="<%=rsVH("location").Name%>" valign="top"><%=rsVH("location")%></td>
			    <td align="left" title="<%=rsVH("locationname").Name%>" valign="top"><%=rsVH("locationname")%></td>
			    <td align="left" title="<%=rsVH("activation").Name%>" valign="top"><%=rsVH("activation")%></td>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
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
<%End if%>
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
