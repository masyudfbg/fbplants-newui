<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Weather</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

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
<%
Cek_IsDailyClosing = ""
txttable=request("txttable")
WIDATE=request("WIDATE")
WIDIVISION=request("WIDIVISION")
TOTRNINCHES=request("TOTRNINCHES")
TOTRNDAYS=request("TOTRNDAYS")
AVGTEMP=request("AVGTEMP")
AVGWIND=request("AVGWIND")

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then


	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), WIDATE)
		If Cek_IsDailyClosing = "FALSE" then
			   q = " DELETE FROM weather WHERE widate = TO_DATE('" & WIDATE & "','mm/dd/yyyy') AND widivision = '" & WIDIVISION & "' AND fcba = '" & Session("fcBA") & "'"
			   connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & WIDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), WIDATE)
		If Cek_IsDailyClosing = "FALSE" then
			  q="Insert into Weather (WIDATE, WIDIVISION, TOTRNINCHES, TOTRNDAYS, AVGTEMP, AVGWIND,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				"values (to_Date('" & WIDATE & "','MM/DD/YYYY'),'" &  WIDIVISION & "','" & TOTRNINCHES & "','" &  TOTRNDAYS & "','" &  AVGTEMP & "'," & _
				"'" & AVGWIND & "'," & _		
				"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
				'response.Write(q)
				'response.Flush()
				connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & WIDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), WIDATE)
		If Cek_IsDailyClosing = "FALSE" then
			   q = " UPDATE weather" & _
				   " SET totrninches = '" & TOTRNINCHES & "'" & _
				   "    , totrndays = '" & TOTRNDAYS & "'" & _
				   "    , avgtemp = '" & AVGTEMP & "'" & _
				   "    , avgwind = '" & AVGWIND & "'" & _
				   "    , lastupdate = SYSDATE " & _
				   "    , lasttime = TO_CHAR(SYSDATE,'hh24:mi') " & _
				   " WHERE widate = TO_DATE('" & WIDATE & "','mm/dd/yyyy') AND widivision = '" & WIDIVISION & "' AND fcba = '" & Session("fcBA") & "'"
				 connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & WIDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
         cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & 
else
    cmdView	= request("cmdView")
End if

if "" & WIDATE = "" then WIDATE = textDate(date)

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   response.Write  "<div class = ""error"">" & err.description & "</div>"
End if

 '-------------ISI Payment Form 
 %>
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-cloud" aria-hidden="true"></i>Header Weather</div><% end if %>
<form name="frmEntry" method="post" action="Weather.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
		<TR> 
			<td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE WEATHER</strong></font></div></td>
		</TR>
		<tr>
			<td align="left">Date/ Tanggal</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="WIDATE" size="10" value="<%=WIDATE%>"></td>
		</tr>
		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="PrintWeather.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&PeriodDate=<%=WIDATE%>" title="Print Data" target="_blank">
					<img src="IMAGES/Print.gif" border="0">
				</a>
				<!-- HADIR -->
				<%
				dataModule 		= "Weather"
				dataSubmodule 	= "Weather Information"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>">
					<i class="fa fa-check-square-o"></i>&nbsp;Hadir
				</a>
				<!-- END:HADIR -->
			</td>
		</tr>
	</table>  
</form>

<%
if cmdView = "True" Then

	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q = " SELECT widate, widivision, totrninches, totrndays, avgtemp, avgwind, fcentry, fcedit, fcip, fcba" & _
	    " FROM weather" & _
	    " WHERE fcba = '" & session("fcBA") & "'" & _
		"       AND widate = to_date('" & WIDATE & "','mm/dd/yyyy')"
		
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
   rsGAD.Open q, connstr
%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="Weather.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">
		<tr>
			<td align="left">Division</td>
			<td align="left">
				<select name="WIDIVISION"<%if txtJenis="Edit" then%> disabled <%end if%>>
				<%
				q = " SELECT fccode, fcname FROM parameter WHERE fcba = '" &  Session("fcBA") & "' AND parcode = 'F0001'" & _
				    " UNION ALL" & _
				    " SELECT fccode, fcname FROM parameter WHERE fcba = '" &  Session("fcBA") & "' AND parcode = 'OMBROMETER_LOCATION'"
				Set rsComboA = Server.CreateObject("ADODB.Recordset")
				
				rsComboA.Open q, connstr
				if not rsComboA.eof then
					Do While Not rsComboA.eof
				%>
					<option value="<%=rsComboA("fccode")%>" <% If WIDIVISION = rsComboA("fccode") Then %> Selected <%End if%> ><%=rsComboA("fcname")%></option> 
				<% 
						rsComboA.movenext
					loop
				End if
				rsComboA.Close : set rsComboA = Nothing : q = ""
				%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<TR>
			<td align="left">Rain Inches</td>
			<td align="left"><input type="text" size="15" name="TOTRNINCHES" value="<%=TOTRNINCHES%>"></td>
		</tr>
		<tr>
			<td align="left">Rain Days</td>
			<td align="left"><input type="text" size="15" name="TOTRNDAYS" value="<%=TOTRNDAYS%>"></td>
		</tr>
		<tr>
			<td align="left">Temperature</td>
			<td align="left"><input type="text" size="15" name="AVGTEMP" value="<%=AVGTEMP%>"></td>
		</tr>
		<tr>
			<td align="left">Wind Speed</td>
			<td align="left"><input type="text" size="15" name="AVGWIND" value="<%=AVGWIND%>"></td>
		</tr>
	    <TR>
			<td colspan="2" align="Left">
				<input type="hidden" name="WIDATE" value="<%=WIDATE%>">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="WIDIVISION" value="<%=WIDIVISION%>">
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
   </form>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines" >
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE</TD>	
	<TD align="left">DIVISION</TD>	
	<TD align="left">RAIN INCHES</TD>	
	<TD align="left">RAIN DAYS</TD>	
	<TD align="left">TEMPERATURE</TD>	
	<TD align="left">WIND SPEED</TD>	
	</TR>
	</thead>
	<tbody>

        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-5
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">
	
			<a href="Weather.asp?WIDATE=<%=rsGAD("WIDATE")%>&txttable=<%=request("txttable")%>&WIDIVISION=<%=rsGAD("WIDIVISION")%>&TOTRNINCHES=<%=rsGAD("TOTRNINCHES")%>&TOTRNDAYS=<%=rsGAD("TOTRNDAYS")%>&AVGTEMP=<%=rsGAD("AVGTEMP")%>&AVGWIND=<%=rsGAD("AVGWIND")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="Weather.asp?WIDATE=<%=rsGAD("WIDATE")%>&txttable=<%=request("txttable")%>&WIDIVISION=<%=rsGAD("WIDIVISION")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
   </tbody>
   </TABLE>
   </div>
<%End if%>
</div></div>
<!--#include file="bottom.jsp"-->
</body>
</html>
