<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - BBC Distribution</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
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
 <!-------------awal!------------->
 <%

Dim qSupplyingDeclaration, rsSupplyingDeclarationInfo, tempQuantity, tempDecType
Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, v_FCCODE, v_FCNAME
Dim qSupplyingInfo, rsSupplyingInfo, qDeclarationType, rsDeclarationType
Dim txtmonth, txtyear 
Dim parcode, totalstand, ScriptName
Dim QUANTITY, censuscode
Dim REMARKS
Dim TRANSDATE
Dim cmdView
Dim reference 
Dim distribution1, distribution2, distribution3, distribution4
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

txtyear =request("txtyear")
censuscode =request("censuscode")
quantity=request("quantity")
remarks=request("remarks")
fieldcode=request("fieldcode")
distribution1 = Request("distribution1")
distribution2 = Request("distribution2")
distribution3 = Request("distribution3")
distribution4 = Request("distribution4")

mon = Month(Session("v_Costbook"))
If len(txtmonth)=0 then txtmonth=mon

Yr = Year(Session("v_Costbook"))

If len(txtyear)= 0 then txtyear = Yr
transdate = txtmonth & "/1/" & txtyear

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   txtyear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   txtmonth	=	Month(Session("v_Costbook"))
end if


Function CheckFieldCode(initialFieldCode, initialFcba, initialDate)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from dual " & _
					" where field_api.GetValid_Activation('" & initialFieldCode & "', '" & initialFcba & "', to_date('"& initialDate &"', 'mm/dd/yyyy')) = 'Y'"

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

	CheckFieldCode = initialAvailable
End Function

Function CheckPrimary(initialFieldCode, initialFcba, initYear, initMonth)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	" select count(1) total" & _
	                " from supplyingprogram " & _
					" where fieldcode ='" & initialFieldCode & "'" & _
					" and year = '" & initYear & "'" & _
					" and month = '" & initMonth & "'" & _
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

Function CheckPlantingDate(initFieldCode, initFcba, initDate)
	Dim initialAvail
	Dim rsAvail, strAvail
	
	strAvail =	" select count (1) total" & _
	            " from dual " & _
	            " where to_char(field_api.GetValid_Plantingdate('" & initFieldCode & "', '" & initFcba & "', to_date('"& initDate &"', 'mm/dd/yyyy')),'yyyy')= to_char(to_date('"& initDate &"', 'mm/dd/yyyy'),'yyyy')"
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		initialAvail = Cdbl(rsAvail("Total"))
	Else
		initialAvail = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	CheckPlantingDate = initialAvail
End Function

Function F_CurrentTotalStand(initFieldCode, initFcba, initDate)
	Dim initialAvail
	Dim rsAvail, strAvail
	
	strAvail =	" select field_api.getValid_POKOK('" & initFieldCode & "', '" & initFcba & "', to_date('"& initDate &"', 'mm/dd/yyyy')) total" & _
	            " from dual " 
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		initialAvail = Cdbl(rsAvail("Total"))
	Else
		initialAvail = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	F_CurrentTotalStand = initialAvail
End Function

If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then

			q = " delete from bbcdistribution " & _
						" 		where periodyear = '" & txtyear & "'" & _
						"       and censuscode = '" & censuscode & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
				' Response.Write q
				connStr.execute(q)
				cmdView	= "True"

	elseif Request("txtJenis")="New" then		
			If Len(Trim(censuscode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Census Code Harus diisi, tidak boleh kosong"
			Else

			tempdistribution = cdbl(to_nol(distribution1)) + cdbl(to_nol(distribution2)) + cdbl(to_nol(distribution3)) + cdbl(to_nol(distribution4)) = 100
			If tempdistribution = 0 Then
				Err.Number = 1001
				Err.Description = "Total distribution harus sama dengan 100."
			Else
				q = " INSERT INTO bbcdistribution" & _
					" (periodyear, censuscode, distribution1, distribution2, distribution3, distribution4, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) VALUES (" & _
					" '" & txtyear & "'" & _
					" , '" & censuscode & "'" & _
					" , '" & distribution1 & "'" & _
					" , '" & distribution2 & "'" & _
					" , '" & distribution3 & "'" & _
					" , '" & distribution4 & "'" & _
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
	cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			If Len(Trim(censuscode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Census Code Harus diisi, tidak boleh kosong"
			Else

				tempdistribution = cdbl(to_nol(distribution1)) + cdbl(to_nol(distribution2)) + cdbl(to_nol(distribution3)) + cdbl(to_nol(distribution4)) = 100
				If tempdistribution = 0 Then
					Err.Number = 1001
					Err.Description = "Total distribution haarus sama dengan 100."
				Else
					q =	" Update bbcdistribution" & _
						" set  distribution1 = '" & distribution1 & "'" & _
						"     , distribution2 = '" & distribution2 & "'" & _
						"     , distribution3 = '" & distribution3 & "'" & _
						"     , distribution4 = '" & distribution4 & "'" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where periodyear = '" & txtyear & "'" & _
						" 		and censuscode = '" & censuscode & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
		 			' response.write q
					connStr.execute(q)
				End If
			End If
		cmdView	= "True"
	else
		cmdView	= "True"
 	End if
	Else
	 	cmdview = request("cmdview")
 	End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")


If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('"& txtmonth &"'||'/01/'||'" & txtyear &"', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_char(to_date('"& txtmonth &"'||'/01/'||'" & txtyear &"','mm/dd/yyyy'),'yyyy')= to_char(currentdate,'yyyy')" & _
	        "       and upper(system) = upper('costbook')"
			
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if

'response.Write strNurseryCulling
oRecordSet.Close
Set oRecordSet = Nothing

%>

 <!-------------ISI Supplying Declaration Form!------------->
 
	<form name="frmEntry" method="post" action="<%=ScriptName%>">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Year</td>
				<% 
					Dim mon, Yr, MinYr, MaxYr, MM

					MinYr = Yr - 2
					MaxYr = Yr  + 2
				%>
			<td align="left">
				<select name="txtYear">
				<%
					For i = MinYr to MaxYr
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(txtyear) then%> Selected <%End If%>><%=i%></option>
				<%
					Next
				%>
				</select>
			</td>
		</tr>		
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				<!-- <img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> -->
				<!-- HADIR -->
				<!-- <%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Nursery"
				dataSubmodule 	= "Culling"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a> -->
				<!-- END:HADIR -->
			</td>
		</tr>
		</table>
	</form>
<%
if cmdView = "True" and len(trim(txtyear)) > 0 Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= " SELECT * FROM bbcdistribution " & _
		  " WHERE fcba = '" & session("fcBA") & "'" & _
		  " AND periodyear = '" & txtyear & "'" 

	'response.Write q

	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<form action="bbcdistribution.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
			<tr> 
	            <td align="left"><font face="Verdana" size="1">Census Code</font></td>
					<td align="left"> 
						<select name="censuscode" required="required" style="background:#FFFFCC;" title="CENSUS CODE Character a s/d z or 1 s/d 9">
							<option value=""></option> 
								<% 	Dim rsComboV
								q = "SELECT * FROM PARAMETER WHERE PARCODE = 'BBC' and fcba = '" & Session("fcba") & "'"
								Set rsComboV = Server.CreateObject("ADODB.Recordset")
								rsComboV.CursorType = 3
								rsComboV.CursorLocation = 3
								rsComboV.Open q, connstr
								if not rsComboV.eof then%>
								<%
								Do While Not rsCombov.eof
								%>
									<option value="<%=rsComboV("FCCODE")%>" 
										<%if cstr(""&censuscode)=cstr(""&rsComboV("FCCODE")) then %> selected 
										<%End If%>><%=rsComboV("FCNAME")%>
									</option> 
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
				<td align="left" width="20%">Distribution 1</td>
				<td align="left">
					<input type="text" size="15" name="distribution1" style="background:#FFFFCC;" value="<%=distribution1%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
					<font color="#FF0000">%</font>
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Distribution 2</td>
				<td align="left">
					<input type="text" size="15" name="distribution2" style="background:#FFFFCC;" value="<%=distribution2%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
					<font color="#FF0000">%</font>
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Distribution 3</td>
				<td align="left">
					<input type="text" size="15" name="distribution3" style="background:#FFFFCC;" value="<%=distribution3%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
					<font color="#FF0000">%</font>
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Distribution 4</td>
				<td align="left">
					<input type="text" size="15" name="distribution4" style="background:#FFFFCC;" value="<%=distribution4%>" autocomplete="off" maxlength="32" onkeyup="hanyaAngka(this);" >
					<font color="#FF0000">%</font>
				</td>
			</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="txtyear" value="<%=txtyear%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<%If val_IsTrue = "Y" Then%>
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtJenis" value="Edit">
						<input Type="Submit"  name="Add" value="Update">
					<%Else%>
						<input type="hidden" name="txtJenis" value="New">
						<input Type="Submit"  name="Add" value="Add">
					<%End if%>

					<input Type="Reset"  name="Reset" value="Reset">
					<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				<%End if%>
				</td>
			</tr>
		</table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="1" >&nbsp;</td>
		<td align="left">CENSUS CODE</td>
		<td align="left">DISTRIBUTION 1</td>
		<td align="left">DISTRIBUTION 2</td>
		<td align="left">DISTRIBUTION 3</td>
		<td align="left">DISTRIBUTION 4</td>
	</tr>
   
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
				<%If val_IsTrue = "Y" Then%>
					<a href="bbcdistribution.asp?txtyear=<%=txtyear%>&censuscode=<%=rsVH("censuscode")%>&distribution1=<%=rsVH("distribution1")%>&distribution2=<%=rsVH("distribution2")%>&distribution3=<%=rsVH("distribution3")%>&distribution4=<%=rsVH("distribution4")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>

					<a href="bbcdistribution.asp?txtyear=<%=txtyear%>&censuscode=<%=rsVH("censuscode")%>&distribution1=<%=rsVH("distribution1")%>&distribution2=<%=rsVH("distribution2")%>&distribution3=<%=rsVH("distribution3")%>&distribution4=<%=rsVH("distribution4")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtyear")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				<% else %>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> 
				<%end if%>
				
			</td>
			    <td align="left" title="<%=rsVH("censuscode").Name%>" valign="top"><%=rsVH("censuscode")%></td>
			    <td align="left" title="<%=rsVH("distribution1").Name%>" valign="top"><%=rsVH("distribution1")%></td>
			    <td align="left" title="<%=rsVH("distribution2").Name%>" valign="top"><%=rsVH("distribution2")%></td>
			    <td align="left" title="<%=rsVH("distribution3").Name%>" valign="top"><%=rsVH("distribution3")%></td>
			    <td align="left" title="<%=rsVH("distribution4").Name%>" valign="top"><%=rsVH("distribution4")%></td>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
