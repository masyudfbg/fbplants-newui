<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - Upkeep Rotation</title>
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
 <!-------------awal------------->
<%
Cek_IsDailyClosing = ""
 
txtYear = request("txtYear")
activitycode = request("activitycode")
groupactivitycode = request("groupactivitycode")
txtMonth = request("txtMonth")
fieldcode = request("fieldcode")
mandays = request("mandays")
runningnumber = request("runningnumber")
lineno = request("lineno")

Yr = Year(Session("v_Costbook"))
if "" & txtYear = "" then txtYear = Yr

function f_get_jobcode(year__, fieldcode__, job__)
	status__ = "NOSTATUS"
	result__ = "0000000"
	qMaxLineNo = " SELECT status" & _
				 " FROM budget_field" & _
				 " WHERE budget_year = '" & year__ & "'" & _
				 "       AND budget_field_version = '1'" & _
				 "       AND fccode = '" & fieldcode__ & "'" & _
				 "       AND fcba = '" & Session("fcBA") & "' "
	
	Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
	rsqMaxLineNo.Open qMaxLineNo, connstr

	if not rsqMaxLineNo.eof then
		status__ = rsqMaxLineNo("status")
	else
		status__ = "NOSTATUS"
	end if
	rsqMaxLineNo.close
	set rsqMaxLineNo = Nothing
	
	if status__ = "NOSTATUS" then
		result__ = "0000000"
	else
		qMaxLineNo = " SELECT param_1, param_2" & _
					 " FROM parameter" & _
					 " WHERE fccode = '" & job__ & "'" & _
					 "       AND parcode = 'UR001'" & _
					 "       AND fcba = '" & Session("fcBA") & "' "

		Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
		rsqMaxLineNo.Open qMaxLineNo, connstr

		if not rsqMaxLineNo.eof then
			param1__ = rsqMaxLineNo("param_1")
			param2__ = rsqMaxLineNo("param_2")
		else
			param1__ = "0000000"
			param2__ = "0000000"
		end if
		rsqMaxLineNo.close
		set rsqMaxLineNo = Nothing
		
		if cstr(""&status__) = "TBM" then
			result__ = param1__
		elseif cstr(""&status__) = "TM" then
			result__ = param2__
		else
			result__ = "0000000"
		end if
	end if
	
	f_get_jobcode = result__
end function


if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
			q = " DELETE FROM upkeeprotation WHERE fcba = '" & Session("fcBA") & "' AND lineno = '" & lineno & "'"

			' response.write q
			connStr.execute(q)

		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		has_fieldcode = 0 : lineno = 0 : has_record = 0 : p_jobcode = ""
		
		qMaxLineNo = " SELECT COUNT(1) totalrecord " & _
					 " FROM budget_field" & _
					 " WHERE budget_year = '" & txtYear & "'" & _
					 "       AND budget_field_version = '1'" & _
					 "       AND fccode = '" & fieldcode & "'" & _
					 "       AND fcba = '" & Session("fcBA") & "'"
						
		Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
		rsqMaxLineNo.Open qMaxLineNo, connstr

		if not rsqMaxLineNo.eof then
			has_fieldcode = cdbl(rsqMaxLineNo("totalrecord"))
		end if
		rsqMaxLineNo.close
		set rsqMaxLineNo = Nothing
		
		if has_fieldcode > 0 then
		
			qMaxLineNo = " SELECT COUNT(1) totalrecord " & _
						 " FROM upkeeprotation ur" & _
						 " LEFT JOIN parameter pr ON (pr.fcba = ur.fcba AND pr.parcode = 'UR001' AND pr.fccode = ur.activitycode)" & _
						 " WHERE ur.period_year = '" & txtYear & "'" & _
						 "       AND ur.fcba = '" & Session("fcBA") & "'" & _
						 "       AND pr.param_3 = '" & groupactivitycode & "'" & _
						 "       AND ur.fieldcode = '" & fieldcode & "'" & _
						 "       AND ur.period_month = '" & txtMonth & "'"
							
			Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
			rsqMaxLineNo.Open qMaxLineNo, connstr
			
			if not rsqMaxLineNo.eof then
				has_record = cdbl(rsqMaxLineNo("totalrecord"))
			end if
			rsqMaxLineNo.close
			set rsqMaxLineNo = Nothing
			
			if has_record = 0 then
				if cstr(""&p_jobcode) <> "0000000" then
			
					qMaxLineNo = " SELECT NVL(MAX(lineno),0) maxLineNo " & _
								 " FROM upkeeprotation " & _
								 " WHERE fcba = '" & Session("fcBA") & "' "
									
					Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
					rsqMaxLineNo.Open qMaxLineNo, connstr

					if not rsqMaxLineNo.eof then
						lineno = cdbl(to_nol(rsqMaxLineNo("maxLineNo")))
					end if
					rsqMaxLineNo.close
					set rsqMaxLineNo = Nothing
			
					p_jobcode = f_get_jobcode(txtYear,fieldcode,activitycode)
					
					qMaxLineNo = " SELECT COUNT(1) totalrecord " & _
								 " FROM upkeeprotation" & _
								 " WHERE period_year = '" & txtYear & "'" & _
								 "       AND fcba = '" & Session("fcBA") & "'" & _
								 "       AND activitycode = '" & activitycode & "'" & _
								 "       AND period_month = '" & txtMonth & "'" & _
								 "       AND SUBSTR(orderdata,1,1)||LPAD(TO_NUMBER(SUBSTR(orderdata,2,4)),4,0) = SUBSTR('" & runningnumber & "',1,1)||LPAD(TO_NUMBER(SUBSTR('" & runningnumber & "',2,4)),4,0)"
					'response.write qMaxLineNo				
					Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
					rsqMaxLineNo.Open qMaxLineNo, connstr

					if not rsqMaxLineNo.eof then
						has_runningnumber = cdbl(to_nol(rsqMaxLineNo("totalrecord")))
					end if
					rsqMaxLineNo.close
					set rsqMaxLineNo = Nothing
					
					if has_runningnumber = 0 then
			
						lineno = lineno +1
						q = " INSERT INTO upkeeprotation ( lineno" & _
							"                              , period_year" & _
							"                              , period_month" & _
							"                              , jobcode" & _
							"                              , fieldcode" & _
							"                              , mandays" & _
							"                              , orderdata" & _
							"                              , activitycode" & _
							"                              , fcentry" & _
							"                              , fcedit" & _
							"                              , fcip" & _
							"                              , fcba" & _
							"                              , lastupdate" & _
							"                              , lasttime)" & _
							" VALUES ( '" & lineno & "'" & _
							"          , '" & txtYear & "'" & _
							"          , '" & txtMonth & "'" & _
							"          , '" & p_jobcode & "'" & _
							"          , '" & fieldcode & "'" & _
							"          , '" & mandays & "'" & _
							"          , '" & runningnumber & "'" & _
							"          , '" & activitycode & "'" & _
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
						err.description = "Data Running Number "&runningnumber&" sudah ada, Silahkan perbaiki dan cek kembali inputan anda"
					end if
				else
					err.number = 1001
					err.description = "Status Blok tidak sesuai, Silahkan cek kembali inputan anda"
				end if
			else
				err.number = 1001
				err.description = "Data Upkeep, Blok:" & fieldcode & " Aktivitas:" & groupactivitycode & " Bulan: " & Bulan(txtMonth) & " sudah di entry, Silahkan cek kembali inputan anda"
			End If
		else
			err.number = 1001
			err.description = "Blok " & fieldcode & " tidak ada didalam budget blok, Silahkan cek kembali inputan anda"
		End If
			
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			q = " UPDATE upkeeprotation" & _
                " SET fieldcode = '" & fieldcode & "'" & _					
                "     , mandays = '" & mandays & "'" & _
                "     , orderdata = '" & runningnumber & "'" & _
                "     , fcedit = '" & Session("CodeName") & "'" & _
                "     , lastupdate = SYSDATE " & _
                "     , lasttime = TO_CHAR(SYSDATE, 'HH24:MI') "& _
                " WHERE fcba = '" & Session("fcBA") & "' AND lineno = '" & lineno & "'"
					
			' response.write q
			connStr.execute(q)

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
				" FROM upkeeprotation" & _
				" WHERE fcba = '" & Session("fcBA") & "'" & _
				"       AND lineno = '" & lineno & "'"
				
	set oRecordSet = server.CreateObject("Adodb.recordset")
	oRecordSet.Open oSQLQuery, connstr
	If Not oRecordSet.Eof then
		txtYear = oRecordSet("period_year")
		activitycode = oRecordSet("activitycode")
		txtMonth = oRecordSet("period_month")
		fieldcode = oRecordSet("fieldcode")
		mandays = oRecordSet("mandays")
		runningnumber = oRecordSet("orderdata")
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
            " from DUAL"

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

 <form name="frmEntry" method="post" action="UpkeepRotation.asp">
	<table class="data" style="width:70%;">
		<tr>
			<td align="left" style="width:20%;">Year</td>
			<td align="left">
				<select name="txtYear">
				<%
					MinYr = Yr -2
					MaxYr = Yr +2
					
					for xYr=MinYr to MaxYr
					if cint(txtYear) = cint(xYr) then
						response.write "<option value ='"&xYr&"' selected>"&xYr&"</option>"
					else
						response.write "<option value ='"&xYr&"'>"&xYr&"</option>"
					end if
					next
				%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Activity</td>
			<td align="left">
				<select name="activitycode">
				<%
					strsectionname	=	" SELECT fccode, fcname, param_3 from parameter where fcba = '" & Session("fcba") & "' AND parcode = 'UR001' order by fccode asc"
						
					Set rssectionname = Server.CreateObject("ADODB.Recordset")
					rssectionname.Open strsectionname, connstr, adOpenStatic 
					
					if not rssectionname.eof then
						Do While Not rssectionname.eof
							if cstr(""&activitycode) = cstr(""&rssectionname("fccode")) then
								groupactivitycode = cstr(""&rssectionname("param_3"))
								response.write "<option value ='"&rssectionname("fccode")&"' selected>"&rssectionname("fccode")&": "&rssectionname("fcname")&"</option>"
							else
								response.write "<option value ='"&rssectionname("fccode")&"'>"&rssectionname("fccode")&": "&rssectionname("fcname")&"</option>"
							end if
							rssectionname.movenext
						loop
					end if
				%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;
			</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" and "" & txtYear <> "" and "" & activitycode <> "" Then	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q = " SELECT ur.*" & _
		" FROM upkeeprotation ur" & _
		" LEFT JOIN job j ON (j.fcba = ur.fcba AND j.fccode = ur.jobcode)" & _
		" WHERE ur.period_year = '" & txtYear & "'" & _
        "       AND ur.activitycode = '" & activitycode & "'" & _
        "       AND ur.fcba = '" & session("fcBA") & "'" & _
        " ORDER BY ur.period_year, ur.jobcode, ur.period_month, ur.orderdata"
	'response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.Open q, connstr
%>
   <form action="UpkeepRotation.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" style="width:70%;">
		<!--Yang tidak bisa di edit adalah PK Table-->
		<tr>
			<td align="left" style="width:20%;">Month</td>
			<td align="left">
				<select name="txtMonth" <%IF Request("txtJenis")="Edit" then%>disabled style="background:#cccccc;"<%end if%>>
				<%
					
					for xMn=1 to 12
					if cint(txtMonth) = cint(xMn) then
						response.write "<option value ='"&xMn&"' selected>"&Bulan(xMn)&"</option>"
					else
						response.write "<option value ='"&xMn&"'>"&Bulan(xMn)&"</option>"
					end if
					next
				%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="Left">Field Code</td>
			<td align="Left">
				<input type="text" size="15" name="fieldcode" value="<%=fieldcode%>" <%IF Request("txtJenis")="Edit" then%>disabled style="background:#cccccc;"<%end if%>>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="Left">Mandays</td>
			<td align="Left">
				<input type="text" size="15" name="mandays" value="<%=mandays%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="Left">Running Number</td>
			<td align="Left">
				<input type="text" size="15" maxlength="5" id="runningnumber" name="runningnumber" value="<%=runningnumber%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
	    <TR><td colspan="2" align="Left">
	      <input type="hidden" name="txtYear" value="<%=txtYear%>">
	      <input type="hidden" name="activitycode" value="<%=activitycode%>">
	      <input type="hidden" name="groupactivitycode" value="<%=groupactivitycode%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%If val_IsTrue = "Y" Then%>
		  <%IF Request("txtJenis")="Edit" then%>
			<input type="hidden" name="txtMonth" value="<%=txtMonth%>">
			<input type="hidden" name="lineno" value="<%=lineno%>">
			<input type="hidden" name="fieldcode" value="<%=fieldcode%>">
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update" onclick="return validasiform();">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add" onclick="return validasiform();">
		  <%End if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	<TD align="left">No</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">Year</TD>
	<TD align="left">Jobcode</TD>
	<TD align="left">Month</TD>
	<TD align="left">Fieldcode</TD>
	<TD align="left">Mandays</TD>
	<TD align="left">Order Data</TD>
	</tr>
   
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
				<%If val_IsTrue = "Y" Then%>
					<td>
						<a href="UpkeepRotation.asp?lineno=<%=rsVH("lineno")%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0">
						</a>
					</td>
					<td>
						<a href="UpkeepRotation.asp?lineno=<%=rsVH("lineno")%>&txtYear=<%=rsVH("period_year")%>&activitycode=<%=rsVH("activitycode")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');">
							<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0">
						</a>
					</td>
				<%else%>
					<TD COLSPAN="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
				<%End if%>
				<td align="left" title="<%=rsVH("period_year").Name%>" valign="top"><%=rsVH("period_year")%></td>
				<td align="left" title="<%=rsVH("jobcode").Name%>" valign="top"><%=rsVH("jobcode")%></td>
				<td align="left" title="<%=rsVH("period_month").Name%>" valign="top"><%=rsVH("period_month")%></td>
				<td align="left" title="<%=rsVH("fieldcode").Name%>" valign="top"><%=rsVH("fieldcode")%></td>
				<td align="left" title="<%=rsVH("mandays").Name%>" valign="top"><%=rsVH("mandays")%></td>
				<td align="left" title="<%=rsVH("orderdata").Name%>" valign="top"><%=rsVH("orderdata")%></td>
			</tr>
		  <%
		  rsVH.MoveNext
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
<%end if%>
  <!-------------akhir------------->
</div></div>
  <!--#include file="bottom.jsp"-->
  <script type="text/javascript">
    function validasiform() {
	
		var rn = document.getElementById('runningnumber').value;
        if (rn === '') {
            alert("running number harus diisi");
            return false;
        } else {
			var str1;
			var str2;
			str1 = rn.substr(0, 1);
			//alert(str1);
			str2 = rn.substr(1, 2);
			if (str1 === "A" || str1 === "B" || str1 === "C" || str1 === "D" || str1 === "E" || str1 === "F" || str1 === "G" || str1 === "H" || str1 === "I" || str1 === "J" || str1 === "K" || str1 === "L") {
				if(isNaN(str2)){
					alert("digit kedua sampai keempat harus angka!");
					return false;
				}
			} else {
				alert("digit pertama harus huruf A-L!");
				return false;
			}
		}
        
    }
</script>
</body>
</html>
