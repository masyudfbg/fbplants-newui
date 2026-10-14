<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtFcba, v_txtYear, v_txtMonth,v_txtDay,v_txtDate, v_txtAfdeling, v_txtBlok, v_txtColor
Dim ScriptName, v_weekid
Dim rsPusingan, strsql
Dim NO, cClass, nday,ColorID, varDefaultWeek
Dim rsDefaultBa, strDefaultBa, varDefaultBa, strDefaultWeek, rsDefaultWeek
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName,varCompanyCode
Dim oRS, oSQL

datefrom = Request("datefrom")
dateuntil = Request("dateuntil")

If "" & datefrom = "" then datefrom = Month(TextDate(date)) &"/1/" & Year(TextDate(date))
If "" & dateuntil = "" then dateuntil = TextDate(date)


strweek = " SELECT NVL(RPT_PIEZOTRANSACTION_API.filter_date(TO_DATE('" & datefrom & "', 'mm/dd/yyyy'), TO_DATE('" & dateuntil & "', 'mm/dd/yyyy')), '''W3092016''') result FROM DUAL"
Set rsweek = Server.CreateObject("ADODB.Recordset")
rsweek.Open strweek, connstr, adOpenStatic
if not rsweek.eof then
	listweek = rsweek("result")
else
	listweek = "'W3092016'"
end if
rsweek.Close
Set rsweek = Nothing

'response.write listweek

strweek = " SELECT weekid FROM weekmaster_tab WHERE fromdate >= TO_DATE('" & datefrom & "', 'mm/dd/yyyy') AND todate <= TO_DATE('" & dateuntil & "', 'mm/dd/yyyy') ORDER BY fromdate"
arr_week_eof = true
Set rsweek = Server.CreateObject("ADODB.Recordset")
rsweek.Open strweek, connstr, adOpenStatic
if not rsweek.eof then
	arr_week = rsweek.getrows()
	arr_week_eof = false
end if
rsweek.Close
Set rsweek = Nothing

strweek = " SELECT colour, lowlevel, maxlevel FROM waterlevelstatus_tab"
arr_color_eof = true
Set rsweek = Server.CreateObject("ADODB.Recordset")
rsweek.Open strweek, connstr, adOpenStatic
if not rsweek.eof then
	arr_color = rsweek.getrows()
	arr_color_eof = false
end if
rsweek.Close
Set rsweek = Nothing

Function span_yearmonth(data_array__, yearmonth__)
	result_ = 0
	arr_1 = data_array__
	for indx = LBound(arr_1, 2) to UBound(arr_1, 2)
		if Mid(arr_1(0, indx),3,6) = Mid(yearmonth__,3,6) then
			result_ = result_ + 1
		end if
	next
	span_yearmonth = result_
End Function

Function desc_yearmonth(yearmonth__)
	result_ = ""
	
	if Mid(yearmonth__,3,2) = "01" then
		result_ = "Jan"
	elseif Mid(yearmonth__,3,2) = "02" then
		result_ = "Feb"
	elseif Mid(yearmonth__,3,2) = "03" then
		result_ = "Mar"
	elseif Mid(yearmonth__,3,2) = "04" then
		result_ = "Apr"
	elseif Mid(yearmonth__,3,2) = "05" then
		result_ = "Mei"
	elseif Mid(yearmonth__,3,2) = "06" then
		result_ = "Jun"
	elseif Mid(yearmonth__,3,2) = "07" then
		result_ = "Jul"
	elseif Mid(yearmonth__,3,2) = "08" then
		result_ = "Agu"
	elseif Mid(yearmonth__,3,2) = "09" then
		result_ = "Sep"
	elseif Mid(yearmonth__,3,2) = "10" then
		result_ = "Okt"
	elseif Mid(yearmonth__,3,2) = "11" then
		result_ = "Nov"
	elseif Mid(yearmonth__,3,2) = "12" then
		result_ = "Des"
	end if
	
	result_ = result_ & " " & Mid(yearmonth__,5,4)
	
	desc_yearmonth = result_
End Function

Function coloring(data_array__, waterlevel__)
	result_ = ""
	arr_1 = data_array__
	if "" & waterlevel__ = "" then
		result="fff"
	else
	for indx = LBound(arr_1, 2) to UBound(arr_1, 2)
		if Cdbl(To_Nol(waterlevel__)) >= Cdbl(To_Nol(arr_1(1, indx))) and Cdbl(To_Nol(waterlevel__)) <= Cdbl(To_Nol(arr_1(2, indx))) then
			result_ = arr_1(0, indx)
		end if
	next
	end if
	
	result_ = "#" & result_
	
	coloring = result_
End Function

v_weekid		=	Request("txtweekid")

v_txtFcba = Request("txtFcba")


strDefaultBa = " SELECT"& _
				"     BU.FCCODE as FCBA,"& _
				"     BU.FCNAME as FCBANAME," & _
			    "     BU.fccompanycode  , BU.FCCOMPANYNAME as COMPANYNAME" & _
				" FROM BUSINESSUNIT BU"& _
				" WHERE "& _
				"     BU.FCCODE = '" & v_txtFcba & "'"
'response.write strDefaultBa
				
Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
	varFcbaName = rsDefaultBa("FCBANAME")
	varCompanyName = rsDefaultBa("COMPANYNAME")
	varCompanyCode = rsDefaultBa("fccompanycode")
Else
	varDefaultBa = ""
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing


If "" & v_txtFcba = "" Then v_txtFcba = varDefaultBa

if "" & v_txtFcba <> "" and v_txtFcba <> "ALL" then
	wherefcba = "           AND pt.fcba = '" & v_txtFcba & "'"
end if
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
Dim strPusingan, titleSite, wherecons, v_txtfcbacondition, v_weekidcondition
strPusingan	= " SELECT * FROM (" & _
              "     SELECT businessunit_api.get_plantation(pt.fcba) Plantation" & _
              "            , businessunit_api.get_Groupregion(pt.fcba) Region" & _
              "            , businessunit_api.get_Region(pt.fcba) groupkebun" & _
              "            , businessunit_api.get_CompanyCode(pt.fcba) Company" & _
              "            , pt.fcba" & _
              "            , field_api.getValid_Afdeling(fieldcode, pt.fcba, sysdate) afdeling" & _
              "            , p.fieldcode" & _
              "            , field_api.getValid_plantingyear(fieldcode, pt.fcba, sysdate) TT" & _
              "            , field_api.GetValid_hectarageplanted(fieldcode, pt.fcba, sysdate) Ha" & _
              "            , pt.weekid, pt.waterlevel" & _
              "     FROM PIEZOTRANSACTION_TAB pt" & _
              "     LEFT JOIN piezo_tab p ON (pt.fcba = p.fcba AND pt.piezoid = p.piezoid)" & _
              "     WHERE pt.weekid IN (" & listweek & ")" & _
              " " & wherefcba & "" & _
              "           AND p.ISACTIVE = 'TRUE'          " & _
              " ) PIVOT" & _
              " (" & _
              "   MAX(waterlevel)" & _
              "   FOR weekid IN (" & listweek & ")" & _
              " )" & _
              " ORDER BY fcba, afdeling, fieldcode"
			  
			  
			  
			    if v_weekid <> "ALL"  then
				'v_weekidcondition = " and weekid = '"& v_weekid &"'"
				end if	
				if v_txtfcba <> "ALL" then
				'v_txtfcbacondition = " where  fcba = '"& v_txtfcba & "'"
				end if	
				
	whereCons = " "	

'response.write strPusingan & whereCons

arr_data_eof = true
Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan & whereCons , connstr, adOpenStatic

if not rsPusingan.eof then
	arr_data = rsPusingan.getrows()
	arr_data_eof = false
end if
rsPusingan.Close
Set rsPusingan = Nothing

if not arr_data_eof then
	actualcol = ubound(arr_data,1)
	spanstatuspiezo = actualcol -8
	totalcol = 2
	throwspan = 3
else
	actualcol = 10
	throwspan = 2
end if
'response.write totalcol
%>


<html>

<head>
<title>ePlaS Reporting - Summary Weekly Piezo Meter</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
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
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>

<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>


<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>

	<div id="wrapper">
	 <!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
				<div class="printhide">
					<table id="form-table" class="centered-margin">	
						<tr>
							<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<font face="verdana" size="1">FCBA</font>
							</td>
							<td>
								<select name="txtFcba" >
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT FCCODE, FCNAME FROM BUSINESSUNIT"& _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL"& _
												" AND FCTYPE = 'E'"& _
												" ORDER BY FCCODE"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									
									If rsFcba.Bof Or rsFcba.eof Then
									%>
									<option value="ALL">--ALL--</option>
									<%
									End If
									
                                    if not rsFcba.eof then %>
									<option value="ALL">--ALL--</option>
									<%
									Do While Not rsFcba.eof
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if v_txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
									<%
										rsFcba.movenext
									loop
									end if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">From</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" maxlength="10" name="datefrom" size="10" value="<%=datefrom%>" title="Type Date mm/dd/yyyy" >
							</td>
							<td>
								<font face="verdana" size="1">To</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" maxlength="10" name="dateuntil" size="10" value="<%=dateuntil%>" title="Type Date mm/dd/yyyy" >
							</td>
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table width="100%" border="0" bordercolor="#CFA4BC" class="xthf-site" id="product-table">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=actualcol + totalcol%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>Summary Weekly Piezo Meter<h2>
	   <%
		dim rskbn,kbn
		
		kbn = "SELECT FCCODE, FCNAME FROM BUSINESSUNIT WHERE FCCODE = '" & v_txtFcba & "'"
		
	    Set rskbn = Server.CreateObject("ADODB.Recordset")
		rskbn.Open kbn, connstr, adOpenStatic

		If Not rskbn.EOF Then
	   %>													
				<h4>KEBUN : <%=rskbn("FCNAME")%></h4><br>
		<%  End If %>
			<%if datefrom = dateuntil then%>
				<h5>PERIODE : <%=dateuntil%></h5><br>
			<%else%>
				<h5>PERIODE : <%=datefrom%> - <%=dateuntil%></h5><br>
			<%end if%>
			</div>
		 </th>
	</tr>
<% 
dim selec_day
	if len(trim(DAY(v_txtDate))) = 0 then
		selec_day = GetLastDay(v_txtMonth, v_txtYear)
	else
		selec_day = DAY(v_txtDate)
	end if
%>
	<tr>
		<th width="15" rowspan="<%=throwspan%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>No</strong></div></th>
		<th width="178" rowspan="<%=throwspan%>" bgcolor="#D0A8C0" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Plantation</strong></div></th>
		<th width="68" rowspan="<%=throwspan%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Region</strong></div></th>
		<th width="130" rowspan="<%=throwspan%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>GroupKebun</strong></div></th>
		<th width="95" rowspan="<%=throwspan%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Company</strong></div></th>
		<th width="67" rowspan="<%=throwspan%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>FCBA</strong></div></th>
		<th width="102" rowspan="<%=throwspan%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Afdeling</strong></div></th>
		<th width="103" rowspan="<%=throwspan%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>FieldCode</strong></div></th>
		<th width="29" rowspan="<%=throwspan%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>TT</strong></div></th>
		<th width="40" rowspan="<%=throwspan%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Luas</strong></div></th>
	<%if not arr_data_eof then%>
		<th colspan="<%=spanstatuspiezo%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Status Piezo Meter </strong></div></th>

	</tr>
  <tr>
	<%if not arr_week_eof then%>
		<%for ii=0 to ubound(arr_week,2)%>
			<%curr_yearmonth = Mid(arr_week(0,ii),3,6)%>
			<%if curr_yearmonth <> prev_yearmonth then%>
				<th colspan="<%=span_yearmonth(arr_week, arr_week(0,ii))%>" style="background-color:#00A300;color:#fff;"><div align="center"><strong><%=desc_yearmonth(arr_week(0,ii))%></strong></div></th>
			<%end if%>
	<!--
    <th colspan="4" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Okt 2016</strong></div></th>
    <th colspan="5" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Nov 2016</strong></div></th>
    <th colspan="4" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Des 2016</strong></div></th>
	<th colspan="4" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Jan 2017</strong></div></th>
	<th colspan="4" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Feb 2017</strong></div></th>
	<th colspan="5" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Mar 2017</strong></div></th>
	-->
			<%prev_yearmonth = curr_yearmonth%>
		<%next%>
	<%end if%>
  </tr>
  <tr>
	<%if not arr_week_eof then%>
		<%for ii=0 to ubound(arr_week,2)%>
			<th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;">
				<a target="_blank" href="http://172.30.1.122/pzo/Content/PZO_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&IDWEEK=<%=arr_week(0,ii)%>">
					<div align="center"><font color="white"><strong><%="M" & Mid(arr_week(0,ii),2,1)%></strong></font></div>
				</a>
			</th>
		<%next%>
	<%end if%>
  </tr>
<%end if%>
  </thead></tbody>
	
<%
if not arr_data_eof then
	NO = 1
	for xrow=0 to ubound(arr_data,2)
		response.write "<tr>"
		response.write "<td align='center'>" & NO & "</td>"
		for xcol=0 to ubound(arr_data,1)
			if xcol > 8 then
				if not arr_color_eof then
					bg_color = coloring(arr_color, arr_data(xcol,xrow))
				else
					bg_color = "#fff"
				end if
			else
				bg_color = "#fff"
			end if
%>
			<td style="background-color:<%=bg_color%>;" align="center"><%=arr_data(xcol,xrow)%></td>
<%
		next
		response.write "</tr>"
		NO = NO +1
	next
else
	response.write "<tr><td colspan='36'>No Data Available</td></tr>"
end if

Dim StrAkhir, rsAkhir, whereCons1

StrAkhir = " SELECT orderby,upper(status) status,Colour, count(status) jml, round(100*(count(*) / sum(count(*)) over ()),1) perc" & _
           " FROM ( SELECT fcba," &_
           "               waterlevelstatus_api.GET_Status(waterlevel) Status," &_ 
           "               waterlevelstatus_api.GET_Colour(waterlevel) colour, " &_ 
           "               waterlevelstatus_api.GET_LevelId(waterlevel) orderby " &_ 			  
           "        FROM PIEZOTRANSACTION_TAB  " & _
           "        WHERE weekid IN (" & listweek & ")"
if v_txtfcba <> "ALL" then
	StrAkhir = StrAkhir & " AND fcba = '"& v_txtfcba & "'"
end if
StrAkhir = StrAkhir & " )"


	
	
'response.write StrAkhir & " Group by Status,Colour order by orderby" 
Set rsAkhir = Server.CreateObject("ADODB.Recordset")
rsAkhir.Open StrAkhir & " Group by orderby,Status,Colour order by orderby" , connstr, adOpenStatic

dim Nilai
%>
</tbody>

</table>
<div align="center">

<table width="35%">
<tr></tr><tr></tr><tr></tr><tr></tr><tr></tr>
<tr><td colspan="4"><div align="left"><strong>Index</strong></div></td></tr>
<%  Do while not rsAkhir.eof        
      
%>
       <tr>
	       <% if  "FFFFFF" = rsAkhir("COLOUR") then  %>
	             <td >999</td>
		    <% else %>
			     <td width="3" bgcolor="<%=rsAkhir("COLOUR")%>"</td>
			<% End If %>
	       <td width="10">&nbsp&nbsp:</td>
		   <td width="15"><%=rsAkhir("STATUS")%></td>
		   <td width="45" align="right">( <%=rsAkhir("perc")%> %)</td>
	   </tr>
<%
	  rsAkhir.MoveNext            		
	loop
rsAkhir.Close
set rsAkhir = Nothing
%>
</table>
</div>

   <table width="90%" class="data">
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <%
		DIM rsttd,ttd
		
		ttd = "SELECT FCMANAGER,FCKTU FROM BusinessUnit WHERE FCCODE = '" & v_txtFcba & "'"
		
	    Set rsttd = Server.CreateObject("ADODB.Recordset")
		rsttd.Open ttd, connstr, adOpenStatic

		If Not rsttd.EOF Then
	   %>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCKTU")%></u></strong></td>
	   </tr>
	<%  End If
%> 
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
</body>
   </html>
