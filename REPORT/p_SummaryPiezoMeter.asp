<%@Language=VBScript%>
<% Option Explicit %>
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
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName
Dim oRS, oSQL

If Len(Trim(Request("txtDate"))) = 0 Then
	v_txtDate		=	date
else
	v_txtDate		=	Request("txtDate")
End If

v_weekid		=	Request("txtweekid")
'response.Write v_weekid
v_txtFcba		=	Request("txtFcba")
'response.write v_txtFcba

strDefaultBa = " SELECT"& _
				"     FCCODE as FCBA"& _
				" FROM BUSINESSUNIT BU"& _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) "& _
				" WHERE"& _
				"     ISDEFAULT = 'TRUE'"& _
				"     AND FCTYPE = 'E'"& _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing


If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
Dim strPusingan, titleSite, wherecons
strPusingan	= "select businessunit_api.get_plantation(fcba) Plantation, businessunit_api.get_Groupregion(fcba) Region,"& _
			  " businessunit_api.get_Region(fcba) groupkebun, businessunit_api.get_CompanyCode(fcba) Company,"& _
			  " fcba,field_api.getValid_Afdeling(fieldcode, fcba, sysdate) afedling,"& _
			  " fieldcode,field_api.getValid_plantingyear(fieldcode, fcba, sysdate) TT, Piezoid,"& _
			  " field_api.GetValid_hectarageplanted(fieldcode, fcba, sysdate) Ha,W4092016Status,"& _
			  " W4092016waterlevel, W4092016colour, W1102016Status, W1102016waterlevel, W1102016colour, W2102016Status,"& _
			  " W2102016waterlevel, W2102016colour, W3102016Status, W3102016waterlevel, W3102016colour,"& _
			  " W4102016Status, W4102016waterlevel, W4102016colour, W1112016Status, W1112016waterlevel, W1112016colour,"& _
			  " W2112016Status, W2112016waterlevel, W2112016colour, W3112016Status, W3112016waterlevel,"& _
			  " W3112016colour, W4112016Status, W4112016waterlevel, W4112016colour, W5112016Status, W5112016waterlevel,"& _
			  " W5112016colour, W1122016Status, W1122016waterlevel, W1122016colour, W2122016Status, W2122016waterlevel,"& _
			  " W2122016colour, W3122016Status, W3122016waterlevel, W3122016colour, W4122016Status, W4122016waterlevel,"& _
			  " W4122016colour from (select fcba, piezoid, piezo_api.GET_FIELDCODE(fcba, Piezoid) fieldcode, max(case weekid"& _
              " when 'W4092016' then status end) W4092016Status, max(case weekid when 'W4092016' then waterlevel"& _
              " end) W4092016waterlevel, max(case weekid when 'W4092016' then WaterlevelStatus_Api.GET_Colour(waterlevel)"& _
              " end) W4092016colour, max(case weekid when 'W1102016' then status end) W1102016Status,"& _
              " max(case weekid when 'W1102016' then waterlevel end) W1102016waterlevel,"& _
              " max(case weekid when 'W1102016' then WaterlevelStatus_Api.GET_Colour(waterlevel)"& _
              " end) W1102016colour, max(case weekid when 'W2102016' then"& _
              " status end) W2102016Status, max(case weekid when 'W2102016' then waterlevel"& _
              " end) W2102016waterlevel, max(case weekid when 'W2102016' then WaterlevelStatus_Api.GET_Colour(waterlevel)end) W2102016colour,"& _
              " max(case weekid when 'W3102016' then status end) W3102016Status, max(case weekid when 'W3102016' then waterlevel end) W3102016waterlevel,"& _
              " max(case weekid when 'W3102016' then WaterlevelStatus_Api.GET_Colour(waterlevel) end) W3102016colour, max(case weekid when 'W4102016' then"& _
              " status end) W4102016Status, max(case weekid when 'W4102016' then waterlevel end) W4102016waterlevel,"& _
              " max(case weekid when 'W4102016' then WaterlevelStatus_Api.GET_Colour(waterlevel) end) W4102016colour,"& _
              " max(case weekid when 'W1112016' then status end) W1112016Status, max(case weekid when 'W1112016' then"& _
              " waterlevel end) W1112016waterlevel, max(case weekid when 'W1112016' then WaterlevelStatus_Api.GET_Colour(waterlevel) end) W1112016colour,"& _
              " max(case weekid when 'W2112016' then status end) W2112016Status, max(case weekid when 'W2112016' then waterlevel end) W2112016waterlevel,"& _
              " max(case weekid when 'W2112016' then WaterlevelStatus_Api.GET_Colour(waterlevel) end) W2112016colour, max(case weekid when 'W3112016' then"& _
              " status end) W3112016Status, max(case weekid when 'W3112016' then waterlevel end) W3112016waterlevel, max(case weekid when 'W3112016' then"& _
              " WaterlevelStatus_Api.GET_Colour(waterlevel) end) W3112016colour, max(case weekid when 'W4112016' then status end) W4112016Status, max(case weekid"& _
              " when 'W4112016' then waterlevel end) W4112016waterlevel, max(case weekid when 'W4112016' then WaterlevelStatus_Api.GET_Colour(waterlevel)"& _
              " end) W4112016colour, max(case weekid when 'W5112016' then status end) W5112016Status, max(case weekid when 'W5112016' then waterlevel"& _
              " end) W5112016waterlevel, max(case weekid when 'W5112016' then WaterlevelStatus_Api.GET_Colour(waterlevel) end) W5112016colour, max(case weekid"& _
              " when 'W1122016' then status end) W1122016Status, max(case weekid when 'W1122016' then waterlevel end) W1122016waterlevel, max(case weekid"& _
              " when 'W1122016' then WaterlevelStatus_Api.GET_Colour(waterlevel)end) W1122016colour, max(case weekid when 'W2122016' then status end) W2122016Status,"& _
              " max(case weekid when 'W2122016' then waterlevel end) W2122016waterlevel, max(case weekid when 'W2122016' then WaterlevelStatus_Api.GET_Colour(waterlevel)"& _
              " end) W2122016colour, max(case weekid when 'W3122016' then status end) W3122016Status, max(case weekid when 'W3122016' then waterlevel end) W3122016waterlevel,"& _
              " max(case weekid when 'W3122016' then WaterlevelStatus_Api.GET_Colour(waterlevel) end) W3122016colour, max(case weekid when 'W4122016' then status"& _
              " end) W4122016Status, max(case weekid when 'W4122016' then waterlevel end) W4122016waterlevel, max(case weekid when 'W4122016' then WaterlevelStatus_Api.GET_Colour(waterlevel)"& _
              " end) W4122016colour from PIEZOTRANSACTION_TAB "
	wherecons = " where weekid = '"& v_weekid &"' and PIEZO_API.ISACTIVE(fcba, Piezoid) = 'TRUE' group by fcba, piezoid) where fcba = '"& v_txtfcba & "'" 
			    if v_weekid = "ALL"  then
			    whereCons = "group by fcba, piezoid) order by  fcba,field_api.getValid_Afdeling(fieldcode, fcba, sysdate) ,fieldcode "	
				end if	  

'response.write strPusingan & whereCons
Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan & whereCons , connstr, adOpenStatic


%>


<html>

<head>
<title>ePlaS Reporting - Summary Weekly Piezzo Meter</title>
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
						<!--<td><font face="verdana" size="1">TAHUN</font></td>
					<% 
						Dim mon, Yr, MinYr, MaxYr, i, MM
						mon = Month(Session("v_Costbook"))
						If len(v_txtMonth)=0 then v_txtMonth=mon

						Yr = Year(Session("v_Costbook"))
						MinYr = Yr - 2
						MaxYr = Yr  + 2
						If len(v_txtYear)= 0 then v_txtYear = Yr
					%>
							<td width="4">
								<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
								</select>
							</td>
							<td><font face="verdana" size="1">BULAN</font></td>
							<td width="2">
								<select name="txtMonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td>-->
							<td>
								<!-- <font face="verdana" size="1">DATE</font> -->
							</td>
							<td width="2">
								<!--<input class="datepicker" type="text" name="txtDate" value="<%'=v_txtDate%>" /> -->
							</td>
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
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT"& _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL"& _
												" AND FCTYPE = 'E'"& _
												" ORDER BY ORDERDATA"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									
									If rsFcba.Bof Or rsFcba.eof Then
									%>
									<option value=""></option>
									<%
									End If
									
									if not rsFcba.eof then
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
								<font face="verdana" size="1">MINGGU KE-</font>
							</td>
							<td>
								<select name="txtweekid" >
								
								    <%
									Dim rsWeek, strWeek, v_txtWeek
									strWeek	=	" SELECT * FROM WEEKMASTER_TAB ORDER BY todate "
									
									Set rsWeek = Server.CreateObject("ADODB.Recordset")
									rsWeek.Open strWeek, connstr, adOpenStatic
									
									If rsWeek.Bof Or rsWeek.eof Then
									%>
									<option value="ALL"> --ALL-- </option>
									<%
									End If
								    if not rsWeek.eof then %>
									<option value="ALL"> --ALL-- </option>
									<%
									Do While Not rsWeek.eof
									%>
									<option value="<%=rsWeek("WEEKID")%>" <%if v_txtWeek = rsWeek("WEEKNAME") then%> Selected <%End If%>><%=rsWeek("WEEKNAME")%></option>
									<%
										rsWeek.movenext
									loop
									end if
									rsWeek.Close
									set rsWeek = Nothing
									%>
							   </select>
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
         <th width="178" colspan="<%=6+GetLastDay(v_txtMonth, v_txtYear)%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>
				 Summary Weekly Piezo Meter <br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				<h4>s/d Tgl&nbsp;:&nbsp;<%=v_txtDay%>&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;&nbsp;<%=v_txtYear%></h4>
		   </div>
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
    <th width="15" rowspan="3" style="background-color:#00A300;color:#fff;"><div align="center"><strong>No</strong></div></th>
	<th width="178" rowspan="3" bgcolor="#D0A8C0" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Plantation</strong></div></th>
    <th width="68" rowspan="3" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Region</strong></div></th>
    <th width="130" rowspan="3" style="background-color:#00A300;color:#fff;"><div align="center"><strong>GroupKebun</strong></div></th>
    <th width="95" rowspan="3" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Company</strong></div></th>
    <th width="67" rowspan="3" style="background-color:#00A300;color:#fff;"><div align="center"><strong>FCBA</strong></div></th>
    <th width="102" rowspan="3" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Afdeling</strong></div></th>
    <th width="103" rowspan="3" style="background-color:#00A300;color:#fff;"><div align="center"><strong>FieldCode</strong></div></th>
    <th width="29" rowspan="3" style="background-color:#00A300;color:#fff;"><div align="center"><strong>TT</strong></div></th>
    <th width="40" rowspan="3" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Luas</strong></div></th>
    <th colspan="12" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Status Piezo Meter </strong></div></th>
  </tr>
  <tr>
    <th colspan="4" style="background-color:#00A300;color:#fff;"><div align="center"><strong>Oktober</strong></div></th>
    <th colspan="4" style="background-color:#00A300;color:#fff;"><div align="center"><strong>November</strong></div></th>
    <th colspan="4" style="background-color:#00A300;color:#fff;"><div align="center"><strong>December</strong></div></th>
  </tr>
  <tr>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M1</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M2</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M3</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M4</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M1</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M2</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M3</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M4</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M1</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M2</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M3</strong></div></th>
    <th width="24" align="center" valign="middle" style="background-color:#00A300;color:#fff;"><div align="center"><strong>M4</strong></div></th>
	</tr></thead></tbody>
	
<%
if rsPusingan.bof and rsPusingan.eof then
%>
	<tr>
		<td colspan="22">NO DATA AVAILABLE</td>
	</tr>
<% else
NO = 1
		Do while not rsPusingan.eof
			if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
<tr class="<%=cClass%>">
		<td align="center"><%=NO%></td>
		<td align="center"><%=rsPusingan("PLANTATION")%></td>
		<td align="center"><%=rsPusingan("REGION")%></td>
		<td align="center"><%=rsPusingan("GROUPKEBUN")%></td>
		<td align="center"><%=rsPusingan("COMPANY")%></td>
		<td align="center"><%=rsPusingan("FCBA")%></td>
		<td align="center"><%=rsPusingan("AFEDLING")%></td>
		<td align="center"><%=rsPusingan("FIELDCODE")%></td>
		<td align="center"><%=rsPusingan("TT")%></td>
		<td align="right"><%=rsPusingan("HA")%></td>	
		<td bgcolor="<%=rsPusingan("W1102016COLOUR")%>" align="center"><font color="#FFFFFF"><%=rsPusingan("W1102016WATERLEVEL")%></font></td>
		<td bgcolor="<%=rsPusingan("W2102016COLOUR")%>" align="center"><%=rsPusingan("W2102016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W3102016COLOUR")%>" align="center"><%=rsPusingan("W3102016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W4102016COLOUR")%>" align="center"><%=rsPusingan("W4102016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W1112016COLOUR")%>" align="center"><%=rsPusingan("W1112016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W2112016COLOUR")%>" align="center"><%=rsPusingan("W2112016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W3112016COLOUR")%>" align="center"><%=rsPusingan("W3112016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W4112016COLOUR")%>" align="center"><%=rsPusingan("W4112016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W1122016COLOUR")%>" align="center"><%=rsPusingan("W1122016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W2122016COLOUR")%>" align="center"><%=rsPusingan("W2122016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W3122016COLOUR")%>" align="center"><%=rsPusingan("W3122016WATERLEVEL")%></td>
		<td bgcolor="<%=rsPusingan("W4122016COLOUR")%>" align="center"><%=rsPusingan("W4122016WATERLEVEL")%></td>
  </tr>

<%
            NO = NO + 1			
			rsPusingan.MoveNext            		
		loop
end if

rsPusingan.Close
set rsPusingan = Nothing
%>
</tbody>
</table>
<div align="center">
<table width="20%">
<tr><td colspan="3"><div align="left"><strong>Index</strong></div></td></tr>
<tr><td  width="21" bgcolor="FF0000"></td><td width="9">:</td><td width="220">Kekurangan Air</td></tr>
<tr><td  width="21" bgcolor="FFFF00"></td><td>:</td><td>Tergenang</td></tr>
<tr><td  width="21" bgcolor="00FF00"></td><td>:</td><td>Agak Tergenang</td></tr>
<tr><td  width="21" bgcolor="009933"></td><td>:</td><td>Normal</td></tr>
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
