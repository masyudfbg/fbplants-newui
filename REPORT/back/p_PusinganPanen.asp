<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtAfdeling, v_txtBlok
Dim ScriptName
Dim rsPusingan, strPusingan, strsql
Dim NO, cClass, nday


v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtAfdeling	=	Request("txtAfdeling")
v_txtBlok		=	Request("txtBlok")

If v_txtYear = "" Then
	v_txtYear = Year(Session("v_Costbook"))
End If

If v_txtMonth = "" Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strPusingan	=	" SELECT" & _
				"      HI.FCBA" & _
				"      , TO_NUMBER (TO_CHAR (HI.ONDATE, 'YYYY')) TAHUN" & _
				"      , TO_NUMBER (TO_CHAR (HI.ONDATE, 'MM')) BULAN" & _
				"      , FF.DIVISION" & _
				"      , FF.HARVESTINGSECTION" & _
				"      , HI.FIELDCODE" & _
				"      , TO_NUMBER (TO_CHAR (FF.PLANTINGDATE, 'YYYY')) TAHUNTANAM" & _
				"      , HECTARAGEPLANTED" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 1, ORIGINALNUMBER, 0)) DAY_1" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 1, COLORIDENTIFICATION, 0)) COLOR_1" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 2, ORIGINALNUMBER, 0)) DAY_2" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 2, COLORIDENTIFICATION, 0)) COLOR_2" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 3, ORIGINALNUMBER, 0)) DAY_3" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 3, COLORIDENTIFICATION, 0)) COLOR_3" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 4, ORIGINALNUMBER, 0)) DAY_4" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 4, COLORIDENTIFICATION, 0)) COLOR_4" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 5, ORIGINALNUMBER, 0)) DAY_5" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 5, COLORIDENTIFICATION, 0)) COLOR_5" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 6, ORIGINALNUMBER, 0)) DAY_6" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 6, COLORIDENTIFICATION, 0)) COLOR_6" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 7, ORIGINALNUMBER, 0)) DAY_7" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 7, COLORIDENTIFICATION, 0)) COLOR_7" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 8, ORIGINALNUMBER, 0)) DAY_8" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 8, COLORIDENTIFICATION, 0)) COLOR_8" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 9, ORIGINALNUMBER, 0)) DAY_9" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 9, COLORIDENTIFICATION, 0)) COLOR_9" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 10, ORIGINALNUMBER, 0)) DAY_10" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 10, COLORIDENTIFICATION, 0)) COLOR_10" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 11, ORIGINALNUMBER, 0)) DAY_11" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 11, COLORIDENTIFICATION, 0)) COLOR_11" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 12, ORIGINALNUMBER, 0)) DAY_12" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 12, COLORIDENTIFICATION, 0)) COLOR_12" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 13, ORIGINALNUMBER, 0)) DAY_13" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 13, COLORIDENTIFICATION, 0)) COLOR_13" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 14, ORIGINALNUMBER, 0)) DAY_14" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 14, COLORIDENTIFICATION, 0)) COLOR_14" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 15, ORIGINALNUMBER, 0)) DAY_15" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 15, COLORIDENTIFICATION, 0)) COLOR_15" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 16, ORIGINALNUMBER, 0)) DAY_16" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 16, COLORIDENTIFICATION, 0)) COLOR_16" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 17, ORIGINALNUMBER, 0)) DAY_17" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 17, COLORIDENTIFICATION, 0)) COLOR_17" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 18, ORIGINALNUMBER, 0)) DAY_18" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 18, COLORIDENTIFICATION, 0)) COLOR_18" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 19, ORIGINALNUMBER, 0)) DAY_19" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 19, COLORIDENTIFICATION, 0)) COLOR_19" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 20, ORIGINALNUMBER, 0)) DAY_20" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 20, COLORIDENTIFICATION, 0)) COLOR_20" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 21, ORIGINALNUMBER, 0)) DAY_21" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 21, COLORIDENTIFICATION, 0)) COLOR_21" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 22, ORIGINALNUMBER, 0)) DAY_22" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 22, COLORIDENTIFICATION, 0)) COLOR_22" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 23, ORIGINALNUMBER, 0)) DAY_23" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 23, COLORIDENTIFICATION, 0)) COLOR_23" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 24, ORIGINALNUMBER, 0)) DAY_24" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 24, COLORIDENTIFICATION, 0)) COLOR_24" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 25, ORIGINALNUMBER, 0)) DAY_25" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 25, COLORIDENTIFICATION, 0)) COLOR_25" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 26, ORIGINALNUMBER, 0)) DAY_26" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 26, COLORIDENTIFICATION, 0)) COLOR_26" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 27, ORIGINALNUMBER, 0)) DAY_27" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 27, COLORIDENTIFICATION, 0)) COLOR_27" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 28, ORIGINALNUMBER, 0)) DAY_28" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 28, COLORIDENTIFICATION, 0)) COLOR_28" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 29, ORIGINALNUMBER, 0)) DAY_29" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 29, COLORIDENTIFICATION, 0)) COLOR_29" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 30, ORIGINALNUMBER, 0)) DAY_30" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 30, COLORIDENTIFICATION, 0)) COLOR_30" & _
				"      , SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 31, ORIGINALNUMBER, 0)) DAY_31" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 31, COLORIDENTIFICATION, 0)) COLOR_31" & _
				" FROM " & _
				"      HARVESTING_INTERVAL_TAB HI" & _
				" INNER JOIN FIELD FF ON (HI.FCBA = FF.FCBA AND HI.FIELDCODE = FF.FCCODE)" & _
				" WHERE " & _
				"      FF.ACTIVATION = 'Y' AND FF.STATUS = 'TM' AND FF.HARVESTINGSECTION IS NOT NULL" & _
				"      AND HI.FCBA = '" & Session("FCBA") & "'"
				
If Len(Trim(v_txtYear)) > 0 Then
	strPusingan	=	strPusingan	&	"      AND TO_NUMBER (TO_CHAR (HI.ONDATE, 'YYYY')) = " & v_txtYear
End If
If Len(Trim(v_txtMonth)) > 0 Then
	strPusingan	=	strPusingan	&	"      AND TO_NUMBER (TO_CHAR (HI.ONDATE, 'MM')) = " & v_txtMonth
End If		
If v_txtAfdeling <> "ALL" Then
	strPusingan	=	strPusingan	&	"      AND FF.DIVISION = '" & v_txtAfdeling & "'"
Else
	strPusingan	=	strPusingan	&	" "
End If
If v_txtBlok <> "ALL" Then
	strPusingan	=	strPusingan	&	"      AND HI.FIELDCODE = '" & v_txtBlok & "'"
Else
	strPusingan	=	strPusingan	&	" "
End If

	strPusingan	=	strPusingan	&	" GROUP BY " & _
									"      HI.FCBA" & _
									"      , TO_NUMBER (TO_CHAR (HI.ONDATE, 'YYYY'))" & _
									"      , TO_NUMBER (TO_CHAR (HI.ONDATE, 'MM'))" & _
									"      , FF.DIVISION" & _
									"      , FF.HARVESTINGSECTION" & _
									"      , HI.FIELDCODE" & _
									"      , HECTARAGEPLANTED" & _
									"      , TO_NUMBER (TO_CHAR (FF.PLANTINGDATE, 'YYYY'))"
									
									
	strPusingan	=	strPusingan	&	" ORDER BY " & _
									"      HI.FCBA" & _
									"      , TO_NUMBER (TO_CHAR (HI.ONDATE, 'YYYY'))" & _
									"      , TO_NUMBER (TO_CHAR (HI.ONDATE, 'MM'))" & _
									"      , FF.DIVISION" & _
									"      , FF.HARVESTINGSECTION" & _
									"      , HI.FIELDCODE"
									' "      , TO_NUMBER (TO_CHAR (FF.PLANTINGDATE, 'YYYY'))" & _
									
									
									


'response.write strPusingan
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Pusingan</title>
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
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
				<div class="printhide">
					<table id="form-table" class="centered-margin">	
						<tr>
							<td><font face="verdana" size="1">TAHUN</font></td>
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
							</td>
							<td>
								<font face="verdana" size="1">AFDELING</font>
							</td>
							<td>
								<select name="txtAfdeling" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsAfdeling, strAfdeling
									strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE"
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = v_txtAfdeling then%>
									<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE")%></option>  
									<%End if
										rsAfdeling.movenext
									loop
									end if
									rsAfdeling.Close
									set rsAfdeling = Nothing
									%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">BLOK</font>
							</td>
							<td>
								<select name="txtBlok" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsBlok, strBlok
									strBlok	=	" SELECT FCCODE, FCNAME FROM FIELD WHERE FCBA = '" & Session("FCBA") & "'" & _
												" AND DX_BLOCK_STATUS = 'TM' ORDER BY FCCODE"
									
									Set rsBlok = Server.CreateObject("ADODB.Recordset")
									rsBlok.Open strBlok, connstr, adOpenStatic 
									
									if not rsBlok.eof then
									Do While Not rsBlok.eof
									if rsBlok("FCCODE") = v_txtBlok then%>
									<option value="<%=rsBlok("FCCODE")%>" Selected><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option> 
									<%Else%>
									<option value="<%=rsBlok("FCCODE")%>"><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option>  
									<%End if
										rsBlok.movenext
									loop
									end if
									rsBlok.Close
									set rsBlok = Nothing
									%>
								</select>
							</td>
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
			<div id="header">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>LAPORAN PUSINGAN PANEN</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
		</div>

<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">SEKSI</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">TT</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">LUAS</th>
		<th colspan="<%=GetLastDay(v_txtMonth, v_txtYear)%>" style="background-color:#00A300;color:#fff;">TANGGAL</th>
	</tr>
	<tr>
		<%for nday = 1 to GetLastDay(v_txtMonth, v_txtYear)%>
		<th style="background-color:#00A300;color:#fff;"><%=nday%></th>
		<%next%>
	</tr>

<%
Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan, connstr, adOpenStatic

Dim xNum

if rsPusingan.eof and rsPusingan.eof then
%>
	<tr>
		<td colspan="<%=5+GetLastDay(v_txtMonth, v_txtYear)%>">NO DATA AVAILABLE</td>
	</tr>
<%end if


		while not rsPusingan.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=rsPusingan("DIVISION")%></td>
				<td align="center"><%=rsPusingan("HARVESTINGSECTION")%></td>
				<td align="center"><%=rsPusingan("FIELDCODE")%></td>
				<td align="center"><%=rsPusingan("TAHUNTANAM")%></td>
				<td align="center"><%=rsPusingan("HECTARAGEPLANTED")%></td>
				<%For xNum = 1 to GetLastDay(v_txtMonth, v_txtYear)%>
				<td align="right" bgcolor="<%=rsPusingan("color_"&xNum&"")%>"><%=rsPusingan("day_"&xNum&"")%></td>
				<%Next%>
			</tr>


<%
			rsPusingan.MoveNext
		Wend
rsPusingan.Close
set rsPusingan = Nothing
%>

</table>

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
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
