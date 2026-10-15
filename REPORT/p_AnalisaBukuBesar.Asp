<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, v_txtJobCode
Dim ScriptName
Dim v_txtFcba, rsDefaultBa, strDefaultBa, varDefaultBa, no
Dim rsBaInfo, strBaInfo, varCompanyName, varFcbaName
Dim v_txtSourceType, v_txtSourceCode, v_txtTargetType, v_txtTargetCode, v_txtReferenceNo
Dim v_txtYear, v_txtMonth
Dim rsJob, strJob
Dim Currsourcetype, Prevsourcetype, Currsourcetypedesc, Prevsourcetypedesc
Dim SubTotalPositive : SubTotalPositive = 0
Dim SubTotalNegative : SubTotalNegative = 0
Dim GrandTotalPositive : GrandTotalPositive = 0
Dim GrandTotalNegative : GrandTotalNegative = 0
Dim NewSourceType
Dim n_SaldoSub : n_SaldoSub = 0
Dim n_SaldoGrand : n_SaldoGrand = 0
Dim n_SaldoRow : n_SaldoRow = 0

ScriptName			= Request.ServerVariables("SCRIPT_NAME")
v_txtYear			= Request("txtYear")
v_txtMonth			= Request("txtMonth")
v_txtJobCode		= Request("txtJobCode")
v_txtSourceType		= Request("txtSourceType")
v_txtSourceCode		= Request("txtSourceCode")
v_txtTargetType		= Request("txtTargetType")
v_txtTargetCode		= Request("txtTargetCode")
v_txtReferenceNo	= Request("txtReferenceNo")
v_txtFcba			= Request("txtFcba")

If len(trim(v_txtYear)) = 0 then v_txtYear = Year(Session("v_Costbook"))
If len(trim(v_txtMonth)) = 0 then v_txtMonth = Month(Session("v_Costbook"))

Function IsPostBack()
     If Request.ServerVariables("Request_Method") = "POST" Then
          IsPostBack = True
     Else
          IsPostBack = False
     End If
	End Function
	
strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
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

'If (v_txtJobCode = "" Or Len(Trim(v_txtJobCode)) = 0) Then
'	v_txtJobCode = "1110103"
'End If

strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
Else
	varFcbaName = ""
	varCompanyName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing



			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			'response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			


			
%>
<html>

<head>
<title>ePlaS Reporting - Analisa Buku Besar</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<link href="../plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="../plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="../plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="../plugins/responsivegridsystem/12cols.css" media="all">
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->

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
		<form name="frmBTL" method="post" action="<%=ScriptName%>">
		<div class="printhide">
			<table id="form-table" class="centered-margin">
		
				<tr>
					<td><font face="verdana" size="1">TAHUN</font></td>
					<td>
						<% 
							mon = Month(Session("v_Costbook"))
							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 2
							
						%>
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
					<td>
						<select name="txtMonth">
						<%
							For i = 1 to 15
						%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=BulanAP(i)%></option>
						<%
							Next
						%>
						</select>
					</td>
					<td><font face="verdana" size="1"><%=titleSite%></font></td>
					<td>
						<select name="txtFcba" >
							<%
							Dim rsFcba, strFcba
							strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
										" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
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
					<td><font face="verdana" size="1">JOBCODE</font></td><td><select name="txtJobCode" class="chosen-select"><option value=""></option>
					<%
					Dim rsJobCode, strJobCode
					If Len(trim(v_txtFcba)) = 0 Then
						strJobCode	=	" SELECT ' ' FCCODE, 'ALL' FCNAME FROM DUAL UNION SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
					ELse
						strJobCode	=	" SELECT ' ' FCCODE, 'ALL' FCNAME FROM DUAL UNION SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & v_txtFcba & "' ORDER BY FCCODE ASC"
					End IF
					
					Set rsJobCode = Server.CreateObject("ADODB.Recordset")
					rsJobCode.Open strJobCode, connstr, adOpenStatic 
					
					if not rsJobCode.eof then
					Do While Not rsJobCode.eof
					if rsJobCode("FCCODE") = v_txtJobCode then%>
							<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option> 
						 <%Else%>
							<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>  
						<%End if
						rsJobCode.movenext
					loop
					end if
					rsJobCode.Close
					set rsJobCode = Nothing
					%>
					</select></td>
				</tr>
				<tr>
					<td><font face="verdana" size="1">SOURCE TYPE</font></td><td><select name="txtSourceType" ><option value=""></option>
					<%
					Dim rsSourceType, strSourceType
					strSourceType	=	" select locationtypecode fccode, locationtypename fcname from locationtype order by locationtypecode"
					
					Set rsSourceType = Server.CreateObject("ADODB.Recordset")
					rsSourceType.Open strSourceType, connstr, adOpenStatic 
					
					if not rsSourceType.eof then
					Do While Not rsSourceType.eof
					if rsSourceType("FCCODE") = v_txtSourceType then%>
							<option value="<%=rsSourceType("FCCODE")%>" Selected><%=rsSourceType("FCCODE")%></option> 
						 <%Else%>
							<option value="<%=rsSourceType("FCCODE")%>"><%=rsSourceType("FCCODE")%></option>  
						<%End if
						rsSourceType.movenext
					loop
					end if
					rsSourceType.Close
					set rsSourceType = Nothing
					%>
					</select></td>
					<td><font face="verdana" size="1">SOURCE CODE</font></td>
					<td><input type="text" name="txtSourceCode" width="5" value="<%=v_txtSourceCode%>"></td>
					<td><font face="verdana" size="1">TARGET TYPE</font></td><td><select name="txtTargetType" ><option value=""></option>
					<%
					Dim rsTargetType, strTargetType
					strTargetType	=	" select locationtypecode fccode, locationtypename fcname from locationtype order by locationtypecode"
					
					Set rsTargetType = Server.CreateObject("ADODB.Recordset")
					rsTargetType.Open strTargetType, connstr, adOpenStatic 
					
					if not rsTargetType.eof then
					Do While Not rsTargetType.eof
					if rsTargetType("FCCODE") = v_txtTargetType then%>
							<option value="<%=rsTargetType("FCCODE")%>" Selected><%=rsTargetType("FCCODE")%></option> 
						 <%Else%>
							<option value="<%=rsTargetType("FCCODE")%>"><%=rsTargetType("FCCODE")%></option>  
						<%End if
						rsTargetType.movenext
					loop
					end if
					rsTargetType.Close
					set rsTargetType = Nothing
					%>
					</select></td>
					<td><font face="verdana" size="1">TARGET CODE</font></td>
					<td><input type="text" name="txtTargetCode" width="5" value="<%=v_txtTargetCode%>"></td>
				</tr>
				<tr>
					<td><font face="verdana" size="1">REFERENCE NO</font></td>
					<td><input type="text" name="txtReferenceNo" width="5" value="<%=v_txtReferenceNo%>"></td>
					<td><input type="submit" value="View" /></td>
				</tr>
			</table>
		</div>
		</form>	
	</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
	<tr>
         <th colspan="16" style="background-color:#FFFFFF; border:none;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>ANALISA BUKU BESAR</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=BulanAP(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	
<%
if IsPostBack then

If (v_txtJobCode = "" Or Len(Trim(v_txtJobCode)) = 0) Then
	strJob		= 	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & v_txtFcba & "'"
else
	strJob		= 	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & v_txtFcba & "' AND FCCODE = '" & v_txtJobCode & "'"
End If

strJob		=	strJob & " ORDER BY FCCODE ASC"


Set rsJob = Server.CreateObject("ADODB.Recordset")
rsJob.Open strJob, connstr, adOpenStatic 

if not rsJob.eof then
	while not rsJob.eof%>
<tr>
<th colspan="14" align="left" style="background-color:#FFFFFF; border:none; color:#000000;"><%=rsJob("FCCODE")%>&nbsp;<%=rsJob("FCNAME")%></th>
</tr>


	<tr>
	    <th style="color:#FFFFFF" rowspan="2">JOB CODE</th>
		<th style="color:#FFFFFF" rowspan="2">JOB CODE DESCRIPTION</th>
		<th style="color:#FFFFFF" rowspan="2">SOURCE CODE</th>
		<th style="color:#FFFFFF" rowspan="2">SOURCE CODE DESCRIPTION</th>
		<th style="color:#FFFFFF" rowspan="2">TARGET TYPE</th>
		<th style="color:#FFFFFF" rowspan="2">TARGET TYPE DESCRIPTION</th>
		<th style="color:#FFFFFF" rowspan="2">TARGET CODE</th>
		<th style="color:#FFFFFF" rowspan="2">TARGET CODE DESCRIPTION</th>
		<th style="color:#FFFFFF" rowspan="2">REFERENCE NO</th>
		<th style="color:#FFFFFF" rowspan="2">STOCK CODE</th>
		<th style="color:#FFFFFF" rowspan="2">TAHUN TANAM</th>
		<th style="color:#FFFFFF" rowspan="2">DIVISI</th>
		<th style="color:#FFFFFF" rowspan="2">KETERANGAN</th>
		<th style="color:#FFFFFF" colspan="3">JUMLAH (Rp)</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF">Debet</th>
		<th style="color:#FFFFFF">Credit</th>
		<th style="color:#FFFFFF">Saldo</th>
	</tr>
	</thead>
	<tbody>
<%
Dim oSQL, oRS
Dim n_Opening : n_Opening = 0



oSQL = " Select opening_balance opening" & _
	   " from trialbalance" & _
	   " where fcba = '" & v_txtFcba & "'"
	   
'If Len(Trim(v_txtJobCode)) > 0 then
'	oSQL = oSQL + "       and fccode = '" & v_txtjobCode & "'"
'End If
	oSQL = oSQL + "       and fccode = '" & rsJob("FCCODE") & "'" 
If Len(Trim(v_txtYear)) > 0 then
	oSQL = oSQL + "       and period_year = " & v_txtYear & ""
End If
If Len(Trim(v_txtMonth)) > 0 then
	oSQL = oSQL + "       and period_month = " & v_txtMonth & ""
End If

	oSQL = oSQL + "  Order by FCCODE "

'response.write oSQL 

Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open oSQL, connstr, adOpenStatic


If Not oRS.EOF then
'	n_Opening = n_Opening + Cdbl(oRS("opening"))
    n_Opening = Cdbl(oRS("opening")) 
Else
    n_Opening = 0 
'	n_Opening = n_Opening + 0
End If

oRS.Close
set oRS = Nothing

n_SaldoSub = n_SaldoSub + n_Opening
n_SaldoGrand = n_SaldoGrand + n_Opening
n_SaldoRow = n_SaldoRow + n_Opening
%>
		<tr bgcolor="#9CB770">
			<td align="left" colspan="13" style="padding-left:10px;"><strong><%=PCase("Saldo Awal")%></strong></td>
			<%If n_Opening = 0 Then%>
			<td align="Right"><strong><%=FormatNumber(0, 2)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(0, 2)%></strong></td>
			<%Else%>
				<%If n_Opening > 0 Then%>
				<td align="Right"><strong><%=FormatNumber(n_Opening, 2)%></strong></td>
				<td align="Right"><strong><%=FormatNumber(0, 2)%></strong></td>
				<%ElseIf n_Opening < 0 Then%>
				<td align="Right"><strong><%=FormatNumber(0, 2)%></strong></td>
				<td align="Right"><strong><%=FormatNumber(n_Opening, 2)%></strong></td>
				<%End If%>
			<%End If%>
			<td align="Right"><strong><%=FormatNumber(n_SaldoSub, 2)%></strong></td>
		</tr>
<%
Dim rsAnalisaBukuBesar, strAnalisaBukuBesar, rowno
rowno		=	0
strAnalisaBukuBesar		=	" select fcba" & _
							"        , jobcode" & _
							"        , jobname" & _
							"        , orderdata" & _
							"        , sourcetype" & _
							"        , sourcetypedesc" & _
							"        , sourcecode" & _
							"        , sourcecodedesc" & _
							"        , targettype" & _
							"        , targettypedesc" & _
							"        , targetcode" & _
							"        , targetcodedesc" & _
							"        , referenceno" & _
							"        , stockcode" & _
							"        , PlantingYear" & _
							"        , Division" & _
							"        , Sum(amount) amount" & _
							"        , remarks" & _
							" from view_bukudua" & _
							" where fcba = '" & v_txtFcba & "'"

if  LEN(TRIM(v_txtYear)) > 0 then
	'strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and to_number(to_char (tdate, 'yyyy')) = " & v_txtYear
	strAnalisaBukuBesar	= strAnalisaBukuBesar & " and accountingyear = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	'strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and to_number(to_char (tdate, 'mm')) = " & v_txtMonth
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and accountingperiod = " & v_txtMonth
end if
'if  LEN(TRIM(v_txtJobCode)) > 0 then
	'strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and jobcode = '" & v_txtjobCode & "'"
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and jobcode = '" & rsJob("FCCODE") & "'"
	
'end if
if  len(trim(v_txtSourceType)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and sourcetype = '" & v_txtSourceType & "'"
end if
if  len(trim(v_txtSourceCode)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and sourcecode = '" & v_txtSourceCode & "'"
end if
if  len(trim(v_txtTargetType)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and targettype = '" & v_txtTargetType & "'"
end if
if  len(trim(v_txtTargetCode)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and targetcode = '" & v_txtTargetCode & "'"
end if
if  len(trim(v_txtReferenceNo)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and referenceno = '" & v_txtReferenceNo & "'"
end if
	strAnalisaBukuBesar	=	strAnalisaBukuBesar &   " group by fcba" & _
													"          , jobcode" & _
													"          , jobname" & _
													"          , orderdata" & _
													"          , sourcetype" & _
													"          , sourcetypedesc" & _
													"          , sourcecode" & _
													"          , sourcecodedesc" & _
													"          , targettype" & _
													"          , targettypedesc" & _
													"          , targetcode" & _
													"          , targetcodedesc" & _
													"          , referenceno" & _
													"          , stockcode" & _
													"          , PlantingYear" & _
													"          , Division" & _
													"          , remarks"

	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " Order by orderdata"

	
	Set rsAnalisaBukuBesar = Server.CreateObject("ADODB.Recordset")
	rsAnalisaBukuBesar.Open strAnalisaBukuBesar, connstr, adOpenStatic
	if Not rsAnalisaBukuBesar.eof then
		rsAnalisaBukuBesar.MoveFirst
		While Not rsAnalisaBukuBesar.EOF
				rowno	=	rowno	+	1
				Currsourcetype = rsAnalisaBukuBesar("sourcetype")
				Currsourcetypedesc = rsAnalisaBukuBesar("sourcetypedesc")
				If Cdbl(rsAnalisaBukuBesar("amount")) = 0 Then
					GrandTotalPositive = GrandTotalPositive + 0
					GrandTotalNegative = GrandTotalNegative + 0
				Else
					If Cdbl(rsAnalisaBukuBesar("amount")) > 0 Then
						GrandTotalPositive = GrandTotalPositive + Cdbl(rsAnalisaBukuBesar("amount"))
					ElseIf Cdbl(rsAnalisaBukuBesar("amount")) < 0 Then
						GrandTotalNegative = GrandTotalNegative + Cdbl(rsAnalisaBukuBesar("amount"))
					End If
				End If

				If Currsourcetypedesc <> Prevsourcetypedesc Then
%>
<%If Len(Trim(Prevsourcetypedesc)) = 0 Then%>
		<tr>
			<td colspan="14" align="left" style="padding-left:10px;" bgcolor="#F4F4A0"><strong><%=Currsourcetypedesc%></strong></td>
		</tr>
<%
SubTotalPositive = 0
SubTotalNegative = 0
'n_SaldoSub = n_SaldoSub + SubTotalPositive + SubTotalNegative
%>
<%Else%>
		<tr bgcolor="#9CB770">
			<td align="left" colspan="13" style="padding-left:10px;"><strong>Sub Total <%=PCase(Prevsourcetypedesc)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(SubTotalPositive, 2)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(SubTotalNegative, 2)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(n_SaldoSub, 2)%></strong></td>
		</tr>
		<tr>
			<td colspan="14" align="left" style="padding-left:10px;" bgcolor="#F4F4A0"><strong><%=Currsourcetypedesc%></strong></td>
		</tr>
<%
SubTotalPositive = 0
SubTotalNegative = 0
'n_SaldoSub = n_SaldoSub + SubTotalPositive + SubTotalNegative
%>
<%End If%>
		<tr>
		    <td align="left"><%=rsAnalisaBukuBesar("jobcode")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("jobname")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("sourcecode")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("sourcecodedesc")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("targettype")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("targettypedesc")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("targetcode")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("targetcodedesc")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("referenceno")%></td>
			<%if left(rsAnalisaBukuBesar("stockcode"), 14) = "linenovoucher-" then%>
				<td align="left">&nbsp;</td>
			<%else%>
				<td align="left"><%=rsAnalisaBukuBesar("stockcode")%></td>
			<%end if%>
			
			<td align="left"><%=rsAnalisaBukuBesar("PlantingYear")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("Division")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("remarks")%></td>
			<%If Cdbl(rsAnalisaBukuBesar("amount")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<td align="right"><%=FormatNumber(0, 2)%></td>				
			<%Else%>
				<%If Cdbl(rsAnalisaBukuBesar("amount")) > 0 Then%>
					<td align="right"><%=FormatNumber(rsAnalisaBukuBesar("amount"), 2)%></td>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%ElseIf Cdbl(rsAnalisaBukuBesar("amount")) < 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<td align="right"><%=FormatNumber(rsAnalisaBukuBesar("amount"), 2)%></td>
				<%End If%>
			<%End If%>
			<%n_SaldoRow = n_SaldoRow + Cdbl(rsAnalisaBukuBesar("amount"))%>
			<td align="right"><%=FormatNumber(n_SaldoRow, 2)%></td>
		</tr>
<%
			If Cdbl(rsAnalisaBukuBesar("amount")) = 0 Then
				SubTotalPositive = SubTotalPositive + 0
				SubTotalNegative = SubTotalNegative + 0
			Else
				If Cdbl(rsAnalisaBukuBesar("amount")) > 0 Then
					SubTotalPositive = SubTotalPositive + Cdbl(rsAnalisaBukuBesar("amount"))
				ElseIf Cdbl(rsAnalisaBukuBesar("amount")) < 0 Then
					SubTotalNegative = SubTotalNegative + Cdbl(rsAnalisaBukuBesar("amount"))
				End If
			End If
					n_SaldoSub = n_SaldoSub + Cdbl(rsAnalisaBukuBesar("amount"))
%>

	<%Else%>
		<tr>
		    <td align="left"><%=rsAnalisaBukuBesar("jobcode")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("jobname")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("sourcecode")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("sourcecodedesc")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("targettype")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("targettypedesc")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("targetcode")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("targetcodedesc")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("referenceno")%></td>
			
			<%if left(rsAnalisaBukuBesar("stockcode"), 14) = "linenovoucher-"  then%>
				<td align="left">&nbsp;</td>
			<%else%>
				<td align="left"><%=rsAnalisaBukuBesar("stockcode")%></td>
			<%end if%>
			
			<td align="left"><%=rsAnalisaBukuBesar("PlantingYear")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("Division")%></td>
			<td align="left"><%=rsAnalisaBukuBesar("remarks")%></td>
			<%If Cdbl(rsAnalisaBukuBesar("amount")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 2)%></td>
				<td align="right"><%=FormatNumber(0, 2)%></td>
			<%Else%>
				<%If Cdbl(rsAnalisaBukuBesar("amount")) > 0 Then%>
					<td align="right"><%=FormatNumber(rsAnalisaBukuBesar("amount"), 2)%></td>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%ElseIf Cdbl(rsAnalisaBukuBesar("amount")) < 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
					<td align="right"><%=FormatNumber(rsAnalisaBukuBesar("amount"), 2)%></td>
				<%End If%>
			<%End If%>
			<%n_SaldoRow = n_SaldoRow + Cdbl(rsAnalisaBukuBesar("amount"))%>
			<td align="right"><%=FormatNumber(n_SaldoRow, 2)%></td>
		</tr>
<%
			If Cdbl(rsAnalisaBukuBesar("amount")) = 0 Then
				SubTotalPositive = SubTotalPositive + 0
				SubTotalNegative = SubTotalNegative + 0
			Else
				If Cdbl(rsAnalisaBukuBesar("amount")) > 0 Then
					SubTotalPositive = SubTotalPositive + Cdbl(rsAnalisaBukuBesar("amount"))
				ElseIf Cdbl(rsAnalisaBukuBesar("amount")) < 0 Then
					SubTotalNegative = SubTotalNegative + Cdbl(rsAnalisaBukuBesar("amount"))
				End If
			End If
			n_SaldoSub = n_SaldoSub + Cdbl(rsAnalisaBukuBesar("amount"))
%>

<%End If%>
<%
		
		Prevsourcetype = Currsourcetype
		Prevsourcetypedesc = Currsourcetypedesc
		NewSourceType = rsAnalisaBukuBesar("sourcetypedesc")
		rsAnalisaBukuBesar.MoveNext
		Wend
		end if

rsJob.MoveNext
Wend
end if

n_SaldoSub = n_SaldoSub
n_SaldoGrand = n_SaldoGrand + GrandTotalPositive + GrandTotalNegative
%>
		<tr bgcolor="#9CB770">
			<td align="left" colspan="13" style="padding-left:10px;"><strong>Sub Total <%=PCase(Prevsourcetypedesc)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(SubTotalPositive, 2)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(SubTotalNegative, 2)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(n_SaldoSub, 2)%></strong></td>
		</tr>
		<tr bgcolor="#9CB770">
			<td align="left" colspan="13" style="padding-left:10px;"><strong>Grand Total</strong></td>
			<td align="Right"><strong><%=FormatNumber(GrandTotalPositive, 2)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(GrandTotalNegative, 2)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(n_SaldoGrand, 2)%></strong></td>
		</tr>
<%
rsAnalisaBukuBesar.Close
set rsAnalisaBukuBesar = Nothing
rsJob.Close
set rsJob = Nothing


else
%>
					<tr>
						<td colspan="<%=16%>">NO DATA AVAILABLE</td>
					</tr>
<%end if%>			
</tbody>
</table>
</div>
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="/plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">	
	// $(".chosen-select").empty();
	
	var config = {
	  '.chosen-select'           : {},
	  '.chosen-select-deselect'  : {allow_single_deselect:true},
	  '.chosen-select-no-single' : {disable_search_threshold:10},
	  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
	  '.chosen-select-width'     : {width:"45%"}
	}
	for (var selector in config) {
	  $(selector).chosen(config[selector]);
	}
	// $(".chosen-select").trigger("chosen:updated");
</script>
</body>
</html>