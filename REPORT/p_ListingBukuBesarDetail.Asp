<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i
Dim ScriptName
Dim v_txtFcba, rsDefaultBa, strDefaultBa, varDefaultBa, no
Dim rsBaInfo, strBaInfo, varCompanyName, varFcbaName
Dim v_txtSourceType, v_txtSourceCode, v_txtTargetType, v_txtTargetCode, v_txtReferenceNo
Dim v_txtJobCode,v_txtJobCodeto 
Dim v_txtYear, v_txtMonth , v_txtMonthTO
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
v_txtMonthTO		= Request("txtMonthTO")
v_txtSourceType		= Request("txtSourceType")
v_txtSourceCode		= Request("txtSourceCode")
v_txtTargetType		= Request("txtTargetType")
v_txtTargetCode		= Request("txtTargetCode")
v_txtReferenceNo	= Request("txtReferenceNo")
v_txtFcba			= Request("txtFcba")
v_txtjobcode        = Request("txtJobCode")
v_txtjobcodeto      = Request("txtJobCodeto")

If len(trim(v_txtYear)) = 0 then v_txtYear = Year(Session("v_Costbook"))
If len(trim(v_txtMonth)) = 0 then v_txtMonth = Month(Session("v_Costbook"))
If len(trim(v_txtMonthTO)) = 0 then v_txtMonthTO = Month(Session("v_Costbook"))

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

If (v_txtJobCode = "" Or Len(Trim(v_txtJobCode)) = 0) Then
	v_txtJobCode = ""
End If

If (v_txtJobCodeto = "" Or Len(Trim(v_txtJobCodeto)) = 0) Then
	v_txtJobCodeto = ""
End If


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
			
strJob		= 	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & v_txtFcba & "' AND FCCODE = '" & v_txtJobCode & "'"
strJob		=	strJob & " ORDER BY FCCODE ASC"

Set rsJob = Server.CreateObject("ADODB.Recordset")
rsJob.Open strJob, connstr, adOpenStatic 
			
%>
<html>

<head>
<title>ePlaS Reporting - Buku Besar Detail</title>
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
					<td><font face="verdana" size="1">PERIOD</font></td>
					<td>
					<table>
						<tr>
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
							<font face="verdana" size="1">s/d</font>
							</td>
							
							<td>
								<select name="txtMonthTO">
								<%
									For i = 1 to 15
								%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonthTO) then%> Selected <%End If%>><%=BulanAP(i)%></option>
								<%
									Next
								%>
								</select>
							</td>
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
					
					</Table>
					<tr>
					<td><font face="verdana" size="1"><%=titleSite%></font></td>
					<td>
						<select name="txtFcba" >
							<%
							Dim rsFcba, strFcba
							strFcba	=	" SELECT 'ALL' FCCODE,' ' FCCOMPANYCODE FROM DUAL  UNION SELECT FCCODE, FCCOMPANYCODE FROM BUSINESSUNIT" & _
										" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
										" ORDER BY FCCOMPANYCODE"
							
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
				  </tr>
				  </tr>
                  <tr>					
					<td>
					<font face="verdana" size="1">JOBCODE</font></td>
					<td>
					<select name="txtJobCode" class="chosen-select"><option value=""></option>
					<%
					Dim rsJobCode, strJobCode
					strJobCode	=	"SELECT ' ' FCCODE,'--ALL--' FCNAME FROM DUAL UNION SELECT DISTINCT FCCODE, FCNAME FROM JOB WHERE FCBA IN (SELECT FCCODE FROM BUSINESSUNIT WHERE FCCOMPANYCODE='" & Session("COMPANYCODE") & "') ORDER BY FCCODE ASC"
					
					'If Len(trim(v_txtFcba)) = 0 Then
					'	strJobCode	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
					'ELse
				   '	strJobCode	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & v_txtFcba & "' ORDER BY FCCODE ASC"
					'End IF
					
					
					Set rsJobCode = Server.CreateObject("ADODB.Recordset")
					rsJobCode.Open strJobCode, connstr, adOpenStatic 
					
					if not rsJobCode.eof then
					Do While Not rsJobCode.eof
					%>
							<option value="<%=rsJobCode("FCCODE")%>" <%if v_txtjobcode = rsJobCode("FCCODE") then%> Selected <%End If%>><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>
					<%
						rsJobCode.movenext
					loop
					end if
					rsJobCode.Close
					set rsJobCode = Nothing
					%>
					</select>
					</td>
					<td ><font face="verdana" size="1">TO</font></td>
					<td>
					<select name="txtJobCodeto" class="chosen-select"><option value=""></option>
					<%
					
					strJobCode	=	"SELECT ' ' FCCODE,'--ALL--' FCNAME FROM DUAL UNION SELECT DISTINCT FCCODE, FCNAME FROM JOB WHERE FCBA IN (SELECT FCCODE FROM BUSINESSUNIT WHERE FCCOMPANYCODE='" & Session("COMPANYCODE") & "') ORDER BY FCCODE ASC"
					
					'If Len(trim(v_txtFcba)) = 0 Then
					'	strJobCode	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
					'ELse
				   '	strJobCode	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & v_txtFcba & "' ORDER BY FCCODE ASC"
					'End IF
					
					Set rsJobCode = Server.CreateObject("ADODB.Recordset")
					rsJobCode.Open strJobCode, connstr, adOpenStatic 
					
					if not rsJobCode.eof then
					Do While Not rsJobCode.eof
					%>
							<option value="<%=rsJobCode("FCCODE")%>" <%if v_txtjobcodeto = rsJobCode("FCCODE") then%> Selected <%End If%>><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>
					<%
					rsJobCode.movenext
					loop
					end if
					rsJobCode.Close
					set rsJobCode = Nothing
					%>
					</select>
					
				</tr>
				<tr>
					<td>
					<font face="verdana" size="1">SOURCE TYPE</font></td>
					<td>
					<select name="txtSourceType" ><option value=""></option>
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
					</select>
					</td>
					<td>
					<font face="verdana" size="1">TARGET TYPE</font></td>
					<td>
					<select name="txtTargetType" ><option value=""></option>
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
					</select>
					</td>
					<tr>
					<td><font face="verdana" size="1">NO.JOURNAL</font></td>
					<td><input type="text" name="txtReferenceNo" width="5" value="<%=v_txtReferenceNo%>"></td>
					</tr>
					<tr>
					<td><input type="submit" value="View" /></td>
					</tr>
			</table>
		</div>
		</form>	
	</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
	<tr>
         <th colspan="15" style="background-color:#FFFFFF; border:none;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>BUKU BESAR DETAIL</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=BulanAP(v_txtMonth)%>&nbsp;s/d&nbsp;<%=BulanAP(v_txtMonthTO)%> <%=v_txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	

	<tr>
	    <th style="color:#FFFFFF" rowspan="2">Kode.Unit</th>
		<th style="color:#FFFFFF" rowspan="2">Kode.Akun</th>
		<th style="color:#FFFFFF" rowspan="2">Nama.Akun</th>
		<th style="color:#FFFFFF" rowspan="2">No.Journal</th>
		<th style="color:#FFFFFF" rowspan="2">No.Reference</th>
		<th style="color:#FFFFFF" rowspan="2">Tgl</th>
		<th style="color:#FFFFFF" rowspan="2">Kode.Source</th>
		<th style="color:#FFFFFF" rowspan="2">Nama.Source</th>
		<th style="color:#FFFFFF" rowspan="2">Keterangan</th>
		<th style="color:#FFFFFF" rowspan="2">Nama.Target</th>
		<th style="color:#FFFFFF" rowspan="2">Kode.Target</th>
		<th style="color:#FFFFFF" colspan="4">Jumlah (Rp)</th>
	</tr>
	<tr>
	    <th style="color:#FFFFFF">Saldo Awal</th>
		<th style="color:#FFFFFF">Debet</th>
		<th style="color:#FFFFFF">Credit</th>
		<th style="color:#FFFFFF">Saldo Akhir</th>
	</tr>
	</thead>
	<tbody>
<%
Dim rsAnalisaBukuBesar, strAnalisaBukuBesar, rowno,sumdebit,sumkredit,sumawal,sumakhir
Dim array_data, array_eof, xx,xx_sum
rowno		=	0
sumdebit  = 0
sumkredit = 0 
sumawal = 0 
sumakhir = 0
if v_txtFcba <> "ALL" then
strAnalisaBukuBesar		=	" SELECT fcba,header, jobcode, jobname,journalno,referenceno, tdate, sourcetype, sourcetypedesc,  remarks, targettypedesc, targetcode, opening_balance,debet, credit,ending_balance,period_month,period_year" & _
							" FROM  " & _
							" ( " & _
							" SELECT '1' Header,fcba,'' journalno,'' referenceno, sysdate tdate, fccode jobcode, fcname jobname,'' sourcetype, '' sourcetypedesc,  '' remarks, '' targettypedesc, '' targetcode, opening_balance,debet,credit,ending_balance,0 lineno,to_number(period_month) period_month,to_number(period_year) period_year "& _ 
							" FROM trialbalance a WHERE a.fcba='" & v_txtFcba & "' AND a.period_month=" & v_txtMonth & " AND a.period_year= " & v_txtYear & "" & _ 
							" UNION " & _
							" SELECT '2' Header,fcba, journalno,referenceno, tdate, jobcode, jobname,sourcetype, sourcetypedesc,  remarks, targettypedesc, targetcode, 0 opening_balance,CASE WHEN amount > 0 THEN amount ELSE 0 END debet,CASE WHEN amount< 0 THEN amount ELSE 0 END credit,0 ending_balance,lineno,to_number(to_char(tdate,'MM')) period_month,to_number(to_char(tdate,'YYYY')) period_year " & _ 
							" FROM  view_bukudua a  WHERE a.fcba='" & v_txtFcba & "' AND (a.accountingperiod between  " & v_txtMonth & " AND " & v_txtMonthTO & ") AND a.accountingyear=" & v_txtYear & "" & _ 
							" )x " & _
							" WHERE 1=1 "
							

if LEN(TRIM(v_txtjobcode)) > 0 AND LEN(TRIM(v_txtjobcodeto)) > 0 Then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and (jobcode between '" & trim(v_txtjobCode) & "' and '" & trim(v_txtjobCodeto) & "') " 
	else
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and jobcode like '" & trim(v_txtjobCode) & "%'"
end if

ELSE
strAnalisaBukuBesar		=	" SELECT fcba,header, jobcode, jobname,journalno,referenceno, tdate, sourcetype, sourcetypedesc,  remarks, targettypedesc, targetcode, opening_balance,debet, credit,ending_balance,period_month,period_year" & _
							" FROM  " & _
							" ( " & _
							" SELECT '1' Header,fcba,'' journalno,'' referenceno, sysdate tdate, fccode jobcode, fcname jobname,'' sourcetype, '' sourcetypedesc,  '' remarks, '' targettypedesc, '' targetcode, opening_balance,debet,credit,ending_balance,0 lineno,to_number(period_month) period_month,to_number(period_year) period_year "& _ 
							" FROM trialbalance a WHERE a.fcba Like '%' AND a.period_month=" & v_txtMonth & " AND a.period_year= " & v_txtYear & "" & _ 
							" UNION " & _
							" SELECT '2' Header,fcba, journalno,referenceno, tdate, jobcode, jobname,sourcetype, sourcetypedesc,  remarks, targettypedesc, targetcode, 0 opening_balance,CASE WHEN amount > 0 THEN amount ELSE 0 END debet,CASE WHEN amount< 0 THEN amount ELSE 0 END credit,0 ending_balance,lineno,to_number(to_char(tdate,'MM')) period_month,to_number(to_char(tdate,'YYYY')) period_year " & _ 
							" FROM  view_bukudua a  WHERE a.fcba Like '%' AND (a.accountingperiod between  " & v_txtMonth & " AND " & v_txtMonthTO & ") AND a.accountingyear=" & v_txtYear & "" & _ 
							" )x " & _
							" WHERE 1=1 "

strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and (jobcode between '" & trim(v_txtjobCode) & "' and '" & trim(v_txtjobCodeto) & "') " 

End IF
    

if  len(trim(v_txtSourceType)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and sourcetype = '" & v_txtSourceType & "'"
end if

if  len(trim(v_txtTargetType)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and targettype = '" & v_txtTargetType & "'"
end if

if  len(trim(v_txtReferenceNo)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " and journalno like '%" & v_txtReferenceNo & "%'"
end if
	
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " ORDER BY fcba,Header,tdate,journalno,lineno"

	
'Response.Write strAnalisaBukuBesar

Set rsAnalisaBukuBesar = Server.CreateObject("ADODB.Recordset")
rsAnalisaBukuBesar.Open strAnalisaBukuBesar, connstr, adOpenStatic
array_eof = true
if not rsAnalisaBukuBesar.EOF then
	array_data = rsAnalisaBukuBesar.getrows()
	array_eof = false
end if
rsAnalisaBukuBesar.Close
set rsAnalisaBukuBesar = Nothing

if not array_eof then
	For xx = LBound(array_data, 2) To UBound(array_data, 2)
			
			
%>

<% if array_data(1, xx) = "1" then %>
       <tr bgcolor="#F4F4A0">
	   <td align="left"><%=array_data(0, xx)%></strong></td>
	   <td align="left"><strong><%=array_data(2, xx)%></strong></td>
	   <td align="left"><strong><%=array_data(3, xx)%></strong></td>
	   <td align="left" colspan="8" style="padding-left:10px;"></td>
	   <td align="right"><strong><%=FormatNumber(array_data(12, xx),0)%></strong></td>
	   <td align="right"><strong><%=FormatNumber(array_data(13, xx),0)%></strong></td>
	   <td align="right"><strong><%=FormatNumber(array_data(14, xx),0)%></strong></td>
	   <td align="right"><strong><%=FormatNumber(array_data(15, xx),0)%></strong></td>
	   </tr>

<% 
			sumawal = sumawal + Cdbl(array_data(12, xx))
	else 
			sumdebit =  sumdebit + Cdbl(array_data(13, xx))
			sumkredit = sumkredit + Cdbl(array_data(14, xx))
%>     
		<tr>
		    <td align="left"><%=array_data(0, xx)%></td>
			<td align="left"></td>
			<td align="left"></td>
			<td align="left"><a href="p_listingBukuBesar.asp?txtFcba=<%=array_data(0, xx)%>&txtYear=<%=array_data(17, xx)%>&txtMonth=<%=array_data(16, xx)%>&txtReferenceNo=<%=array_data(4, xx)%>"  target="_blank"><%=array_data(4, xx)%></a></td>
			<td align="left"><%=array_data(5, xx)%></td>
			<td align="left"><%=array_data(6, xx)%></td>
			<td align="left"><%=array_data(7, xx)%></td>
			<td align="left"><%=array_data(8, xx)%></td>
			<td align="left"><%=array_data(9, xx)%></td>
			<td align="left"><%=array_data(10, xx)%></td>
			<td align="left"><%=array_data(11, xx)%></td>
			<td align="right"><%=FormatNumber(0, 0)%></td>
			<td align="right"><%=FormatNumber(array_data(13, xx), 0)%></td>
			<td align="right"><%=FormatNumber(array_data(14, xx), 0)%></td>
			<td align="right"><%=FormatNumber(0, 0)%></td>
		</tr>
<% end if %>

<%
	next
else %>
					<tr>
						<td colspan="14">NO DATA AVAILABLE</td>
					</tr>   
<%
end if
sumakhir = sumawal+sumdebit+sumkredit
%>

<tr bgcolor="#9CB770">
<td align="Left" colspan="11" ><strong>Grand Total</strong></td>
<td align="right"><strong><%=FormatNumber(sumawal, 0)%></strong></td>
<td align="right"><strong><%=FormatNumber(sumdebit, 0)%></strong></td>
<td align="right"><strong><%=FormatNumber(sumkredit, 0)%></strong></td>
<td align="right"><strong><%=FormatNumber(sumakhir, 0)%></strong></td>


</tr>
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