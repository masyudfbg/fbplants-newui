<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%

Function FormatString(byVal StringToFormat, byVal FormatPattern)
    Dim i, sTemp, sResult
    sTemp = CStr(StringToFormat)
    sResult = FormatPattern
    Do Until InStr(sResult, "#") = 0
        i = InStr(sResult, "#")
        sResult = Replace(sResult, "#", Left(sTemp, 1), 1, 1)
        sTemp = Mid(sTemp, 2)
        if Len(sTemp) = 0 Then sResult = Left(sResult, i)
    Loop
    FormatString = sResult
End Function

Dim v_txtYear, v_txtMonth, v_txtGangCode
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim ScriptName
Dim rsTransferGaji, strTransferGaji
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName
Dim NO, cClass, nTotalPekerja, nTotalPasangan, nTotalTanggungan, nSubTotal

v_txtFcba = Request("txtFcba")
v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")
v_txtGangCode=	Request("txtGangCode")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE'" & _
               "       and fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("fcba")
Else
	varDefaultBa = Session("fcba")
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

strcurrentdate = " select currentdate" & _
                 " from periodcontrol" & _                 
                 " where fcba = '" & v_txtFcba & "'"

Set rscurrentdate = Server.CreateObject("ADODB.Recordset")
rscurrentdate.Open strcurrentdate, connstr, adOpenStatic

If Not rscurrentdate.EOF Then
	varDefaultYear	=	Year(rscurrentdate("currentdate"))
	varDefaultMonth	=	Month(rscurrentdate("currentdate"))
Else
	varDefaultYear	=	Year(Session("v_Employee"))
	varDefaultMonth	=	Month(Session("v_Employee"))
End If
rscurrentdate.Close
set rscurrentdate = Nothing

If (v_txtYear = "" Or Len(Trim(v_txtYear)) = 0) Then v_txtYear = varDefaultYear
If (v_txtMonth = "" Or Len(Trim(v_txtMonth)) = 0) Then v_txtMonth = varDefaultMonth

if len(v_txtmonth) > 1 then	
   v_txtMonth = FormatString(v_txtMonth,"##")
else
   v_txtMonth = FormatString(v_txtMonth,"0#")
end if 
	
strTransferGaji	=	" select c.fccode, " & _
					"	   c.fcname, " & _
					"	   c.bankname, " & _
					"	   c.accountno, " & _
					"	   c.ownershipofaccount, " & _
					"	   c.TotalAmount TotalAmount" & _
					"  From view_transfergaji c " & _
					" where fcba = '" & v_txtFcba & "'  AND YearMonth = '" & v_txtYear  & v_txtMonth & "'"
					
if  LEN(TRIM(v_txtGangCode)) > 0 then
	strTransferGaji	=	strTransferGaji & " AND c.GANGCODE = '" & v_txtGangCode & "'"
end if

	strTransferGaji	=	strTransferGaji & " ORDER BY c.fccode, c.SECTIONNAME, c.PRESENTGRADE "
'response.write strTransferGaji

strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			"     , FCMANAGER" & _
			"     , FCKTU" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varManagerName = rsBaInfo("FCMANAGER")
	varKTUName = rsBaInfo("FCKTU")
Else
	varFcbaName = ""
	varCompanyName = ""
	varManagerName = ""
	varKTUName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

Set rsTransferGaji = Server.CreateObject("ADODB.Recordset")
rsTransferGaji.Open strTransferGaji, connstr, adOpenStatic 
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - List Payroll Transfer</title>
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
			<!--#include file="../print_excel.asp"-->
			<!--#include file="menutop.asp"-->
		</div>
		<form action="<%=ScriptName%>" method="post" name="frmEntry">
		<div class="printhide">
		<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
		<table id="form-table" class="centered-margin">
			<tr>
				<td>
					<font face="verdana" size="1">TAHUN</font>
				</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 2
							If len(v_txtYear)= 0 then v_txtYear = Yr
						%>
				<td width="4">
					<select name="txtYear">					
						<%For i = MinYr to MaxYr%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
						<%Next%>
					</select>
				</td>
				<td>
					<font face="verdana" size="1">BULAN</font>
				</td>
				<td width="2">
					<select name="txtMonth">
						<%For i = 1 to 12%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
						<%Next%>
					</select>
				</td>		
				<td><font face="verdana" size="1">GANGCODE</font></td>
				<td><select name="txtGangCode" ><option value=""></option>
					<%
					Dim rsGangCode, strGangCode
					strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE ACTIVATION = 'Y' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE ASC"
					
					Set rsGangCode = Server.CreateObject("ADODB.Recordset")
					rsGangCode.Open strGangCode, connstr, adOpenStatic 
					
					if not rsGangCode.eof then
					Do While Not rsGangCode.eof
					if rsGangCode("FCCODE") = v_txtGangCode then%>
							<option value="<%=rsGangCode("FCCODE")%>" Selected><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option> 
						 <%Else%>
							<option value="<%=rsGangCode("FCCODE")%>"><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option>  
						<%End if
						rsGangCode.movenext
					loop
					end if
					rsGangCode.Close
					set rsGangCode = Nothing
					%>
					</select>
		    </td>
			
				<td><input type="submit" value="View" /></td>
			</tr>
		</table>
		</div>
	</form>

</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="11" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
		<h4>PT. <%=varCompanyName%></h4>
		<h5><%=varFcbaName%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>DAFTAR GAJI KARYAWAN TRANSFER</h2>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
 </th>
  </tr>
	<tr>
		<th rowspan="2">NO</th>
		<th rowspan="2">NO PEKERJA</th>
		<th rowspan="2">NAMA PEKERJA</th>		
		<th rowspan="2">NAMA BANK</th>
		<th rowspan="2">NO REKENING</th>
		<th rowspan="2">NAMA PEMEGANG REKENING</th>
		<th rowspan="2">JUMLAH TRANSFER</th>
	</tr>
</thead>
<tbody>
<%if rsTransferGaji.BOF and rsTransferGaji.EOF then%>
	<tr>
		<td colspan="9" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		nTotalPekerja = 0
		nTotalPasangan = 0
		nTotalTanggungan = 0
		nSubTotal = 0
		while not rsTransferGaji.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center" height="30" width="3%"><%=NO%></td>
				<td align="center" width="4%"><%=rsTransferGaji("FCCODE")%></td>
				<td align="left"><%=rsTransferGaji("FCNAME")%></td>
				<td align="left"><%=rsTransferGaji("BANKNAME")%></td>
				<td align="center"><%=rsTransferGaji("ACCOUNTNO")%></td>
				<td align="left"><%=rsTransferGaji("OWNERSHIPOFACCOUNT")%></td>
				<td align="right"><%=formatnumber(rsTransferGaji("TotalAmount"),0)%></td>
			</tr>
<%
			NO = NO + 1

			nSubTotal = nSubTotal + Cdbl(rsTransferGaji("TotalAmount"))
			rsTransferGaji.MoveNext
		Wend
	rsTransferGaji.Close
	set rsTransferGaji = Nothing
%>
			<tr bgcolor="#F4F4A0">
				
				      <td align="center" colspan="6"><strong>TOTAL</strong></td>
				      <td align="right"><strong><%=formatnumber(nSubTotal,0)%></strong></td>
					 
		  
				
			</tr>
            </tbody>
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=varManagerName%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=varKTUName%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
