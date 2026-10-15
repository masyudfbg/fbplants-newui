<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
 <%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade
Dim ScriptName
Dim rsErptoSap, strErptoSap, rsMandor, strMandor, NOMOR
Dim cClass, nday
Dim total40 , total50 , totalelse
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName, no

v_txtFcba = Request("txtFcba")
v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")
ScriptName = Request.ServerVariables("SCRIPT_NAME")

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

strErptoSap		= 	" SELECT FCBA," & _
					"        DOCUMENT_DATE," & _
					"        DOCUMENT_TYPE," & _
					"        COMPANY_CODE," & _
					"        POSTING_DATE," & _
					"        PERIOD," & _
					"        ACCOUNTINGPERIOD,ACCOUNTINGYEAR," & _
					"        CURRENCY," & _
					"        POSTING_KEY," & _
					"        BUSINESS_AREA," & _
					"        GL_ACCOUNT," & _
					"        COST_CENTER," & _
					"        INTERNAL_ORDER," & _
					"        JOBCODE," & _
					"        SUM(DOCUMENT_CCY) DOCUMENT_CCY," & _
					"        owner" & _
					" FROM ( SELECT FCBA," & _
					"               DOCUMENT_DATE," & _
					"               DOCUMENT_TYPE," & _
					"               COMPANY_CODE," & _
					"               POSTING_DATE," & _
					"               PERIOD," & _
					"               CURRENCY," & _
					"               POSTING_KEY," & _
					"               BUSINESS_AREA," & _
					"               GL_ACCOUNT," & _
					"               COST_CENTER," & _
					"               INTERNAL_ORDER," & _
					"               JOBCODE," & _
					"               ACCOUNTINGPERIOD,ACCOUNTINGYEAR ," & _
					"               SUM (DOCUMENT_CCY) DOCUMENT_CCY," & _
					"               decode( targettype,'FF',erp_report_api.Get_owenership(targetcode, fcba, tdate)) owner" & _
					" FROM ERP_REPORT_SAP_GENERATED" & _
					" GROUP BY FCBA," & _
					"          DOCUMENT_DATE," & _
					"          DOCUMENT_TYPE," & _
					"          COMPANY_CODE," & _
					"          POSTING_DATE," & _
					"          PERIOD," & _
					"          CURRENCY," & _
					"          POSTING_KEY," & _
					"          BUSINESS_AREA," & _
					"          GL_ACCOUNT," & _
					"          COST_CENTER," & _
					"          INTERNAL_ORDER," & _
					"          JOBCODE,ACCOUNTINGPERIOD,ACCOUNTINGYEAR ," & _
					"          decode( targettype,'FF',erp_report_api.Get_owenership(targetcode, fcba, tdate)))" & _
					" WHERE FCBA = '" & v_txtFcba & "'"
										
if  LEN(TRIM(v_txtYear)) > 0 then
	strErptoSap	=	strErptoSap & " AND ACCOUNTINGYEAR = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strErptoSap	=	strErptoSap & " AND ACCOUNTINGPERIOD = " & v_txtMonth
end if

strErptoSap	=	strErptoSap & " GROUP BY FCBA," & _
							" DOCUMENT_DATE," & _
							" DOCUMENT_TYPE," & _
							" COMPANY_CODE," & _
							" POSTING_DATE," & _
							" PERIOD," & _
							" ACCOUNTINGPERIOD,ACCOUNTINGYEAR," & _
							" CURRENCY," & _
							" POSTING_KEY," & _
							" BUSINESS_AREA," & _
							" GL_ACCOUNT," & _
							" COST_CENTER," & _
							" INTERNAL_ORDER," & _
							" JOBCODE," & _
							" OWNER "

	


'response.write strErptoSap
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Fico </title>
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
							<td><font face="verdana" size="1">YEAR</font></td>
							<% 
								Dim mon, Yr, MinYr, MaxYr, i, MM
								mon = varDefaultMonth
								If len(v_txtMonth)=0 then v_txtMonth = mon

								Yr = varDefaultYear
								MinYr = 2015
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
							<td><font face="verdana" size="1">MONTH</font></td>
							<td width="2">
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
							<td><input type="submit" value="View" /></td>
						</tr>
					</table>
				</div>			
			</form>
		</div>

<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="38" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
			<h4>PT. <%=varCompanyName%></h4>
			<h5><%=varFcbaName%></h5>
		</div>
		<div id="subheader" class="centered">
			<h2>INTEGRASI IPLAS - SAP (FICO)</h2>
			<br/>
			<h4>PERIODE&nbsp;:&nbsp;<%=BulanAP(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
		</div>
	</div>
   </th>
  </tr>
  
	<tr>
		<th>Document Date</th>
		<th>Document Type</th>
		<th>Company Code</th>
		<th>Posting Date</th>
		<th>Periode</th>
		<th>Currency</th>
		<th>Reference</th>
		<th>Header Text</th>
		<th>Posting Key</th>
		<th>Bus Area</th>
		<th>Account Code</th>
		<th>Special GL</th>
		<th>Tran Type</th>
		<th>Document Ccy</th>
		<th>Local Ccy</th>
		<th>Group ccy</th>
		<th>Purchasing Doc</th>
		<th>PO Item</th>
		<th>Sales Order</th>
		<th>SO Item</th>
		<th>Quantity</th>
		<th>Unit of Measure</th>
		<th>Tax Code</th>
		<th>Value Date</th>
		<th>Bline Date</th>
		<th>payment term</th>
		<th>Allocation</th>
		<th>Text</th>
		<th>Cost Center</th>
		<th>Material</th>
		<th>Order</th>
		<th>payment reference</th>
		<th>plant</th>
		<th>reference key 1</th>
		<th>reference key 2</th>
		<th>reference key 3</th>
		<th>profit center</th>
		<th>trading partner</th>
	</tr>
</thead>
<tbody>
<%
Set rsErptoSap = Server.CreateObject("ADODB.Recordset")
rsErptoSap.Open strErptoSap, connstr, adOpenStatic


if rsErptoSap.bof and rsErptoSap.eof then
%>
	<tr><td colspan="38">NO DATA AVAILABLE</td></tr>
<%
else
	NO = 1
     
	 total40 = 0
	 total50 = 0
	 totalelse = 0
	while not rsErptoSap.eof
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass ="default"
		end if
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=rsErptoSap("DOCUMENT_DATE")%></td>
		<td align="center"><%=rsErptoSap("DOCUMENT_TYPE")%></td>
		<td align="center"><%=rsErptoSap("COMPANY_CODE")%></td>
		<td align="center"><%=rsErptoSap("POSTING_DATE")%></td>
		<td align="center"><%=rsErptoSap("ACCOUNTINGPERIOD")%></td>
		<td align="center"><%=rsErptoSap("CURRENCY")%></td>
		<td align="center"><%=rsErptoSap("BUSINESS_AREA")%></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("POSTING_KEY")%></td>
		<td align="center"><%=rsErptoSap("BUSINESS_AREA")%></td>
		<td align="center"><%=rsErptoSap("GL_ACCOUNT")%></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="right"><%=FormatNumber(rsErptoSap("DOCUMENT_CCY"),2)%></td>
		<%
		 if rsErptoSap("POSTING_KEY")="40" then
		    total40 = cdbl(total40) +   cdbl(rsErptoSap("DOCUMENT_CCY")) 
		 elseif rsErptoSap("POSTING_KEY")="50" then
            total50 = cdbl(total50) +   cdbl(rsErptoSap("DOCUMENT_CCY")) 
		 else
		    totalelse = cdbl(totalelse) +   cdbl(rsErptoSap("DOCUMENT_CCY")) 
		 end if
		%>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("OWNER")%></td>
		<td align="center"><%=rsErptoSap("JOBCODE")%></td>
		<td align="center"><%=rsErptoSap("COST_CENTER")%></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("INTERNAL_ORDER")%></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"></td>
		

<%
	NO = NO + 1

	rsErptoSap.MoveNext
Wend
end if
rsErptoSap.Close
set rsErptoSap = Nothing
%>
	</tr>
    </tbody>
</table>
  
<table width="90%" class="data">
    <tr>
		<td align="right">Total Debet  (40) = <%=FormatNumber(total40,2)%></td>
		<td>&nbsp;</td>
	</tr>
	 <tr> 
		<td align="right">Total Kredit (50) = <%=FormatNumber(total50,2)%></td>
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
