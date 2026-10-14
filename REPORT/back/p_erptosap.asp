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
Dim NO, cClass, nday

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
ScriptName		=	Request.ServerVariables("SCRIPT_NAME")

strErptoSap		= 	" SELECT FCBA," & _
					" DOCUMENT_TYPE," & _
					" COMPANY_CODE," & _
					" POSTING_DATE," & _
					" PERIOD," & _
					" CURRENCY," & _
					" POSTING_KEY," & _
					" BUSINESS_AREA," & _
					" GL_ACCOUNT," & _
					" COST_CENTER," & _
					" INTERNAL_ORDER," & _
					" JOBCODE," & _
					" SUM(DOCUMENT_CCY) DOCUMENT_CCY, owner" & _
					" FROM ( SELECT FCBA," & _
					" DOCUMENT_TYPE," & _
					" COMPANY_CODE," & _
					" POSTING_DATE," & _
					" PERIOD," & _
					" CURRENCY," & _
					" POSTING_KEY," & _
					" BUSINESS_AREA," & _
					" GL_ACCOUNT," & _
					" COST_CENTER," & _
					" INTERNAL_ORDER," & _
					" JOBCODE," & _
					" SUM (DOCUMENT_CCY) DOCUMENT_CCY, decode( targettype,'FF',erp_report_api.Get_owenership(targetcode,fcba)) owner" & _
					" FROM ERP_REPORT_SAP_GENERATED" & _
					" GROUP BY FCBA," & _
					" DOCUMENT_TYPE," & _
					" COMPANY_CODE," & _
					" POSTING_DATE," & _
					" PERIOD," & _
					" CURRENCY," & _
					" POSTING_KEY," & _
					" BUSINESS_AREA," & _
					" GL_ACCOUNT," & _
					" COST_CENTER," & _
					" INTERNAL_ORDER," & _
					" JOBCODE,TARGETCODE,targettype)" & _
					" WHERE FCBA = '" & Session("FCBA") & "'"
										
if  LEN(TRIM(v_txtYear)) > 0 then
	strErptoSap	=	strErptoSap & " AND TO_NUMBER (TO_CHAR (TO_DATE (POSTING_DATE, 'DD.MM.YYYY'), 'YYYY')) = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strErptoSap	=	strErptoSap & " AND TO_NUMBER (TO_CHAR (TO_DATE (POSTING_DATE, 'DD.MM.YYYY'), 'MM')) = " & v_txtMonth
end if

strErptoSap	=	strErptoSap & " GROUP BY FCBA," & _
							" DOCUMENT_TYPE," & _
							" COMPANY_CODE," & _
							" POSTING_DATE," & _
							" PERIOD," & _
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
<title>iPla'S Reporting - Fico </title>
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
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<!--#include file="menutop.asp"-->
			<div id="header">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
		<div id="subheader" class="centered">
			<h2>INTEGRASI IPLAS - SAP (FICO)</h3>
			<br/>
			<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
		</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td>
<% 
	Dim mon, Yr, MinYr, MaxYr, i, MM
    mon = Month(Session("v_Costbook"))
    If len(v_txtMonth)=0 then v_txtMonth = mon

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
		<td><font face="verdana" size="1">MONTH</font></td>
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
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>

<table border="0" width="100%" id="product-table">
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

<%
Set rsErptoSap = Server.CreateObject("ADODB.Recordset")
rsErptoSap.Open strErptoSap, connstr, adOpenStatic


if rsErptoSap.eof and rsErptoSap.eof then
%>
	<tr><td colspan="39">NO DATA AVAILABLE</td></tr>
<%
else
	NO = 1

	while not rsErptoSap.eof
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass ="default"
		end if
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=rsErptoSap("POSTING_DATE")%></td>
		<td align="center"><%=rsErptoSap("DOCUMENT_TYPE")%></td>
		<td align="center"><%=rsErptoSap("COMPANY_CODE")%></td>
		<td align="center"><%=rsErptoSap("POSTING_DATE")%></td>
		<td align="center"><%=rsErptoSap("PERIOD")%></td>
		<td align="center"><%=rsErptoSap("CURRENCY")%></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="center"><%=rsErptoSap("POSTING_KEY")%></td>
		<td align="center"><%=rsErptoSap("BUSINESS_AREA")%></td>
		<td align="center"><%=rsErptoSap("GL_ACCOUNT")%></td>
		<td align="center"></td>
		<td align="center"></td>
		<td align="right"><%=FormatNumber(rsErptoSap("DOCUMENT_CCY"),0)%></td>
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
