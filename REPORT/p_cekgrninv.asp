<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtYear, v_txtMonth
	Dim ScriptName
	Dim rsGrnInv, strGrnInv
	Dim v_txtPOCode
	Dim NO, cClass
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	v_txtYear	=	Request("txtYear")
	v_txtMonth	=	Request("txtMonth")

	If Len(Trim(v_txtYear)) = 0 Then
		v_txtYear = Year(Session("v_Costbook"))
	End If

	If Len(Trim(v_txtMonth)) = 0 Then
		v_txtMonth = Month(Session("v_Costbook"))
	End If
	
	v_txtPOCode		=	trim(request("txtPOCode"))
	
	
	strGrnInv		=	" SELECT" & _
						"     FCBA" & _
						"     , NOPO" & _
						"     , STOCKCODE" & _
						"     , STOCK_API.GET_STOCKNAME(STOCKCODE, FCBA) STOCKNAME" & _
						"     , SUM(QTY_GRN) QTY_GRN" & _
						"     , SUM(AMT_GRN) AMT_GRN" & _
						"     , SUM(QTY_INV) QTY_INV" & _
						"     , SUM(AMT_INV) AMT_INV" & _
						"     , SUM(QTY_GRN) - SUM(QTY_INV) QTY_VAR" & _
						"     , SUM(AMT_GRN) - SUM(AMT_INV) AMT_VAR" & _
						" FROM (" & _
						"     SELECT " & _
						"         RND.FCBA FCBA" & _
						"         , RN.POCODE NOPO" & _
						"         , RN.RNDATE TANGGAL" & _
						"         , RND.PURCHASEITEMCODE STOCKCODE" & _
						"         , NVL(RND.QUANTITY, 0) QTY_GRN" & _
						"         , NVL(RND.AMOUNTRECEIVED, 0) AMT_GRN" & _
						"         , 0 QTY_INV" & _
						"         , 0 AMT_INV" & _
						"     FROM" & _
						"         RECEIVENOTEDETAIL RND," & _
						"         RECEIVENOTE RN" & _
						"     WHERE (RND.FCBA = RN.FCBA AND RND.RECEIVENOTECODE = RN.RECEIVENOTECODE)" & _
						"     UNION ALL" & _
						"     SELECT" & _
						"         INVD.FCBA FCBA" & _
						"         , INVD.POCODE NOPO" & _
						"         , INV.INVOICEDATE TANGGAL" & _
						"         , INVD.ITEMCODE STOCKCODE" & _
						"         , 0 QTY_GRN" & _
						"         , 0 AMT_GRN" & _
						"         , NVL(INVD.QUANTITY, 0) QTY_INV" & _
						"         , NVL(INVD.AMOUNT, 0) AMT_INV" & _
						"     FROM" & _
						"         INVOICEDETAIL INVD," & _
						"         INVOICE INV" & _
						"     WHERE (INVD.FCBA = INV.FCBA AND INVD.INVOICECODE = INV.INVOICECODE)" & _
						"     ) WHERE FCBA = '"& Session("FCBA") & "'" & _
						"             AND TO_NUMBER(TO_CHAR(TANGGAL, 'YYYY')) = " & v_txtYear & "" & _
						"             AND TO_NUMBER(TO_CHAR(TANGGAL, 'MM')) = " & v_txtMonth & ""

if  len(trim(v_txtPOCode)) > 0 then
    strGrnInv	=	strGrnInv & "AND NOPO = '"& v_txtPOCode &"'"
end if

	strGrnInv	=	strGrnInv &	" GROUP BY" & _
								"     FCBA" & _
								"     , NOPO" & _
								"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv &	" ORDER BY" & _
								"     FCBA" & _
								"     , NOPO" & _
								"     , STOCKCODE"

	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Crosscek GRN vs Inv</title>
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
            	<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
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
						<td>
							<font face="verdana" size="1">No. PO</font>
						</td>
						<td>
							<select name="txtPOCode" >
								<option value="">--ALL--</option>
								<%
								Dim rsPOCode, strPOCode
								strPOCode	=	" SELECT POCODE FROM PURCHASEORDER WHERE FCBA = '" & Session("FCBA") & "' ORDER BY POCODE"
								
								Set rsPOCode = Server.CreateObject("ADODB.Recordset")
								rsPOCode.Open strPOCode, connstr, adOpenStatic 
								
								if not rsPOCode.eof then
								Do While Not rsPOCode.eof
								if rsPOCode("POCODE") = v_txtPOCode then%>
								<option value="<%=rsPOCode("POCODE")%>" Selected><%=rsPOCode("POCODE")%></option> 
								<%Else%>
								<option value="<%=rsPOCode("POCODE")%>"><%=rsPOCode("POCODE")%></option>  
								<%End if
									rsPOCode.movenext
								loop
								end if
								rsPOCode.Close
								set rsPOCode = Nothing
								%>
							</select>
						</td>
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
			
			</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="10" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
			
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>Crosscek GRN vs Invoice</h2>
                <br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
		</div>
</th>
		</tr>
	<tr>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">NO PO</th>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">KODE BARANG</th>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">NAMA BARANG</th>
			<th colspan="2" style="background-color:#00A300;color:#fff;">GRN</th>
			<th colspan="2" style="background-color:#00A300;color:#fff;">INVOICE</th>
			<th colspan="2" style="background-color:#00A300;color:#fff;">SELISIH</th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;">QTY</th>
			<th style="background-color:#00A300;color:#fff;">RP.</th>
			<th style="background-color:#00A300;color:#fff;">QTY</th>
			<th style="background-color:#00A300;color:#fff;">RP.</th>
			<th style="background-color:#00A300;color:#fff;">QTY</th>
			<th style="background-color:#00A300;color:#fff;">RP.</th>
		</tr>
    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="9" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<td align="left"><%=rsGrnInv("NOPO")%></td>
			<td align="left"><%=rsGrnInv("STOCKCODE")%></td>
			<td align="left"><%=rsGrnInv("STOCKNAME")%></td>
			<td align="right"><%=FormatNumber(rsGrnInv("QTY_GRN"), 2)%></td>
			<td align="right"><%=FormatNumber(rsGrnInv("AMT_GRN"), 2)%></td>
			<td align="right"><%=FormatNumber(rsGrnInv("QTY_INV"), 2)%></td>
			<td align="right"><%=FormatNumber(rsGrnInv("AMT_INV"), 2)%></td>
			<td align="right"><%=FormatNumber(rsGrnInv("QTY_VAR"), 2)%></td>
			<td align="right"><%=FormatNumber(rsGrnInv("AMT_VAR"), 2)%></td>
		</tr>
<%
			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
	</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->
