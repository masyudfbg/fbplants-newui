<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtSupplierCode, v_txtInvoiceCode

v_txtSupplierCode	=	trim(Request("txtSupplierCode"))
v_txtInvoiceCode	=	trim(Request("txtInvoiceCode"))
ScriptName			=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
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
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
		<!--#include file="menutop.asp"-->
		<table id="form-table" class="centered-margin">
			<tr>
				<form name="frmAR" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<td><font face="verdana" size="1">SUPPLIER</font></td><td>
				<select name="txtSupplierCode" class="chosen-select"><option value=""></option>
				<%
				Dim rsSupplierCode, strSupplierCode
				strSupplierCode	=	" SELECT FCCODE, FCNAME FROM SUPPLIER WHERE ACTIVATION = 'Y' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
				
				Set rsSupplierCode = Server.CreateObject("ADODB.Recordset")
				rsSupplierCode.Open strSupplierCode, connstr, adOpenStatic 
				
				if not rsSupplierCode.eof then
				Do While Not rsSupplierCode.eof
				if rsSupplierCode("FCCODE") = v_txtSupplierCode then%>
						<option value="<%=rsSupplierCode("FCCODE")%>" Selected><%=rsSupplierCode("FCCODE")%>-<%=rsSupplierCode("FCNAME")%></option> 
					 <%Else%>
						<option value="<%=rsSupplierCode("FCCODE")%>"><%=rsSupplierCode("FCCODE")%>-<%=rsSupplierCode("FCNAME")%></option>  
					<%End if
					rsSupplierCode.movenext
				loop
				end if
				rsSupplierCode.Close
				set rsSupplierCode = Nothing
				%>
				</select></td>
				<td><font face="verdana" size="1">INVOICE CODE</font></td><td><input type="text" name="txtInvoiceCode" value="<%=v_txtInvoiceCode%>" /></td>
				<td><input type="submit" value="View" /></td>
				</div>
				</form>
			</tr>
		</table>
        </div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered" style="width:50%; margin:auto;">
			<h2>LAPORAN DETAIL AGING</h2>			 
		</div>
	</div>
    </th>
  </tr>
  
	<tr>
		<th rowspan="2" style="color:#fff;">KODE SUPPLIER</th>		
		<th rowspan="2" style="color:#fff;">NAMA SUPPLIER</th>
		<th rowspan="2" style="color:#fff;">INVOICE CODE</th>
		<th rowspan="2" style="color:#fff;">INVOICE DATE</th>
		<th rowspan="2" style="color:#fff;">TOTAL INVOICE</th>
		<th rowspan="2" style="color:#fff;">NET BALANCE</th>		
		<th colspan="4" style="color:#fff;">AGING</th>
	</tr>
	<tr>		
		<th style="color:#fff;">0 - 30 Hari</th>
		<th style="color:#fff;">31 - 45 Hari</th>
		<th style="color:#fff;">45 - 90 Hari</th>
		<th style="color:#fff;">> 90 Hari</th>	
	</tr>
    </thead>
    <tbody>
	<%
	Dim rsAging, strAging , noline
	Dim nTotalAmt, nNetbalance, nPosition_0, nPosition_1, nPosition_2, nPosition_3
	
	strAging	=	" SELECT A.SUPPLIERCODE, C.FCNAME SUPPLIERNAME, A.INVOICECODE, B.TOTALAMOUNT, A.INVOICEDATE, A.NETBALANCE," & _
					" DECODE (GETAGINGPOSITIONV2 (A.INVOICEDATE), 0, A.NETBALANCE,0) POSITION_0," & _
					" DECODE (GETAGINGPOSITIONV2 (A.INVOICEDATE), 1, A.NETBALANCE,0) POSITION_1," & _
					" DECODE (GETAGINGPOSITIONV2 (A.INVOICEDATE), 2, A.NETBALANCE,0) POSITION_2," & _
					" DECODE (GETAGINGPOSITIONV2 (A.INVOICEDATE), 3, A.NETBALANCE,0) POSITION_3" & _
					" FROM INVOICE A, (  SELECT INVOICECODE, SUM (AMOUNT) TOTALAMOUNT" & _
					" FROM INVOICEDETAIL WHERE FCBA = '" & Session("FCBA") & "'" & _
					" GROUP BY INVOICECODE) B, SUPPLIER C WHERE A.SUPPLIERCODE = C.FCCODE AND A.FCBA = C.FCBA" & _
					" AND A.INVOICECODE = B.INVOICECODE AND A.NETBALANCE <> 0" & _
					" AND A.FCBA = '" & Session("FCBA") & "'"	 

if  len(trim(v_txtSupplierCode)) > 0 then
    strAging	=	strAging &	" AND A.SUPPLIERCODE ='" & v_txtSupplierCode & "'"
end if
if  len(trim(v_txtInvoiceCode)) > 0 then
    strAging	=	strAging &	" AND A.INVOICECODE ='" & v_txtInvoiceCode & "'"
end if

	strAging	=	strAging &	" ORDER BY SUPPLIERCODE ASC, INVOICEDATE DESC"
	
	Set rsAging = Server.CreateObject("ADODB.Recordset")
	rsAging.Open strAging, connstr, adOpenStatic
    
	' response.write strAging
	if Not rsAging.eof then
	
		nTotalAmt	=	0
		nNetbalance	=	0
		nPosition_0	=	0
		nPosition_1	=	0
		nPosition_2	=	0
		nPosition_3	=	0
		
		while Not rsAging.eof  %>
		           
		  <tr>
			<td align="center"><%=rsAging("SUPPLIERCODE")%></td>
			<td align="left"><%=rsAging("SUPPLIERNAME")%></td>
			<td align="left"><%=rsAging("INVOICECODE")%></td>
			<td align="left"><%=rsAging("INVOICEDATE")%></td>
			<td align="right"><%=formatnumber( cdbl(rsAging("TOTALAMOUNT")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsAging("NETBALANCE")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsAging("POSITION_0")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsAging("POSITION_1")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsAging("POSITION_2")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsAging("POSITION_3")),0)%></td>
		  </tr>	
		      
<%
			nTotalAmt	=	nTotalAmt	+	cdbl(rsAging("TOTALAMOUNT"))	
			nNetbalance	=	nNetbalance	+	cdbl(rsAging("NETBALANCE"))
			nPosition_0	=	nPosition_0	+	cdbl(rsAging("POSITION_0"))
			nPosition_1	=	nPosition_1	+	cdbl(rsAging("POSITION_1"))
			nPosition_2	=	nPosition_2	+	cdbl(rsAging("POSITION_2"))
			nPosition_3	=	nPosition_3	+	cdbl(rsAging("POSITION_3"))
			
			rsAging.MoveNext
		wend
	end if
%>
		  <tr style="background-color:#F4F4A0;">
			<td colspan="4" align="center"><strong>TOTAL</strong></td>
			<td align="right"><strong><%=formatnumber(nTotalAmt,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nNetbalance,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nPosition_0,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nPosition_1,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nPosition_2,0)%></strong></td>
			<td align="right"><strong><%=formatnumber(nPosition_3,0)%></strong></td>
		  </tr>	
          </tbody>
</table>
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
<!--#include file="bottom.Asp"-->