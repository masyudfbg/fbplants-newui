<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
<title>ePlaS Reporting - Weighbridge Ticket</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

'--ADDED BY UKIE 14-NOV-2012
v_txtDateout_From= trim(request("Dateout_From"))
v_txtDateout_To= trim(request("Dateout_To"))
v_Posting_Date= trim(request("Posting_Date"))
v_txtPRODUCT_PRODUCTCODE=trim(request("PRODUCT_PRODUCTCODE"))
v_txtCHITNO=trim(request("CHITNO"))
v_txtCUST_CUSTOMERCODE=trim(request("CUST_CUSTOMERCODE"))
'--END ADDING

txtSQL = "SELECT TRANSACTIONTYPE, CHITNO, PRODUCT_PRODUCTCODE, VEHICLENO, CUST_CUSTOMERCODE,customer_api.get_description(cust_customercode,fcba) Custname,FIELD_ESTATE_ESTATECODE,DIVISION,spbno, FIELD_FIELDCODE, DATEIN,TIMEIN, " & _
         " WBRIDGE_WEIGHBRIDGECODEIN, WEIGHTIN, DATEOUT,TIMEOUT,POSTINGDATE, WBRIDGE_WEIGHBRIDGECODEOUT, WEIGHTOUT,abs(WEIGHTIN - WEIGHTOUT ) BRUTO , " & _
		 " GRADINGRECORD_API.GET_BUNCHES(CHITNO,FCBA) BUNCHES, GRADINGRECORD_API.GET_SORTASE(CHITNO,FCBA) SORTASE , " & _
         " abs(WEIGHTIN - WEIGHTOUT) -   GRADINGRECORD_API.GET_SORTASE(CHITNO,FCBA) NETTO  FROM WEIGHBRIDGETICKET" & _
		 " where fcba='" & session("fcBA") & "'"


'--ADDED BY UKIE 14-NOV-2012
if  len(trim(v_txtDateout_From)) > 0 then
    txtSQL = txtSQL & " and DATEOUT >=to_Date('" & v_txtDateout_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDateout_From,1)
end if
if  len(trim(v_txtDateout_To)) > 0 then
    txtSQL =txtSQL &  " and DATEOUT <=to_Date('" & v_txtDateout_To & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDateout_To,1)
end if
'--END ADDING
if  len(trim(v_Posting_Date)) > 0 then
    txtSQL =txtSQL &  " and POSTINGDATE =to_Date('" & v_Posting_Date & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_Posting_Date,1)
end if
if  len(trim(v_txtPRODUCT_PRODUCTCODE)) > 0 then
    txtSQL = txtSQL & " and PRODUCT_PRODUCTCODE='" & v_txtPRODUCT_PRODUCTCODE & "'"
	txtSelect = txtSelect & " PRODUCTCODE " & v_txtPRODUCT_PRODUCTCODE
end if
if  len(trim(v_txtCHITNO)) > 0 then
    txtSQL = txtSQL & " and CHITNO='" & v_txtCHITNO & "'"
	txtSelect = txtSelect & " CHITNO " & v_txtCHITNO
end if
if  len(trim(v_txtCUST_CUSTOMERCODE)) > 0 then
    txtSQL = txtSQL & " and CUST_CUSTOMERCODE='" & v_txtCUST_CUSTOMERCODE & "'"
	txtSelect = txtSelect & " CUSTOMERCODE " & v_txtCUST_CUSTOMERCODE
end if

txtsql	=	txtsql	&	" ORDER BY PRODUCT_PRODUCTCODE, DATEOUT, CHITNO ASC" 
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'response.Write(txtsql)
'response.Flush()
%>
 <body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
		 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="22" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>WEIGHBRIDGE TICKET</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="22" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<!--</div>-->
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>Product Code</th>
					<th>Chit No</th>
					<th>Vehicle No</th>
					<th>Customer Code</th>
					<th>Customer Name</th>
					<th>Date in</th>
					<th>Time in</th>
					<th>Date Out</th>
					<th>Time Out</th>
					<th>Posting Date</th>
					<th>Weighbridge Code in</th>
					<th>Transaction Type</th>
					<% if trim(v_txtPRODUCT_PRODUCTCODE) = "FFB"  then %>
					<th>Estate Code</th>
					<th>SPB No</th>
					<th>Afdeling</th>
					<th>Field Code</th>
					<th>Janjang</th>
					<% END IF%>
					<th>Weighbridge Code Out</th>
					<th>Weight in</th>
					<th>Weight out</th>
					<th>Bruto</th>
					<% if trim(v_txtPRODUCT_PRODUCTCODE) ="FFB" then %>
					<th>Sortase</th>
					<% END IF %>
					<th>Netto</th>
				</tr>
                </thead>
               <tbody>
	<%
	Dim v_WEIGHTIN,v_WEIGHTOUT,v_BRUTO,v_SORTASE,v_NETTO
	v_WEIGHTIN=0
	v_WEIGHTOUT=0
	v_BRUTO=0
	v_SORTASE=0
	v_NETTO=0
	
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCT_PRODUCTCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHITNO")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLENO")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CUST_CUSTOMERCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Custname")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DATEIN")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TIMEIN")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DATEOUT")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TIMEOUT")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("POSTINGDATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WBRIDGE_WEIGHBRIDGECODEIN")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TRANSACTIONTYPE")%></td>
		<% if trim(v_txtPRODUCT_PRODUCTCODE) = "FFB" then %>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FIELD_ESTATE_ESTATECODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SPBNO")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Division")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FIELD_FIELDCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BUNCHES")%></td>
		<% END IF %>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WBRIDGE_WEIGHBRIDGECODEOUT")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("WEIGHTIN"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("WEIGHTOUT"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("BRUTO"),2)%></td>
		<% if trim(v_txtPRODUCT_PRODUCTCODE) = "FFB" then %>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("SORTASE"),2)%></td>
		<% END IF %>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("NETTO"),2)%></td>
		</tr>
	  <%
	  v_WEIGHTIN=v_WEIGHTIN+trim(cdbl(rsCat("WEIGHTIN")))
	  v_WEIGHTOUT=v_WEIGHTOUT+trim(cdbl(rsCat("WEIGHTOUT")))
	  v_BRUTO=v_BRUTO+trim(cdbl(rsCat("BRUTO")))
	  v_SORTASE=v_SORTASE+trim(cdbl(rsCat("SORTASE")))
	  v_NETTO=v_NETTO+trim(cdbl(rsCat("NETTO")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null
	%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
	<% if trim(v_txtPRODUCT_PRODUCTCODE) = "FFB" then %>
	<td align="Center" colspan="18"> <strong><i> <font face="Verdana" size="1">T o t a l</i></strong></td>
	<% ELSE %>
	<td align="Center" colspan="13"> <strong><i> <font face="Verdana" size="1">T o t a l</i></strong></td>
	<% END IF %>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_WEIGHTIN,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_WEIGHTOUT,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_BRUTO,2)%></i></strong></td>
	<% if trim(v_txtPRODUCT_PRODUCTCODE) = "FFB" then %>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_SORTASE,2)%></i></strong></td>
	<% END IF %>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_NETTO,2)%></i></strong></td>
	</tr>
    </tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
	</div>
	</body>
</html>