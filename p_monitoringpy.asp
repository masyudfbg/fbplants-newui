<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>FBPlants Reporting - Monitoring Invoice Belum Diterima</title>
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
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtPODATE_From=trim(request("PODATE_From"))
v_txtPODATE_To=trim(request("PODATE_To"))
v_txtPOCODE=trim(request("POCODE"))
v_txtITEMCODE=trim(request("ITEMCODE"))
v_txtRequestorID=trim(request("REQUESTORID"))
v_txtRowStateId=trim(request("ROWSTATEID"))
compmonipy=trim(request("compmonipy"))
businessunit=trim(request("businessunitmonpy"))

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if

if trim(compmonipy) = "all" then 
		filtercompany = " WHERE b.PLANTATION = 'FBG'"
	else
		if trim(businessunit) = "-" then 
	    	filtercompany = " WHERE b.fccompanycode ='" & compmonipy & "' "
		else 
			filtercompany = " WHERE b.fccode ='" & businessunit & "' "
		end if
end if

txtSQL = " SELECT " & _ 
    " B.FCCOMPANYCODE, " & _ 
    " PO.FCBA, " & _ 
    " S.FCCODE || ' - ' || S.FCNAME FCNAME, " & _ 
    " PO.PODATE, " & _ 
    " PO.POCODE, " & _ 
    " PY.INVOICEDATE AS TANGGALPY, " & _ 
    " PY.INVOICECODE AS NOPY, " & _ 
    " I.INVOICEDATE AS INVOICEPODATE, " & _ 
    " I.INVOICECODE AS INVOICEPOCODE, " & _ 
    " NVL(I.TOTALAMOUNT,0) AS TOTALAMOUNT, " & _ 
    " ( " & _ 
    "    SELECT LISTAGG(PD_UNIQ.VOUCHERCODE, ', ') WITHIN GROUP (ORDER BY PD_UNIQ.VOUCHERCODE) " & _ 
    "    FROM ( " & _ 
    "        SELECT DISTINCT PD.VOUCHERCODE " & _ 
    "        FROM PAYMENTVOUCHERDETAIL PD " & _ 
    "        WHERE (PD.INVOICECODE = PY.INVOICECODE) " & _ 
    "          OR " & _ 
    "          (PD.INVOICECODE = I.INVOICECODE) " & _ 
    "    ) PD_UNIQ " & _ 
    " ) AS ALLVOUCHERCODE, " & _ 
    " ( " & _ 
  " SELECT LISTAGG(TO_CHAR(P.DATECREATED, 'DD-MM-YYYY'), ', ')  " & _ 
  "       WITHIN GROUP (ORDER BY P.DATECREATED) " & _ 
  " FROM PAYMENTVOUCHERDETAIL PD " & _ 
  " JOIN PAYMENTVOUCHER P ON P.VOUCHERCODE = PD.VOUCHERCODE " & _ 
  " WHERE PD.INVOICECODE IN (PY.INVOICECODE, I.INVOICECODE) " & _ 
" ) AS DATECREATED, " & _ 
" ( " & _ 
"    SELECT NVL(SUM(PDD.AMOUNT),0) " & _ 
"    FROM PAYMENTVOUCHERDETAIL PDD " & _ 
"    JOIN PAYMENTVOUCHER PP ON PP.VOUCHERCODE = PDD.VOUCHERCODE " & _ 
"    WHERE PDD.INVOICECODE IN (PY.INVOICECODE, I.INVOICECODE) " & _ 
"  ) AS AMOUNTVOUCHER " & _
" FROM PURCHASEORDER PO " & _ 
" LEFT JOIN INVOICEDPHDR PY " & _ 
"    ON PY.POCODE = PO.POCODE AND PY.FCBA = PO.FCBA " & _ 
" LEFT JOIN INVOICE I " & _ 
"    ON I.POCODE = PO.POCODE AND I.FCBA = PO.FCBA " & _ 
" LEFT JOIN SUPPLIER S ON S.FCCODE=PO.SUPPLIERCODE AND S.FCBA=PO.FCBA " & _ 
" LEFT JOIN BUSINESSUNIT B " & _ 
"    ON PO.FCBA = B.FCCODE " & _ 
filtercompany & _
" AND PO.ROWSTATEID in('Ordered','Confirmed')" 


if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " AND PO.PODATE >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtPODATE_From,1)	
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and PO.PODATE <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtPODATE_To,1)	
end if


	txtSQL = txtSQL & 	" ORDER BY B.FCCOMPANYCODE,PY.FCBA, PO.PODATE"


if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'' response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
		 <!--#include file="print_excel.asp"-->
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>MONITORING PO vs PY vs INVOICE</h2>
				</div>
			</div>
					
		</th>
        </tr>
        		<!--<div class="printhide">-->
						<tr>
							<th colspan="14"  style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
							<%	
							
							Set rsCat = Server.CreateObject("ADODB.Recordset")
							rsCat.CursorType = 3
							rsCat.CursorLocation = 3
							rsCat.Open txtSQL, connstr

							
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr>
						<tr>
							<!--<th>No.</th>-->
							<th width="">No</th>
							<th>Company</th>
							<th>FCBA</th>
							<th>Supplier</th>
							<th>Tanggal PO</th>
							<th>No PO</th>
							<th>Tanggal PY</th>
							<th>No PY</th>
							<th>Tanggal Invoice</th>
							<th>No Invoice</th>
							<th>Jumlah Invoice</th>
							<th>Payment</th>
							<th>Tanggal Payment</th>
							<th>Jumlah Payment</th>
						</tr>
                        </thead>
                        <tbody>
							<%
							Dim txtno
							txtno=1
							While Not rsCat.EOF%>

							<tr>
								
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=txtno%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCCOMPANYCODE")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCBA")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCNAME")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PODATE")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("POCODE")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TANGGALPY")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NOPY")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("INVOICEPODATE")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("INVOICEPOCODE")%> </td>
								<td align="RIGHT" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("TOTALAMOUNT"),0)%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ALLVOUCHERCODE")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DATECREATED")%> </td>
								<td align="RIGHT" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("AMOUNTVOUCHER"),0)%> </td>
								
							</tr>
							  <%

							  rsCat.MoveNext
							  'response.Flush()
							  txtno=txtno+1
							Wend
							rsCat.close()
							rsCat=null%>
							<input  type="hidden" value="<%=txtTable%>" name="txtTable">
							<input  type="hidden" value="<%=txtModule%>" name="txtModule">
							</form>
                            </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	