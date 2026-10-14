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
<title>ePlaS Reporting - Stock Transfer Order </title>
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


 LINENOKEY=Request("LineNoKey")
 ' End Added by NAF
 ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 OTHERITEMDESC=request("OTHERITEMDESC")
 RowStateId = Request("RowStateId") 
 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID")
 fccompanycode = Request("fccompanycode")
 businessunit = Request("businessunit")
 rentby = Request("rentby")
 transportercode = Request("transportercode")
 transportername = Request("transportername")
 transporteraddress = Request("transporteraddress")
 termofpayment = Request("termofpayment")
 dp = Request("dp")
 ppn = Request("ppn")
 stono = Request("stono")
 StockName = Request("StockName")
 unitprice = Request("unitprice")
 totalamount = Request("totalamount")
 doquantity = Request("doquantity")
 ponumber = Request("ponumber")
 stolineno = Request("stolineno")
 stodate_from  = Request("stodate_from")
 stodate_no  = Request("stodate_no")

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

txtSQL = 	"	SELECT sth.company,"  & _
			"       sth.fcba,"  & _
			"       sth.stodate,"  & _
			"       sth.stono,"  & _
			"       sth.pocode,"  & _
			"       sth.rowstateid,"  & _
			"       tr.fcname,"  & _
			"       sth.rentby,"  & _
			"       sth.termofpayment,"  & _
			"       std.itemcode,"  & _
			"       std.stockname,"  & _
			"       std.quantityorder,"  & _
			"       nvl(std.do_quantity,0) do_quantity,"  & _
			"       std.unitprice,"  & _
			"       sth.ppnrp,"  & _
			"       sth.totalamount"  & _
			"  FROM stocktransferorderhdr sth"  & _
			"       INNER JOIN stocktransferorderdtl std"  & _
			"          ON (sth.stono = std.stono AND sth.fcba = std.fcba) " & _ 
			" 		INNER JOIN transporter tr on (tr.fccode=sth.TRANSPORTERCODE and tr.FCBA=sth.FCBA) "  

if  len(trim(fccompanycode)) > 0 then
    txtSQL =txtSQL &  "  where sth.company = '"& fccompanycode &"'"
end if
if  len(trim(businessunit)) > 0 then
    txtSQL =txtSQL &  "  and sth.fcba = '"& businessunit &"'"
end if
if  len(trim(stono)) > 0 then
    txtSQL =txtSQL &  " and sth.stono = '"& stono &"'"
end if
if  len(trim(ponumber)) > 0 then
    txtSQL =txtSQL &  " and sth.pocode = '"& ponumber &"'"
end if
if  len(trim(transportercode)) > 0 then
    txtSQL =txtSQL &  " and sth.transportercode = '"& transportercode &"'"
end if
if  len(trim(rentby)) > 0 then
    txtSQL =txtSQL &  " and sth.rentby = '"& rentby &"'"
end if
if  len(trim(termofpayment)) > 0 then
    txtSQL =txtSQL &  " and sth.termofpayment = '"& termofpayment &"'"
end if
if  len(trim(stodate_from)) > 0 then
    txtSQL = txtSQL & " and stodate >=to_Date('" & stodate_from & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(stodate_from,1)
end if
if  len(trim(stodate_to)) > 0 then
    txtSQL =txtSQL &  " and stodate <=to_Date('" & stodate_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(stodate_to,1)
end if
txtsql	=	txtsql & " Order By sth.stodate,sth.stono"
txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

'' response.Write(txtsql)
response.Flush()
Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open txtSQL, connstr
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
					<h2>STOCK TRANSFER ORDER </h2>
				</div>
			</div>
		 </th>
			</tr>
						<!--<div class="printhide">-->
				<tr>
					<td colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th></th>
					<th>Company</th>
					<th>Business Unit</th>
					<th>Date</th>
					<th>STO No.</th>
					<th>PO No.</th>
					<th>Transporter</th>
					<th>Rent By</th>
					<th>Term Of Payment</th>
					<th>Item Code</th>
					<th>Item Name</th>
					<th>Quantity</th>
					<th>Delivery Quantity</th>
					<th>Unit Price</th>
					<th>PPN (RP)</th>
					<th>Amount</th>
				</tr>
                </thead>
               <tbody>
	<%
	if Not rscat.eof then
		v_quantityorder = 0
		v_do_quantity = 0
		v_unitprice = 0
		v_totalamount = 0
	While Not rscat.EOF

					dim ttl 
					ttl =cdbl(rscat("quantityorder"))*cdbl(rscat("unitprice"))+cdbl(rscat("ppnrp"))
	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
				<a href="printsto.asp?fcba=<%=businessunit%>&stono=<%=rscat("stono")%>">
								<img src="IMAGES/print.gif" alt="Print Data : <%=rscat(1)%>" border="0">
								</a>
			</td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("company")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stodate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stono")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pocode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rentby")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("termofpayment")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("itemcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stockname")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("quantityorder")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("do_quantity")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("unitprice"),0)%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("ppnrp"))%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(ttl,0)%></td>
		</tr>
	  <%
	   v_quantityorder =v_quantityorder+trim(cdbl(rsCat("quantityorder")))
	   v_unitprice =v_unitprice+Trim(cdbl(rsCat("unitprice")))
	   v_totalamount =v_totalamount+Trim(cdbl(rsCat("totalamount")))
	   v_do_quantity = cdbl(v_do_quantity) + Trim(cdbl(rsCat("do_quantity")))
	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='14'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
		<td align="Center" style="<%=cmdstyleLeft%>" colspan="12"> <font face="Verdana" size="1">T o t a l</td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_quantityorder,0)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_do_quantity,0)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_unitprice,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_totalamount,2)%></td>
	</tr>
	</tbody>
</table>
 <table width="90%" class="data">
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=MgrName%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=KTUName%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
       </tbody>
   </table>
   </body>
   </html>