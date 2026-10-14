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
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
<title>ePlaS Reporting - Delivery Order Purchase</title>
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

Code = Request("xCode")
 IsInBridge = Request("IsInBridge")
 Remarks = Request("Remarks")
 
 '---- Data Details
 ' Add by NAF, 2012-12-1
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
 sto_quantity = Request("sto_quantity")
 remarks = Request("remarks")
 delivery_quantity = Request("delivery_quantity")
 do_no = Request("do_no")
 stolineno = Request("stolineno")
 stono = Request("stono")
 StockName = Request("StockName")
 doquantity = Request("doquantity")
 ponumber = Request("ponumber")
 DOLINE = Request("DOLINE")
 stonumber = Request("stonumber")
 do_date_from = Request("do_date_from")
 do_date_to = Request ("do_date_to")

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

txtSQL = 	"  SELECT doh.company," & _
			"       doh.fcba," & _
			"       doh.do_date," & _
			"       doh.do_no," & _
			"       doh.pocode," & _
			"       doh.rowstateid," & _
			"       dod.stono," & _
			"       dod.itemcode," & _
			"       dod.stockname," & _
			"       dod.sto_quantity," & _
			"       dod.delivery_quantity," & _
			"       dod.remarks" & _
			"  FROM deliveryorderpurchasehdr doh" & _
			"       INNER JOIN deliveryorderpurchasedtl dod" & _
			"          ON (doh.do_no = dod.do_no AND doh.fcba = dod.fcba)" 

if  len(trim(fccompanycode)) > 0 then
    txtSQL =txtSQL &  "  where doh.company = '"& fccompanycode &"'"
end if
if  len(trim(businessunit)) > 0 then
    txtSQL =txtSQL &  "  and doh.fcba = '"& businessunit &"'"
end if
if  len(trim(do_no)) > 0 then
    txtSQL =txtSQL &  " and doh.do_no = '"& do_no &"'"
end if
if  len(trim(ponumber)) > 0 then
    txtSQL =txtSQL &  " and doh.pocode = '"& ponumber &"'"
end if
if  len(trim(do_date_from)) > 0 then
    txtSQL = txtSQL & " and do_date >=to_Date('" & do_date_from & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(do_date_from,1)
end if
if  len(trim(do_date_to)) > 0 then
    txtSQL =txtSQL &  " and do_date <=to_Date('" & do_date_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(do_date_to,1)
end if
txtsql	=	txtsql & " Order By doh.do_date, doh.do_no"
txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

' response.Write(txtsql)
' response.Flush()
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
         <th colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>DELIVERY ORDER PURCHASE</h2>
				</div>
			</div>
		 </th>
			</tr>
						<!--<div class="printhide">-->
				<tr>
					<td colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th>Company</th>
					<th>Business Unit</th>
					<th>Date</th>
					<th>Delivery Order</th>
					<th>Purchase Order</th>
					<th>Status</th>
					<th>STO No.</th>
					<th>Item Code</th>
					<th>Item Name</th>
					<th>STO Quantity</th>
					<th>DO Quantity</th>
					<th>Description</th>
				</tr>
                </thead>
               <tbody>
	<%
	if Not rscat.eof then
	While Not rscat.EOF
	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("company")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("do_date")%></td>
			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
			<a href="javascript:launchwin('PrintDeliveryOrder_Purch.Asp?DONo=<%=rscat("do_no")%>&fcba=<%=rscat("fcba")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
			<font face="Verdana" size="1"><%=rscat("do_no")%></font>
			</a>
			</td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pocode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RowStateId")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stono")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("itemcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stockname")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sto_quantity")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("delivery_quantity")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
			
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='20'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
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