<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Report - Sales Contract</title>
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

<%
d_span = 10
creationdate_from = request("creationdate_from")
creationdate_to = request("creationdate_to")
salescontractno = request("salescontractno")
customercode = request("customercode")
productcode = request("productcode")

txtSelect = ""

businessunit= trim(request("businessunit"))

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

txtSQL = " SELECT ddata.transportercontractno" & _
         "        , ddata.creationdate" & _
         "        , ddata.transporter" & _
         "        , ddata.startdate" & _
         "        , ddata.enddate" & _
         "        , ddata.shippingaddress" & _
         "        , nvl(ddata.rate,0) rate" & _
         "        , ddata.unitofmeasure" & _
         "        , nvl(ddata.losstolerance,0) losstolerance" & _
         "        , tr.fcname transportername" & _
		 " FROM transportercontract ddata" & _
		 " LEFT JOIN transporter tr ON (ddata.fcba = tr.fcba AND ddata.transporter = tr.fccode)" & _
		 " WHERE ddata.creationdate BETWEEN TO_DATE('" & creationdate_from & "','mm/dd/yyyy') AND TO_DATE('" & creationdate_to & "','mm/dd/yyyy')" & _
		 "       AND ddata.fcba = '" & businessunit & "'"

if cstr(""&creationdate_from) = cstr(""&creationdate_to) then
	txtSelect = txtSelect & " Date : " & Day(creationdate_to) & " " & MonthName(Month(creationdate_to)) & " " & Year(creationdate_to)
else
	txtSelect = txtSelect & " From " & Day(creationdate_from) & " " & MonthName(Month(creationdate_from)) & " " & Year(creationdate_from)
	txtSelect = txtSelect & " To " & Day(creationdate_to) & " " & MonthName(Month(creationdate_to)) & " " & Year(creationdate_to)
end if

if cstr(""&transportercontractno) <> "" then
    txtSQL = txtSQL & " AND ddata.transportercontractno = '" & transportercontractno & "'"
	txtSelect = txtSelect & " Contract No " & transportercontractno
end if

if cstr(""&transporter) <> "" then
    txtSQL = txtSQL & " AND ddata.transporter = '" & transporter & "'"
	txtSelect = txtSelect & " Transporter " & transporter
end if

if len(trim(cstr(""&txtSelect))) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if

arr_data_eof = true
Set rsCat = connStr.Execute(txtSQL)
if Not rsCat.eof then
	arr_data = rsCat.getrows()
	arr_data_eof = false
end if
%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
		<tr>
			<th colspan="<%=d_span%>" style=" background-color:#FFFFFF; color:#000000; border:none">
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" style="width:50%; margin:auto;">
						<h2>Transporter Contract</h2>
					</div>
				</div>
			</th>
		</tr>
		<tr>
			<th colspan="<%=d_span%>" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
		</tr>
		<tr>
			<th>Contract No</th>
			<th>Creation Date</th>
			<th>Transporter</th>
			<th>Description</th>
			<th>Start</th>
			<th>End</th>
			<th>Shipping Address</th>
			<th>Price</th>
			<th>Uom</th>
			<th>Loss Tolerance (%)</th>
		</tr>
		</thead>
		<tbody>
<%
		if not arr_data_eof then
			for xyy=0 to ubound(arr_data,2)
				response.write "<tr>"
				response.write "	<td align='center'>"&arr_data(0,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(1,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(2,xyy)&"</td>"
				response.write "	<td align='left'>"&arr_data(9,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(3,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(4,xyy)&"</td>"
				response.write "	<td align='left'>"&arr_data(5,xyy)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(arr_data(6,xyy),2)&"</td>"
				response.write "	<td align='center'>"&arr_data(7,xyy)&"</td>"
				response.write "	<td align='center'>"&arr_data(8,xyy)&"</td>"
				response.write "</tr>"
			next
		else
			response.write "<tr><td align='left' colspan='"&d_span&"'>No Data Available</td>"
		end if
		
%>
		</tbody>
		</table>
		</div>
		<input  type="hidden" value="<%=txtTable%>" name="txtTable">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		</form>
	</body>
</html>