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
<title>ePlaS Reporting - Vehicle Activity</title>
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

Dim Amount
StrAmount  = " NVL(vehiclecb_api.get_amount(vh.vehiclecode, vh.jobcode, vh.vehicledate, vh.locationtype, vh.locationcode, vh.unit, vh.fcba),0)"

v_VEHICLEDATE_From=trim(request("VEHDATE_From"))
v_VEHICLEDATE_To=trim(request("VEHDATE_to"))
v_JOBCODE=trim(request("JOBCODE"))
v_LOCATIONTYPE=trim(request("LOCATIONTYPE"))
v_LocationCode=trim(request("LocationCode"))
v_VEHICLEGROUPCODE=trim(request("VEHICLEGROUPCODE"))
v_VEHICLECODE=trim(request("VEHICLECODE"))
categorytransport=request("categorytransport")
EMPLOYEECODE=trim(request("EMPLOYEECODE"))
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

if request("txtKind") = "VehicleRunningAccount" Then
	txtSQL = " SELECT vh.vehiclecode" & _
		     "        , vm.fcname vehiclename" & _
			 "		  , vg.vehiclegroupcode " & _
			 "		  , vg.description" & _
	         "        , vh.vehicledate" & _
	         "        , vh.time" & _
	         "        , vh.locationtype" & _
	         "        , vh.locationcode" & _
	         "        , vh.EMPLOYEECODE" & _
			 "        , emp.fcname employeename" & _
	         "        , vh.helper1" & _
	         "        , vh.helper2" & _
	         "        , vh.jobcode" & _
			 "        , vh.chargeunit" & _
	         "        , vh.chargejob" & _
	         "        , vh.chargetype" & _
	         "        , vh.chargecode" & _
	         "        , vac.fcname categorytransport" & _
	         "        , j.fcname" & _
	         "        , vh.ref" & _
	         "        , vh.remarks" & _
			 "        , vh.type_kontanan" & _
	         "        , vh.hours" & _
	         "        , vh.unit" & _
	         "        , vh.output" & _
	         "        , vh.trip" & _
	         "        , " & StrAmount & " AS amount" & _
             " FROM vehicleactivity vh, job j, vehicleactivitycategory vac,  Employee emp, vehicle vm, vehiclegroup vg" & _
             " WHERE (vh.jobcode = j.fccode AND vh.fcba = j.fcba) " & _
             "       AND vac.fccode = vh.categorytransport AND vh.employeecode = emp.fccode" & _
             " 		 AND vh.vehiclecode = vm.fccode" & _
			 "  	 AND vg.vehiclegroupcode = vm.vehiclegroupcode" & _
			 "		 AND vg.fcba = vm.fcba " & _
			 "		 AND vh.fcba = vm.fcba" & _
			 "       AND vh.fcba = '" & businessunit & "'"
Else
	txtSQL = " SELECT vh.vehiclecode" & _
	         "        , vm.fcname vehiclename" & _
			 "		  , vg.vehiclegroupcode " & _
			 "		  , vg.description" & _
	         "        , vh.vehicledate" & _
	         "        , vh.time" & _
	         "        , vh.locationtype" & _
	         "        , vh.locationcode" & _
	         "        , vh.EMPLOYEECODE" & _
			 "        , emp.fcname employeename" & _
	         "        , vh.helper1" & _
	         "        , vh.helper2" & _
	         "        , vh.jobcode" & _
			 "        , vh.chargeunit" & _
	         "        , vh.chargejob" & _
	         "        , vh.chargetype" & _
	         "        , vh.chargecode" & _
	         "        , vac.fcname categorytransport" & _
	         "        , j.fcname" & _
	         "        , vh.ref" & _
	         "        , vh.remarks" & _
			 "        , vh.type_kontanan" & _
	         "        , vh.hours" & _
	         "        , vh.unit" & _
	         "        , vh.output" & _
	         "        , vh.trip" & _
	         "        , vh.amount" & _
             " FROM vehicleactivity vh, job j, vehicleactivitycategory vac, Employee emp, vehicle vm, vehiclegroup vg" & _
             " WHERE (vh.jobcode = j.fccode AND vh.fcba = j.fcba) " & _
             "       AND vac.fccode = vh.categorytransport AND vh.employeecode = emp.fccode" & _
             " 		 AND vh.vehiclecode = vm.fccode" & _
             " 		 AND vh.fcba = emp.fcba" & _
             "		 AND vh.fcba = vm.fcba" & _
			 "  	 AND vg.vehiclegroupcode = vm.vehiclegroupcode" & _
			 "		 AND vg.fcba = vm.fcba " & _
			 "       AND vh.fcba = '" & businessunit & "'"
End If
	 
if  len(trim(v_VEHICLEDATE_From)) > 0 then
    txtSQL = txtSQL & " AND vh.VEHICLEDATE >= TO_DATE('" & v_VEHICLEDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_VEHICLEDATE_From,1)
end if
if  len(trim(v_VEHICLEDATE_To)) > 0 then
    txtSQL =txtSQL &  " AND vh.VEHICLEDATE <= TO_DATE('" & v_VEHICLEDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_VEHICLEDATE_To,1)
end if
if  trim(v_JOBCODE) <> "" then
    txtSQL = txtSQL & " AND vh.JOBCODE = '" & v_JOBCODE & "'"
	txtSelect = txtSelect & " Job Code " & v_JOBCODE
end if
if  trim(v_LOCATIONTYPE) <> "" then
    txtSQL =txtSQL &  " AND vh.LOCATIONTYPE = '" & v_LOCATIONTYPE & "'"
	txtSelect = txtSelect & " Target Type " & v_LOCATIONTYPE
end if 
' if  trim(v_LocationCode) <> "" then
'     txtSQL =txtSQL &  " AND vh.LOCATIONCODE = '" & v_LocationCode & "'"
' 	txtSelect = txtSelect & " Target Code " & v_LocationCode
' end if 
' if  len(trim(v_VEHICLEGROUPCODE)) > 0 then
'     txtSQL =txtSQL &  " AND UPPER(vh.VEHICLECODE) IN (select FCCODE from Vehicle where VEHICLEGROUPCODE like UPPER('%" & v_VEHICLEGROUPCODE & "%'))"
' 	txtSelect = txtSelect & " Vehicle Group " & v_VEHICLEGROUPCODE
' end if 
if  v_VEHICLECODE <> "-" And len(trim(v_VEHICLECODE)) <> 0 then
    txtSQL =txtSQL &  " AND vh.VEHICLECODE = '" & v_VEHICLECODE & "'"
	txtSelect = txtSelect & " Vehicle Code " & v_VEHICLECODE
end if 
' if  len(trim(EMPLOYEECODE)) > 0 then
'     txtSQL =txtSQL &  " AND (UPPER(vh.EMPLOYEECODE) like UPPER ('%" & EMPLOYEECODE & "%') OR emp.fcname like UPPER ('%"& EMPLOYEECODE &"%'))"
' 	txtSelect = txtSelect & " Employee Code " & EMPLOYEECODE
' end if 
if  categorytransport <> "-" And categorytransport <> ""  then
    txtSQL = txtSQL &  " AND vh.categorytransport = '" & categorytransport & "'"
end if 

txtsql=(txtsql) & " GROUP BY vh.vehiclecode, vm.fcname, vg.vehiclegroupcode, vg.description, vh.vehicledate, vh.time, vh.locationtype, " & _
				  " vh.locationcode, vh.EMPLOYEECODE, emp.fcname, vh.helper1, vh.helper2, vh.jobcode, vh.chargeunit, vh.chargejob, " & _
				  "	vh.chargetype, vh.chargecode, vac.fcname, j.fcname, vh.ref, vh.remarks, vh.type_kontanan, vh.hours, " & _
				  "	vh.unit, vh.output, vh.trip, vh.amount ORDER BY vh.vehicledate, vh.time,vh.vehiclecode "
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="25" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>VEHICLE ACTIVITY</h2>
				</div>
			</div>
				</th>
			</tr>
			<!--<div class="printhide">-->
			<tr>
				<th colspan="25" style=" background-color:#FFFFFF; color:#000000; border:none"><%=txtSelect%></td>
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
						NO=1
					%>
				</tr>
				<tr>
					<th>Code</th>
					<th>Vehicle Name</th>
					<th>Vehicle Group</th>
					<th>Employee Code</th>
					<th>Helper 1</th>
					<th>Helper 2</th>
					<th>Date</th>
					<th>Time</th>
					<th>Job<br>Code</th>
					<th>Job<br>Description</th>
					<th>Target<br>Type</th>
					<th>Target<br>Code</th>
					<th>Category<br>Transport</th>
					<th>Type Kontanan</th>
					<th>Charge Unit</th>
					<th>Charge Job</th>
					<th>Charge Type</th>
					<th>Charge Code</th>
					<th>Trip</th>
					<th>Hours</th>
					<th>Km/Hm</th>
					<th>Output</th>
					<th>Amount</th>
					<th>Reference</th>
					<th>Remarks</th>	
				</tr>
				</thead>
	<tbody>
	<%	
		Dim v_Quantity,v_qtyretrun,v_amountreceived,v_amountreturn
		v_HOURS = 0
		v_UNIT = 0
		v_OUTPUT = 0
		v_AMOUNT = 0
		v_trip = 0
		While Not rsCat.EOF
	%>
	 	<tr>
		 
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLECODE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLENAME")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("vehiclegroupcode")%> : <%=rscat("description")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("EMPLOYEECODE")%> : <%=rscat("EMPLOYEENAME")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HELPER1")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HELPER2")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLEDATE")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TIME")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JOBCODE")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>		
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONTYPE")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONcode")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("categorytransport")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("type_kontanan")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("chargeunit")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("chargejob")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("chargetype")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("chargecode")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(To_Nol(rscat("trip")),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(To_Nol(rscat("Hours")),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(To_Nol(rscat("UNIT")),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(To_Nol(rscat("Output")),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(To_Nol(rscat("Amount")),2)%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ref")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
		</tr>
	  <%
	  v_HOURS= v_HOURS+trim(cdbl(To_Nol(rsCat("Hours"))))
	  v_UNIT = v_UNIT+Trim(cdbl(To_Nol(rsCat("UNIT"))))
	  v_OUTPUT = v_OUTPUT+Trim(cdbl(To_Nol(rsCat("Output"))))
	  v_AMOUNT = v_AMOUNT+trim(cdbl(To_Nol(rsCat("Amount"))))
	  v_trip = v_trip+trim(cdbl(To_Nol(rsCat("trip"))))
	  rsCat.MoveNext
	  response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="18"><b><font face="Verdana" size="1">T o t a l</b></td>
	<td align="Right" style="<%=cmdstyle%>"><b><font face="Verdana" size="1"><%=FormatNumber(v_trip,2)%></b></td>
	<td align="Right" style="<%=cmdstyle%>"><b><font face="Verdana" size="1"><%=FormatNumber(v_HOURS,2)%></b></td>
	<td align="Right" style="<%=cmdstyle%>"><b><font face="Verdana" size="1"><%=FormatNumber(v_UNIT,2)%></b></td>
	<td align="Right" style="<%=cmdstyle%>"><b><font face="Verdana" size="1"><%=FormatNumber(v_OUTPUT,2)%></b></td>
	<td align="Right" style="<%=cmdstyle%>"><b><font face="Verdana" size="1"><%=FormatNumber(v_AMOUNT,2)%></b></td>
	<td align="Right" style="<%=cmdstyle%>" colspan="2"><b><font face="Verdana" size="1"></b></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>