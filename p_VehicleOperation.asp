<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Vehicle Operation</title>
</head>
<body>
<%
v_VEHICLEDATE_From=trim(request("VEHDATE_From"))
v_VEHICLEDATE_To=trim(request("VEHDATE_to"))
v_JOBCODE=trim(request("JOBCODE"))
v_VEHICLEGROUPCODE=trim(request("VEHICLEGROUPCODE"))
v_VEHICLECODE=trim(request("VEHICLECODE"))
vehiclestatus=request("vehiclestatus")

	txtSQL = " SELECT vh.fcba" & _
	         "        , vh.vehicledate" & _
	         "        , vh.vehiclecode" & _
	         "        , v.fcname vehiclename" & _
	         "        , vh.vehiclestatus" & _
	         "        , pr.fcname vehiclestatusname" & _
	         "        , vh.unit" & _
	         "        , vh.fuel" & _
	         "        , vh.harihujan" & _
	         "        , vh.curahhujan" & _
	         "        , vh.remarks" & _
             " FROM vehicleoperation vh" & _
			 " INNER JOIN vehicle v ON (v.fcba = vh.fcba AND v.fccode = vh.vehiclecode)" & _
			 " LEFT JOIN parameter pr ON (pr.fcba = vh.fcba AND pr.fccode = vh.vehiclestatus AND pr.parcode = 'VO001')" & _			 
             " WHERE vh.VEHICLEDATE >= TO_DATE('" & v_VEHICLEDATE_From & "','MM/DD/YYYY')" & _
             "       AND vh.VEHICLEDATE <= TO_DATE('" & v_VEHICLEDATE_To & "','MM/DD/YYYY')" & _
			 "       AND vh.fcba = '" & session("fcBA") & "'"
	 
if  len(trim(v_VEHICLEDATE_From)) > 0 then
	txtSelect = " From " & FormatDateTime(v_VEHICLEDATE_From,1)
end if
if  len(trim(v_VEHICLEDATE_To)) > 0 then
	txtSelect = txtSelect & "  to " & FormatDateTime(v_VEHICLEDATE_To,1)
end if
if  len(trim(v_VEHICLEGROUPCODE)) > 0 then
    txtSQL =txtSQL &  " AND v.VEHICLEGROUPCODE = '" & v_VEHICLEGROUPCODE & "'"
	txtSelect = txtSelect & " Vehicle Group " & v_VEHICLEGROUPCODE
end if 
if  len(trim(v_VEHICLECODE)) > 0 then
    txtSQL =txtSQL &  " AND vh.VEHICLECODE = '" & v_VEHICLECODE & "'"
	txtSelect = txtSelect & " Vehicle Code " & v_VEHICLECODE
end if 
if  "" & vehiclestatus <> "" then
    txtSQL = txtSQL &  " AND vh.vehiclestatus = '" & vehiclestatus & "'"
end if 

txtsql=(txtsql) & " ORDER BY vh.vehicledate, vh.vehiclecode "
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write txtsql
Set rsCat = connStr.Execute(txtSQL)
%>
	<body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>Vehicle Operation</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="9"><%=txtSelect%></td>
						</tr>
						</div>
	</tr>
	<tr>
	    <th>Date</th>
		<th>Code</th>
		<th>Description</th>
		<th>Status</th>
		<!--
		<th>Hm/Km/Jam</th>
		<th>Fuel (Ltr)</th>
		<th>Hari Hujan (Hari)</th>
		<th>Curah Hujan (mm)</th>-->
		<th>Remarks</th>
	</tr>
	<%While Not rsCat.EOF%>
	 	<tr>
		 
		<td align="Left"><font face="Verdana" size="1"><%=rscat("vehicledate")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("vehiclecode")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("vehiclename")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("vehiclestatusname")%></td>
		<!--<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("unit"))%></td>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("fuel"))%></td>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("harihujan"))%></td>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("curahhujan"))%></td>-->
		<td align="Left"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>