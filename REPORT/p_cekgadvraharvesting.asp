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

	If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
	If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	
	v_txtFcba = Request("txtFcba")

	txttdate = txtMonth & "" & txtYear
	v_txtPOCode		=	trim(request("txtPOCode"))
	compare1 = trim(request("compare1"))

	varDefaultBa = Session("fcBA")

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
		
       if  v_txtfcba <> "" then
	       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
		end if
	    ' selectCons = selectCons & ",fcba"
	    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
		   fcbaname = rsCat("fcName")
		   CompanyName=rsCat("FCCOMPANYNAME")
		   MgrName = rsCat("FCMANAGER")
		   KTUName = rsCat("FCKTU")
		   CropName=rsCat("FCCROP")
		   FcbaType=rsCat("fcType")
	else
	    nSpanTotal = nSpanTotal - 1 
	end if
	
	strGrnInv		=	"   SELECT vehicleactivity.fcba," & _
						"          vehicleactivity.vehicledate," & _
						"          vehicleactivity.vehiclecode," & _
						"          attendance_gad.gad_employeecode," & _
						"          attendance_gad.gad_employeename," & _
						"          vehicleactivity.vh_employeecode," & _
						"          vehicleactivity.vh_employeename," & _
						"          harvestingspb.harvesting_employeecode," & _
						"          harvestingspb.harvesting_employeename" & _
						"     FROM (SELECT distinct iplasprod.attendance_gad.fcba," & _
						"                  iplasprod.attendance_gad.fddate," & _
						"                  iplasprod.attendance_gad.employeecode AS gad_employeecode," & _
						"                  iplasprod.employee.fcname AS gad_employeename," & _
						"                  iplasprod.attendance_gad.locationcode" & _
						"             FROM iplasprod.attendance_gad" & _
						"                  INNER JOIN iplasprod.employee" & _
						"                     ON (    iplasprod.attendance_gad.fcba =" & _
						"                                iplasprod.employee.fcba" & _
						"                         AND iplasprod.attendance_gad.employeecode =" & _
						"                                iplasprod.employee.fccode)" & _
						"            WHERE iplasprod.attendance_gad.locationtype = 'VH') attendance_gad" & _
						"          INNER JOIN" & _
						"          (SELECT distinct iplasprod.vehicleactivity.fcba," & _
						"                  iplasprod.vehicleactivity.vehicledate," & _
						"                  iplasprod.vehicleactivity.employeecode AS vh_employeecode," & _
						"                  iplasprod.employee.fcname AS vh_employeename," & _
						"                  iplasprod.vehicleactivity.vehiclecode" & _
						"             FROM iplasprod.vehicleactivity" & _
						"                  INNER JOIN iplasprod.employee" & _
						"                     ON (    iplasprod.vehicleactivity.fcba =" & _
						"                                iplasprod.employee.fcba" & _
						"                         AND iplasprod.vehicleactivity.employeecode =" & _
						"                                iplasprod.employee.fccode)) vehicleactivity" & _
						"             ON (    attendance_gad.fcba = vehicleactivity.fcba" & _
						"                 AND attendance_gad.locationcode = vehicleactivity.vehiclecode" & _
						"                 AND attendance_gad.fddate = vehicleactivity.vehicledate)" & _
						"          INNER JOIN" & _
						"          (SELECT distinct iplasprod.harvestingspb.fcba," & _
						"                  iplasprod.harvestingspb.receptiondate," & _
						"                  iplasprod.harvestingspb.driver AS harvesting_employeecode," & _
						"                  iplasprod.employee.fcname AS harvesting_employeename," & _
						"                  iplasprod.harvestingspb.vehicle" & _
						"             FROM iplasprod.harvestingspb" & _
						"                  INNER JOIN iplasprod.employee" & _
						"                     ON (    iplasprod.harvestingspb.fcba =" & _
						"                                iplasprod.employee.fcba" & _
						"                         AND iplasprod.harvestingspb.driver =" & _
						"                                iplasprod.employee.fccode)) harvestingspb" & _
						"             ON (    vehicleactivity.fcba = harvestingspb.fcba" & _
						"                 AND vehicleactivity.vehiclecode = harvestingspb.vehicle" & _
						"                 AND vehicleactivity.vehicledate = harvestingspb.receptiondate)" & _
						"    WHERE     vehicleactivity.fcba = '"& v_txtFcba &"'" & _
						"          AND TO_NUMBER(TO_CHAR (vehicleactivity.vehicledate, 'mm')) = '"& txtMonth &"'" & _
						"          AND TO_NUMBER(TO_CHAR (vehicleactivity.vehicledate, 'yyyy')) = '"& txtYear &"'" 


	if  cdbl(to_nol(compare1)) = 1 then
	    strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != vehicleactivity.vh_employeecode"
	end if
	if  cdbl(to_nol(compare1)) = 2 then
	    strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != harvestingspb.harvesting_employeecode" 
	end if
	if  cdbl(to_nol(compare1)) = 3 then
	    strGrnInv	=	strGrnInv & " and vehicleactivity.vh_employeecode != harvestingspb.harvesting_employeecode" 
	end if

	' strGrnInv	=	strGrnInv &	" GROUP BY" & _
	' 							"     FCBA" & _
	' 							"     , NOPO" & _
	' 							"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv & " ORDER BY vehicleactivity.fcba," & _
								"          vehicleactivity.vehicledate," & _
								"          vehicleactivity.vehiclecode," & _
								"          attendance_gad.gad_employeecode" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Cek GAD vs VRA vs Harvesting</title>
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
							<font face="verdana" size="1">KEBUN</font>
						</td>
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
										" where " & _
										" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
							  
							
							rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
							if not rsFcba.eof then
								DO while not rsFcba.eof
									if v_txtFcba = rsFcba("fcba") then
									%>			
										<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
									<% Else %>	
										<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
									<% End If
									
									rsFcba.movenext
								loop
							End if
							rsFcba.Close
							set rsFcba = Nothing
							%>
							</select>	
						</td>
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
						<td align="left">
							<font face="verdana" size="1">Compare</font>
						</td>	
						<td align="left">
							<select name="compare1" value="<%=compare1%>" >
								<option value=""></option>
								<option value="1" <% If cstr(""&compare1) = "1" Then %> Selected <%End if%>>Attendance GAD Employee <> VRA Employee</option> 
								<option value="2" <% If cstr(""&compare1) = "2" Then %> Selected <%End if%>>Attendance GAD Employee <> Harvesting Employee</option> 
								<option value="3" <% If cstr(""&compare1) = "3" Then %> Selected <%End if%>>VRA Employee <> Harvesting Employee</option> 
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
				<h4>PT. <%=CompanyName%></h4>
				<h5><%=fcbaname%></h5>
			</div>
			
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>Crosscek GAD vs VRA vs Harvesting</h2>
                <br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
		</div>
</th>
		</tr>
	<tr>
			<th style="background-color:#00A300;color:#fff;">Vehicle Date</th>
			<th style="background-color:#00A300;color:#fff;">Vehicle Code</th>
			<th style="background-color:#00A300;color:#fff;">Attendance GAD Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">Attendance GAD Employee Name</th>
			<th style="background-color:#00A300;color:#fff;">VRA Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">VRA Employee Name</th>
			<th style="background-color:#00A300;color:#fff;">Harvesting Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">Harvesting Employee Name</th>
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
			<td align="center"><%=rsGrnInv("vehicledate")%></td>
			<td align="center"><%=rsGrnInv("vehiclecode")%></td>
			<td align="center"><%=rsGrnInv("gad_employeecode")%></td>
			<td align="center"><%=rsGrnInv("gad_employeename")%></td>
			<td align="center"><%=rsGrnInv("vh_employeecode")%></td>
			<td align="center"><%=rsGrnInv("vh_employeename")%></td>
			<td align="center"><%=rsGrnInv("harvesting_employeecode")%></td>
			<td align="center"><%=rsGrnInv("harvesting_employeename")%></td>
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
