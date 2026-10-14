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
<title>ePlaS Reporting - Workshop Activity</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_WORKSHOPDATE_From=trim(request("DATEACTIVITY_From")) 
v_WORKSHOPDATE_To=trim(request("DATEACTIVITY_To"))
v_JOB=trim(request("Job"))
v_LOCATIONTYPE=trim(request("LOCATIONTYPE"))
v_LOCATIONCODE=trim(request("LOCATIONCODE"))
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

Dim strAmount,strBasis
StrAmount  = "nvl(WORKSHOPCB_API.Get_Amount(ws.WORKSHOPCODE, ws.JOB,ws.WORKSHOPDATE, ws.LOCATIONTYPE,ws.LOCATIONcode,ws.fcba),0)"
StrBasis  =  "nvl(WORKSHOPCB_API.Get_Basis(ws.WORKSHOPCODE, ws.JOB,ws.WORKSHOPDATE, ws.LOCATIONTYPE,ws.LOCATIONcode,ws.fcba),0)"

if request("txtKind") = "WorkshopRunningAccount" Then
	txtSQL = "select ws.WORKSHOPCODE, ws.WORKSHOPDATE, ws.employeecode, employee_api.get_employeename(ws.fcba, ws.employeecode) employeename, ws.TIME, ws.LOCATIONTYPE, ws.LOCATIONCODE,ws.JOB, j.fcname,ws.ref, ws.remarks," & _
     " nvl(sum(ws.HOURS),0) as Hours,nvl(sum(ws.UNIT),0) as Unit,nvl(sum(ws.OUTPUT),0) as Output," & StrAmount & " as Amount," & StrBasis & " as basis" & _
     " from WorkshopActivity ws,job j " & _
     " where (ws.job = j.fccode and ws.fcba = j.fcba) " & _
	 " AND WS.fcba='" & businessunit & "'" & _
	 " AND j.fcba='" & businessunit & "'" 
Else
	txtSQL = "select ws.WORKSHOPCODE, ws.WORKSHOPDATE, ws.employeecode, employee_api.get_employeename(ws.fcba, ws.employeecode) employeename, ws.TIME, ws.LOCATIONTYPE, ws.LOCATIONCODE,ws.JOB, j.fcname, ws.ref, ws.remarks," & _
     " nvl(sum(ws.HOURS),0) as Hours,nvl(sum(ws.UNIT),0) as Unit,nvl(sum(ws.OUTPUT),0) as Output," & StrAmount & " as Amount," & StrBasis & " as basis" & _
     " from WorkshopActivity ws,job j " & _
     " where (ws.job = j.fccode and ws.fcba = j.fcba) " & _
	 " AND WS.fcba='" & businessunit & "'"	
End If


if  len(trim(v_WORKSHOPDATE_From)) > 0 then
    txtSQL = txtSQL & " and Ws.WORKSHOPDATE >=to_Date('" & v_WORKSHOPDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_WORKSHOPDATE_From,1)
end if
if  len(trim(v_WORKSHOPDATE_To)) > 0 then
    txtSQL =txtSQL &  " and Ws.WORKSHOPDATE <=to_Date('" & v_WORKSHOPDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_WORKSHOPDATE_To,1)
end if
if  len(trim(v_JOB)) > 0 then
    txtSQL = txtSQL & " and ws.JOB='" & v_JOB & "'"
	txtSelect = txtSelect & " Job Code " & v_JOB
end if
if  len(trim(v_LOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and ws.LOCATIONTYPE='" & v_LOCATIONTYPE & "'"
	txtSelect = txtSelect & " Target Type " & v_LOCATIONTYPE
end if 
if  len(trim(v_LOCATIONCODE)) > 0 then
    txtSQL =txtSQL &  " and ws.LOCATIONCODE='" & v_LOCATIONCODE & "'"
	txtSelect = txtSelect & " Target Code " & v_LOCATIONCODE
end if 
txtsql=(txtsql) & " group BY ws.fcba, ws.WORKSHOPCODE,ws.employeecode, employee_api.get_employeename(ws.fcba, ws.employeecode), ws.WORKSHOPDATE, ws.TIME, ws.LOCATIONTYPE, ws.LOCATIONCODE,ws.JOB, j.fcname,ws.time,ws.ref, ws.remarks order by ws.WORKSHOPDATE, ws.TIME"
'response.Write(txtsql)
response.Flush()
%>
    <body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>WORKSHOP ACTIVITY</h2>
				</div>
			</div>
		</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
						</tr>
						<!--</div>-->
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1
	Dim v_Quantity,v_qtyretrun,v_amountreceived,v_amountreturn
	v_HOURS=0
	v_UNIT=0
	v_OUTPUT=0
	v_AMOUNT=0
	%>
	</tr>
	<tr>
	    <th>Code</th>
		<th>Employee Code</th>
		<th>Employee Name</th>
		<th>Date</th>
		<th>Time</th>
		<th>Job Code</th>
		<th>Job Description</th>
		<th>Target Type</th>
		<th>Target Code</th>
		<th>Hours</th>
		<th>Unit</th>
		<th>Output</th>
		<th>Amount</th>
		<th>Reference</th>
		<th>Remarks</th>		
	</tr>
    </thead>
    <tbody>
	<%While Not rsCat.EOF%>
	 	<tr>
		 
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WORKSHOPCODE")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Employeecode")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("employeename")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WORKSHOPDATE")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TIME")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JOB")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>		
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONTYPE")%></td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONCODE")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Hours"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNIT"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Output"),2)%></td>
		<% if cdbl(rscat("Basis")) > 0 then %>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber((cdbl(rscat("UNIT"))/cdbl(rscat("Basis"))) * cdbl(rscat("Amount")),2)%></td>
		<% else %>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1">0</td>
		<% End If %>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ref")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
		</tr>
	  <%
	  v_HOURS=v_HOURS+trim(cdbl(rsCat("Hours")))
	  v_UNIT  =v_UNIT+Trim(cdbl(rsCat("UNIT")))
	  v_OUTPUT =v_OUTPUT+Trim(cdbl(rsCat("Output")))
	  if cdbl(rscat("Basis")) > 0 then
	  v_AMOUNT  =v_AMOUNT+ (cdbl(rscat("UNIT"))/cdbl(rscat("Basis"))) * cdbl(rscat("Amount"))
	  end if
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
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="9"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_HOURS,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_UNIT,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_OUTPUT,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_AMOUNT,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
	
	</tr>
    </tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>