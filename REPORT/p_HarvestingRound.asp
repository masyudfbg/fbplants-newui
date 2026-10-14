<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Harvesting Round</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtHRNDYEAR= trim(request("HRNDYEAR"))
v_txtHRNDMONTH= trim(request("HRNDMONTH"))
v_txtFIELDCODE= trim(request("FIELDCODE"))
v_txtCROPCODE= trim(request("CROPCODE"))


txtSQL = "SELECT HRNDYEAR, HRNDMONTH, FIELDCODE, CROPCODE, NUMBEROFROUNDS, HECTARAGE, STANDPERHECTHARVESTED FROM HarvestingRound" & _
		 " where fcba='" & session("fcBA") & "'"

if  len(trim(v_txtHRNDYEAR)) > 0 then
    txtSQL = txtSQL & " and HRNDYEAR='" & v_txtHRNDYEAR & "'"
	txtSelect = txtSelect & " Year " & v_txtHRNDYEAR
end if
if  len(trim(v_txtHRNDMONTH)) > 0 then
    txtSQL =txtSQL &  " and HRNDMONTH='" & v_txtHRNDMONTH & "'"
	txtSelect = txtSelect & " Month " & v_txtHRNDMONTH
end if 
if  len(trim(v_txtFIELDCODE)) > 0 then
    txtSQL =txtSQL &  " and FIELDCODE='" & v_txtFIELDCODE & "'"
	txtSelect = txtSelect & " FieldCode " & v_txtFIELDCODE
end if 
if  len(trim(v_txtCROPCODE)) > 0 then
    txtSQL =txtSQL &  " and CROPCODE='" & v_txtCROPCODE & "'"
	txtSelect = txtSelect & " CropCode " & v_txtCROPCODE
end if 

txtsql=(txtsql) 
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
			<div id="header-wrapper">
				<div id="header">




					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>HARVESTING ROUNDS</h2>
				</div>
			</div>
			<table border="0" width="100%" id="product-table" class="headertable">
				<div class="printhide">







				<tr>
					<td colspan="7"><%=txtSelect%></td>
				</tr>
				</div>
				<%	
				Set rsCat = connStr.Execute(txtSQL)
					if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					
					<th>Year</th>
					<th>Month</th>
					<th>Field Code</th>
					<th>Crop Code</th>
					<th>Number Of Rounds</th>
					<th>Hectarage</th>
					<th>Stand PerHectarage Harvested</th>
					
				</tr>
				<%

				Dim v_Quantity
				v_Quantity=0
				While Not rsCat.EOF%>


		<tr>
		
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HRNDYEAR")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HRNDMONTH")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FIELDCODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CROPCODE")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NUMBEROFROUNDS")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HECTARAGE")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STANDPERHECTHARVESTED")%>&nbsp;</td>
		</tr>
	  <%
	  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">	
</table>
<!--#include file="Bottom_Report.asp"-->
</div>
	</body>
</html>