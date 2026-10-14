<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Harvesting</title>
</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtCHITNO= trim(request("CHITNO"))
v_txtFIELDCODE= trim(request("FIELDCODE"))
v_txtCROPCODE= trim(request("CROPCODE"))
v_txtRECEPTIONDATE_From= trim(request("RECEPTIONDATE_From"))
v_txtRECEPTIONDATE_To= trim(request("RECEPTIONDATE_To"))
v_txtPRODUCTCODE=trim(request("PRODUCTCODE"))

txtSQL = "SELECT CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, RECEPTIONDATE, " & _
                     " VEHICLE, OWN, nvl(sum(QUANTITY),0) as Quantity, ESTATE_WEIGHT, MILL_WEIGHT FROM Harvesting" & _
					 " where fcba='" & session("fcBA") & "'"
if  len(trim(v_txtRECEPTIONDATE_From)) > 0 then
    txtSQL = txtSQL & " and RECEPTIONDATE >=to_Date('" & v_txtRECEPTIONDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtRECEPTIONDATE_From,1)
end if
if  len(trim(v_txtRECEPTIONDATE_To)) > 0 then
    txtSQL =txtSQL &  " and RECEPTIONDATE <=to_Date('" & v_txtRECEPTIONDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtRECEPTIONDATE_To,1)
end if
if  len(trim(v_txtCHITNO)) > 0 then
    txtSQL = txtSQL & " and CHITNO='" & v_txtCHITNO & "'"
	txtSelect = txtSelect & " Chit No. " & v_txtCHITNO
end if
if  len(trim(v_txtFIELDCODE)) > 0 then
    txtSQL =txtSQL &  " and FIELDCODE='" & v_txtFIELDCODE & "'"
	txtSelect = txtSelect & " Field Code " & v_txtFIELDCODE
end if 
if  len(trim(v_txtCROPCODE)) > 0 then
    txtSQL =txtSQL &  " and CROPCODE='" & v_txtCROPCODE & "'"
	txtSelect = txtSelect & " Crop Code " & v_txtCROPCODE
end if 
if  len(trim(v_txtPRODUCTCODE)) > 0 then
    txtSQL =txtSQL &  " and PRODUCTCODE='" & v_txtPRODUCTCODE & "'"
	txtSelect = txtSelect & " Product Code " & v_txtPRODUCTCODE
end if 

txtsql=(txtsql) & " group by CHITNO, FIELDCODE, CROPCODE, PRODUCTCODE, RECEPTIONDATE, " & _
                     " VEHICLE, OWN, ESTATE_WEIGHT, MILL_WEIGHT"
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
					<h2>HARVESTING</h2>
				</div>
			</div>
			<table border="0" width="100%" id="product-table" class="headertable">
				<div class="printhide">







				<tr>
					<td colspan="11"><%=txtSelect%></td>
				</tr>
				</div>
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>No</th>
					<th>Chit No</th>
					<th>Field Code</th>
					<th>Crop Code</th>
					<th>Product Code</th>
					<th>Reception Date</th>
					<th>Vehicle</th>
					<th>Own</th>
					<th>Estate Weight</th>
					<th>Mill Weight</th>
					<th>Quantity</th>
				</tr>
				<%

				Dim v_Quantity, nEstateWeight, nMillweight
				v_Quantity=0
				nEstateWeight=0
				nMillweight=0
				While Not rsCat.EOF%>

				<tr>
					<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHITNO")%></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FIELDCODE")%></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CROPCODE")%></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCTCODE")%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RECEPTIONDATE")%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLE")%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("OWN")%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ESTATE_WEIGHT")%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MILL_WEIGHT")%></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Quantity")%></td>
				</tr>
				<%
				v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
				nEstateWeight=nEstateWeight+trim(cdbl(rsCat("ESTATE_WEIGHT")))
				nMillweight=nMillweight+trim(cdbl(rsCat("MILL_WEIGHT")))
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

					<td align="Center" colspan="8"> <strong><i><font face="Verdana" size="1">T o t a l</i></strong></td>
					<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=nEstateWeight%></i></strong></td>
					<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=nMillweight%></i></strong></td>
					<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=v_Quantity%></i></strong></td>
				</tr>
			</table>
<!--#include file="Bottom_Report.asp"-->
		</div>
	</body>
</html>