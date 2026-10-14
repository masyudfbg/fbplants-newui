<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Harvesting Write Off</title>
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

v_txtHARVESINGVALUEDATE_From=trim(request("HARVESINGVALUEDATE_From"))
v_txtHARVESINGVALUEDATE_tO=trim(request("HARVESINGVALUEDATE_tO"))
v_txtHARVESTINGVALUECODE=trim(request("HARVESTINGVALUECODE"))



txtSQL = " SELECT a.HARVESTINGVALUECODE," & _
				" a.FDDATE," & _
				" b.LINENO, " & _
				" b.DIVISIONINFO, " & _
				" NVL(b.SAMPLE_QUANTITY,0) BUNCH," & _
				" NVL(b.SAMPLE_WEIGHT,0) WEIGHT," & _
				" NVL(b.UNDER_RIPE,0) BUAHMENTAH," & _
				" NVL(b.FINE_RIPE,0) BUAHMASAK," & _
				" NVL(b.OVER_RIPE,0) BUAHTERLALUMATANG," & _
				" NVL(b.LONG_STALK,0) TANGKAIPANJANG," & _
				" NVL(b.EMPTY_BUNCH,0) BUAHKOSONG," & _
				" NVL(b.ABNORMAL_RIPE,0) BUAHABNORMAL," & _
				" NVL(b.EATEN_BY_RAT,0) BUAHDIMAKANTIKUS," & _
				" NVL(b.SPOILED_RIPE,0) BUAHBUSUK," & _
				" NVL(b.DETACHED_RIPE,0) BRONDOLAN," & _
				" NVL(b.RUBBISH,0) SAMPAH," & _
				" NVL(b.QE_1,0) QE_1," & _
				" NVL(b.QE_2,0) QE_2," & _
				" NVL(b.QE_3,0) QE_3," & _
				" NVL(b.QE_4,0) QE_4," & _
				" NVL(b.QE_5,0) QE_5," & _
				" NVL(b.QE_6,0) QE_6," & _
				" NVL(b.QE_7,0) QE_7," & _
				" NVL(b.QE_8,0) QE_8," & _
				" NVL(b.QE_9,0) QE_9," & _
				" NVL(b.QE_10,0) QE_10" & _
				" FROM HARVESTINGVALUE a, HARVESTINGVALUEDETAILS b" & _
				" WHERE a.HARVESTINGVALUECODE= b.HARVESTINGVALUECODE and A.FCBA=B.FCBA"
if  len(trim(v_txtHARVESINGVALUEDATE_From)) > 0 then
    txtSQL = txtSQL & " and a.FDDATE >=to_Date('" & v_txtHARVESINGVALUEDATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtHARVESINGVALUEDATE_From,1)	
end if
if  len(trim(v_txtHARVESINGVALUEDATE_tO)) > 0 then
    txtSQL =txtSQL &  " and a.FDDATE <=to_Date('" & v_txtHARVESINGVALUEDATE_tO & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtHARVESINGVALUEDATE_tO,1)	
end if
if  len(trim(v_txtHARVESTINGVALUECODE)) > 0 then
    txtSQL = txtSQL & " and a.HARVESTINGVALUECODE='" & v_txtHARVESTINGVALUECODE & "'"
	txtSelect = txtSelect & " Hrvesting Value Code " & v_txtHARVESTINGVALUECODE	
end if


if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>HARVESTING VALUE</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="50"><%=txtSelect%></td>
						</tr>
						</div>
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
								 
								<th>Date</th>
								<th>Harvesing Value Code</th>
								<th>Line No</th>
								<th>Division</th>
								<th>BUNCH</th>
								<th>WEIGHT</th>
								<th>BUAH MENTAH</th>
								<th>BUAH MASAK</th>
								<th>BUAH TERLALU MATANG</th>
								<th>TANGKAI PANJANG</th>
								<th>BUAH KOSONG</th>
								<th>BUAH ABNORMAL</th>
								<th>BUAH DIMAKAN TIKUS</th>
								<th>BUAH BUSUK</th>
								<th>BRONDOLAN</th>
								<th>SAMPAH</th>
								<th>QE_1</th>
								<th>QE_2</th>
								<th>QE_3</th>
								<th>QE_4</th>
								<th>QE_5</th>
								<th>QE_6</th>
								<th>QE_7</th>
								<th>QE_8</th>
								<th>QE_9</th>
								<th>QE_10</th>
								
							</tr>
								<%
								Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
								v_Quantity=0
								v_Amount = 0
								While Not rsCat.EOF%>


									<tr>
									<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FDDATE")%> </td>
									 <td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("HARVESTINGVALUECODE")%> </td>
									<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LINENO")%> </td>
									<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DIVISIONINFO")%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("BUNCH")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("WEIGHT")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("BUAHMENTAH")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("BUAHMASAK")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("BUAHTERLALUMATANG")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("TANGKAIPANJANG")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("BUAHKOSONG")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("BUAHABNORMAL")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("BUAHDIMAKANTIKUS")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("BUAHBUSUK")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("BRONDOLAN")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("SAMPAH")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_1")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_2")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_3")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_4")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_5")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_6")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_7")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_8")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_9")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("QE_10")),2)%> </td>

									</tr>
								  <%

								  rsCat.MoveNext
								  'response.Flush()
								  no=no+1
								Wend
								rsCat.close()
								rsCat=null%>
								<input  type="hidden" value="<%=txtTable%>" name="txtTable">
								<input  type="hidden" value="<%=txtModule%>" name="txtModule">
								</form>
								
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</body>	
</HTML>