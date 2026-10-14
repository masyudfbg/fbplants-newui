<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting Reporting - Harvesting Write Off</title>
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

v_txtWRITEOFFDATE_From=trim(request("WRITEOFFDATE_From"))
v_txtWRITEOFFDATE_To=trim(request("WRITEOFFDATE_To"))
v_txtWRITEOFFCODE=trim(request("WRITEOFFCODE"))



txtSQL = " SELECT a.WRITEOFFDATE, " & _
				   " a.WRITEOFFCODE, " & _
				   " b.FIELDCODE, " & _
				   " a.Reference, " & _
				   " a.remarks, " & _
				   " nvl(b.QUANTITY,0) QUANTITY" & _
			  " FROM harvestingwriteoff a, harvestingwriteoffdetail b " & _
			 " WHERE a.WRITEOFFCODE = b.WRITEOFFCODE AND a.fcba = b.fcba " & _
			 "       and b.fcba = '" & Session("Fcba") & "'"

if  len(trim(v_txtWRITEOFFDATE_From)) > 0 then
    txtSQL = txtSQL & " and a.WRITEOFFDATE >=to_Date('" & v_txtWRITEOFFDATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtWRITEOFFDATE_From,1)	
end if
if  len(trim(v_txtWRITEOFFDATE_To)) > 0 then
    txtSQL =txtSQL &  " and a.WRITEOFFDATE <=to_Date('" & v_txtWRITEOFFDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtWRITEOFFDATE_To,1)	
end if
if  len(trim(v_txtWRITEOFFCODE)) > 0 then
    txtSQL = txtSQL & " and a.WRITEOFFCODE='" & v_txtWRITEOFFCODE & "'"
	txtSelect = txtSelect & " Write Off Code " & v_txtWRITEOFFCODE	
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
			<!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>HARVESTING WRITE OFF</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>								 
							<th>Date</th>
							<th>WriteOff Code</th>
							<th>Field Code</th>
							<th>Reference</th>
							<th>Remarks</th>
							<th>Quantity</th>
						</tr>
						</thead>
                        <tbody>
<%
									Set rsCat = Server.CreateObject("ADODB.Recordset")
									rsCat.CursorType = 3
									rsCat.CursorLocation = 3
									rsCat.Open txtSQL, connstr
%>


						<%If rsCat.BOF and rsCat.EOF Then%>
							<tr>
								<td align="left" colspan="6">NO DATA AVAILABLE</td>
							</tr>
						<%End If%>

						<%If Not rsCat.EOF Then
							nCountofField = rsCat.fields.count-1
							NO = 1
							Dim v_Quantity : v_Quantity = 0
							
							While Not rsCat.EOF
						%>
							<tr>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("writeoffdate")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("writeoffcode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FieldCode")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Reference")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Remarks")%> </td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("Quantity")),0)%> </td>
							</tr>
							<%
									v_Quantity = v_Quantity + Cdbl(rscat("Quantity"))
									rsCat.MoveNext
									'response.Flush()
									no = no + 1
								Wend
							End If
								rsCat.close()
								rsCat = null
							%>
								<input  type="hidden" value="<%=txtTable%>" name="txtTable">
								<input  type="hidden" value="<%=txtModule%>" name="txtModule">
								</form>
							<tr bgcolor="#F4F4A0">
								<td align="Center" style="<%=cmdstyleLeft%>" colspan="5"> <font face="Verdana" size="1"><strong>T o t a l</strong></td>
								<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Quantity,0)%></strong></td>
							</tr>
							</tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</body>	
</HTML>