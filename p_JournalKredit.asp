<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>Journal Kredit - iPlaS</title>
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
v_txtDATECREATED_From = request("DATECREATED_From")
v_txtDATECREATED_To = request("DATECREATED_To")
v_txtBANKCODE = request("BANKCODE")
v_txtVOUCHERCODE =  request("VOUCHERCODE")
v_txtJOBCODE = REQUEST("JOBCODE")
v_txtLOCATIONTYPE = request("LOCATIONTYPE")
v_txtITEMCODE = request("ITEMCODE")

txtSQL = " select p.datecreated" & _
         "        , pay.vouchercode" & _
		 "        , pay.lineno" & _
         "        , pay.jobcode" & _
         "        , job.fcname" & _
         "        , pay.locationtype" & _
         "        , pay.locationcode" & _
         "        , pay.amount" & _
         "        , pay.reference" & _
         "        , pay.remarks" & _
		 " from paymentvoucherdetail pay" & _
         "      , job job" & _
         "      , paymentvoucher p" & _
		 " where ( ((pay.jobcode = job.fccode) and (pay.fcba = job.fcba))" & _
         "       and ((pay.vouchercode = p.vouchercode) and (pay.fcba = p.fcba)) )" & _
		 "       and pay.fcba = '" & session("fcBA") & "'"
if  len(trim(v_txtDATECREATED_From)) > 0 then
    txtSQL = txtSQL & " and p.datecreated >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
    txtSQL =txtSQL &  " and p.datecreated <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtBANKCODE)) > 0 then
    txtSQL = txtSQL & " and p.BANKCODE='" & v_txtBANKCODE & "'"
	txtSelect = txtSelect & " Bank Code " & v_txtBANKCODE
end if
if  len(trim(v_txtVOUCHERCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.VOUCHERCODE='" & v_txtVOUCHERCODE & "'"
	txtSelect = txtSelect & " Voucher No " & v_txtVOUCHERCODE
end if 
if  len(trim(v_txtJOBCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.jobcode='" & v_txtJOBCODE & "'"
	txtSelect = txtSelect & "  Job Code " & v_txtJOBCODE
end if 
if  len(trim(v_txtLOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and pay.locationtype='" & v_txtLOCATIONTYPE & "'"
	txtSelect = txtSelect & " Location Type " & v_txtLOCATIONTYPE
end if 
if  len(trim(v_txtITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.locationcode='" & v_txtITEMCODE & "'"
	txtSelect = txtSelect & " Location Code " & v_txtITEMCODE
end if 

txtsql	=	txtsql & " ORDER BY p.datecreated, pay.vouchercode, lineno asc, amount Desc"
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
         <th colspan="10" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
               <div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
                <div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h1 class="judul">Journal Kredit</h1>
				</div>
				<div style="clear:both;"></div>
			</div>
            </th>
			</tr>
				<!--<div class="printhide">-->
				<tr>
					<th colspan="10" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
				<!--</div>-->
				<tr>
					 
					<th>Date</th>
					<th>Code</th>
					<th>Line No</th>
					<th>Activity Code</th>
					<th>Activity Name</th>
					<th>Target Type</th>
					<th>Target Code</th>
					<th>Reference</th>
					<th>Remarks</th>
					<th>Amount</th>
				</tr>
                </thead>
                <tbody>
				<%
				Dim v_Opening, v_Debit, v_Credit, v_Saldo	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO	=	1
				v_Amount	=	0
				While Not rsCat.EOF
				%>
				<tr>
					
					<td align="Center"><%=rscat("Datecreated")%></td>
					<td align="left"><a href="javascript:launchwin('PrintVoucher.Asp?fcCode=<%=rscat("Vouchercode")%>&Mode=Print','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><%=rscat("Vouchercode")%></a></td>
					<td align="Right"><%=rscat("lineno")%></td>
					<td align="left"><%=rscat("Jobcode")%></td>
					<td align="left"><%=rscat("Fcname")%></td>
					<td align="left"><%=rscat("Locationtype")%></td>
					<td align="left"><%=rscat("Locationcode")%></td>
					<td align="left"><%=rscat("Reference")%></td>
					<td align="left"><%=rscat("Remarks")%></td>
					<td align="Right"><%=FormatNumber(rscat("Amount"),2)%></td>
				</tr>
				<%
				v_Amount	=	v_Amount	+	trim(cdbl(rsCat("Amount")))
				'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
				rsCat.MoveNext
				'response.Flush()
				no	=	no	+	1
				Wend
				rsCat.close()
				rsCat	=	null
				%>
				<input  type="hidden" value="<%=txtTable%>" name="txtTable">
				<input  type="hidden" value="<%=txtModule%>" name="txtModule">
				</form>
				<tr bgcolor="#F4F4A0">
					<td align="Center" colspan="9"><strong><i>Total</i></strong></td>
					<td align="Right"><strong><i><%=FormatNumber(v_Amount,2)%></i></strong></td>
				</tr>
                </tbody>
			</table>
			<!--#include file="Bottom_Report.asp"-->
		</div>
	</body>
</html>