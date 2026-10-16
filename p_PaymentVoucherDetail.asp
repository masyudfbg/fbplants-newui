<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
function f_title(transtype_, banktype_)
	result = ""
	if Ucase(banktype_) = "JOURNAL" then result = "Journal Voucher" else result = PCase(transtype_) & " Voucher"
	
	f_title = result
end function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Payment Voucher</title>
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
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtDATECREATED_From= trim(request("DATECREATED_From"))
v_txtDATECREATED_To=trim(request("DATECREATED_To"))
v_txtBANKCODE=trim(request("BANKCODE"))
v_txtVOUCHERCODE=  trim(request("VOUCHERCODE"))
v_txtJOBCODE=Trim(REQUEST("JOBCODE"))
v_txtLOCATIONTYPE= trim(request("LOCATIONTYPE"))
v_txtITEMCODE=trim(request("ITEMCODE"))
transtype = request("transtype")
banktype = request("banktype")
rowstateid = request("rowstateid")
remarks = request("remarks")

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

txtSQL = " select p.datecreated" & _
         "        , p.transtype" & _
         "        , p.rowstateid" & _
		 "        , p.fcentry" & _
         "        , p.fcedit" & _
		 "        , decode(nvl(p.bano,''),'',p.remarks,p.remarks||'('||p.bano||')') headerremarks" & _
		 "        , p.vouchercode" & _
		 "        , pay.lineno" & _
         "        , pay.jobcode" & _
         "        , job.fcname" & _
         "        , pay.locationtype" & _
         "        , pay.locationcode" & _
         "        , pay.reference" & _
		 "        , pay.chargejob" & _
         "        , pay.chargetype" & _
		 "        , pay.chargecode" & _
         "        , nvl(sum(pay.amount),0) as amount" & _
         "        , pay.reference" & _
         "        , pay.remarks" & _
		 "        , p.bankcode" & _
		 "        , bk.fcname bankname" & _
		 "        , p.ACCOUNTINGPERIOD || ' - ' || p.ACCOUNTINGYEAR ACTPERIOD" & _
		 "		  , case when p.paymenttype='C' Then 'Cash' WHEN p.paymenttype='G' Then 'Giro' Else '' End as paymenttype" & _
		 "		  , pay.girono,pay.girodate" & _
		 " from paymentvoucher p" & _
		 " LEFT JOIN paymentvoucherdetail pay ON ((pay.vouchercode = p.vouchercode) and (pay.fcba = p.fcba)) " & _
         " LEFT JOIN job job ON ((pay.jobcode = job.fccode) and (pay.fcba = job.fcba)) " & _
         " INNER JOIN bank bk ON (( p.bankcode = bk.fccode) AND (p.fcba = bk.fcba))"& _
		 " where p.fcba = '" & businessunit & "'"
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
if  len(trim(rowstateid)) > 0 then
    txtSQL =txtSQL &  " and p.rowstateid='" & rowstateid & "'"
	txtSelect = txtSelect & " Status " & rowstateid
end if 
if  len(trim(remarks)) > 0 then
    txtSQL =txtSQL &  " and p.remarks like '%" & remarks & "%'"
	txtSelect = txtSelect & " Paid For " & remarks
end if 
if len(trim(""&transtype)) > 0 then txtSQL = txtSQL & " AND UPPER(p.transtype) = UPPER('" & transtype & "')"
if len(trim(""&banktype)) > 0 then txtSQL = txtSQL & " AND UPPER(bk.banktype) = UPPER('" & banktype & "')"

txtsql=(txtsql) & " group by p.datecreated" & _
				  "          , p.transtype" & _
				  "          , p.rowstateid" & _
                  "          , p.vouchercode" & _
				  "          , p.fcentry" & _
                  "          , p.fcedit" & _
				  "          , pay.lineno" & _
                  "          , pay.jobcode" & _
                  "          , job.fcname" & _
                  "          , pay.reference" & _
                  "          , pay.locationtype" & _
                  "          , pay.locationcode" & _
				  "          , pay.chargejob" & _
				  "          , pay.chargetype" & _
				  "          , pay.chargecode" & _
                  "          , pay.REFERENCE" & _
				  "          , p.bano,p.remarks " & _
                  "          , pay.remarks" & _
				  "          , p.bankcode" & _
				  "          , bk.fcname" & _
				  "          , p.ACCOUNTINGPERIOD" & _
				  "          , p.ACCOUNTINGYEAR" & _
				  "          , p.paymenttype" & _
				  "          , pay.girono" & _
				  "          , pay.girodate" 
				  
txtsql	=	txtsql & " Order By p.datecreated, p.vouchercode, lineno asc, amount Desc"
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
         <table border="0" width="100%" id="product-table" class="xthf-site" data-rpt-fit="1">
         <thead>
         <tr>
         <th colspan="24" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
               <div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
                <div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h1 class="judul">Bank Activity</h1>
				</div>
				<div style="clear:both;"></div>
			</div>
            </th>
			</tr>
				<!--<div class="printhide">-->
				<tr>
					<th colspan="24" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
				<!--</div>-->
				<tr>
					 
					<th colspan="2"></th>
					<th>Date</th>
					<th>Code</th>
					<th>Payment Type</th>
					<th>Giro No</th>
					<th>Giro Date</th>
					<th>Bank</th>
					<th>Line No</th>
					<th>Activity Code</th>
					<th>Activity Name</th>
					<th>Target Type</th>
					<th>Target Code</th>
					<th>Charge Job</th>
					<th>Charge Type</th>
					<th>Charge Code</th>
					<%if Request("txttable") = "02-paymentvoucher" or Request("transtype") = "02-payment" then%>
						<th>Paid For</th>
					<%else if Request("txttable") = "01-receivevoucher" or Request("transtype") = "01-receive" then%>
						<th>Receive By</th>
					<%else%>
						<th>Journal By</th>
					<%end if
					end if%>
					<th>Remarks</th>
					<th>Reference</th>
					<th>Amount</th>
					<th>Status</th>
					<th>Act Periode</th>
					<th >FCENTRY</th>
					<th >FCEDIT</th>
				</tr>
                </thead>
                <tbody>
				<%
				Dim v_Opening, v_Debit, v_Credit, v_Saldo, txtrowstateid	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO	=	1
				v_Amount	=	0
				While Not rsCat.EOF
					txtrowstateid = rscat("rowstateid")
				%>
				<tr>
					<%
					if request("txtTable") ="03-journalvoucher" then
						txtFileReportName = "print_journal_voucher"
					end if 
					if request("txtTable") ="01-receivevoucher" then 
						if  txtrowstateid = "Confirmed" then 
							txtFileReportName = "printrecvoucherfin"
						
						else 
							txtFileReportName = "printvoucherfinapprove"
						end if
					else
						
						if  txtrowstateid = "Confirmed" then 
							txtFileReportName = "printvoucherfin"
						
						else 
							txtFileReportName = "printvoucherfinapprove"
						end if
					end if

					Dim bln, thn
					bln = Month(Date)
					thn = Year(Date)

					qLineMax = " select KASBANK_CLOSING from EMPALLDEDSTATUS where fcba = '"& session("fcBA") &"' and YEAR = "&thn&" and MONTH = '"&bln&"'" 
					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					%>
					<td valign="middle" width="10">
						<%if rsqLineMax("KASBANK_CLOSING") = "FALSE" then%>
							<a href="PaymentVoucherDetail.Asp?fcCode=<%=rscat("VoucherCode")%>&txtlineno=<%=rscat("lineno")%>&transtype=<%=transtype%>&banktype=<%=banktype%>&xCode=<%=rscat("bankcode")%>&txtJenis=Edit">
								<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
							</a>
						<%end if%>
					</td>
					
					<%if  txtrowstateid = "Planned" then%> 
					<td>
						
							
					</td>
					<%else%>
					<td>
							<a href="<%=txtFileReportName%>.asp?fcba=<%=businessunit%>&vouchercode=<%=rscat("Vouchercode")%>">
							<img src="IMAGES/print.gif" alt="Print Data : <%=rscat(1)%>" border="0">
							</a>
					</td>
					<%end if%>
										
					<td align="Center"><%=rscat("Datecreated")%></td>
					<td align="center"><%=rscat("Vouchercode")%></td>
					<td align="center"><%=rscat("paymenttype")%></td>
					<td align="center"><%=rscat("girono")%></td>
					<%if rscat("paymenttype")="Giro" then%>
					<td align="center"><%=rscat("girodate")%></td>
					<%else%>
					<td align="center"></td>
					<%end if%>
					<td align="center"><%=rscat("bankname")%></td>
					<td align="center"><%=rscat("lineno")%></td>
					<td align="center"><%=rscat("Jobcode")%></td>
					<td align="center"><%=rscat("Fcname")%></td>
					<td align="center"><%=rscat("Locationtype")%></td>
					<td align="center"><%=rscat("Locationcode")%></td>
					<td align="center"><%=rscat("chargejob")%></td>
					<td align="center"><%=rscat("chargetype")%></td>
					<td align="center"><%=rscat("chargecode")%></td>
					<td align="left"><%=rscat("headerremarks")%></td>
					<td align="left"><%=rscat("Remarks")%></td>
					<td align="left"><%=rscat("reference")%></td>
					<td align="Right"><%=FormatNumber(rscat("Amount"),2)%></td>
					<td align="center"><%=rscat("rowstateid")%></td>
					<td align="center"><%=rscat("ACTPERIOD")%></td>
					<td align="center"><%=rscat("fcentry")%></td>
					<td align="center"><%=rscat("fcedit")%></td>
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
				<input  type="hidden" value="<%=txtKind%>" name="txtKind">
				<input  type="hidden" value="<%=txtModule%>" name="txtModule">
				</form>
				<tr bgcolor="#F4F4A0">
					<td align="Right" colspan="19"><strong><i>Total</i></strong></td>
					<td align="Right"><strong><i><%=FormatNumber(v_Amount,2)%></i></strong></td>
					<td align="Right"></td>
				</tr>
                </tbody>
			</table>
			<!--#include file="Bottom_Report.asp"-->
		</div>
	</body>
</html>