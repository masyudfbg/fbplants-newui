<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Stock Issue</title>
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

v_txtSIVDATE_From=Trim(REQUEST("SIVDATE_From"))
v_txtSIVDATE_To=Trim(REQUEST("SIVDATE_To"))
v_txtSIVCODE=Trim(REQUEST("SIVCODE"))
v_txtSTOCKCODE=Trim(REQUEST("STOCKCODE"))
v_txtJOBCODE=Trim(REQUEST("JOBCODE"))
v_txtLOCATIONTYPE=Trim(REQUEST("LOCATIONTYPE"))
v_txtLOCATIONCODE=Trim(REQUEST("LOCATIONCODE"))

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

txtSQL = "SELECT issue.fcba,issue.postingflg,issue.storecode,issue.sivdate,sub.FCNAME kodesub,iss.sivcode, iss.stockcode, stock.fcname, iss.jobcode, iss.locationcode,issue.rowstate," & _
		 "iss.locationtype, nvl(sum(iss.quantity),0) as quantity, nvl(sum(iss.price),0) as price, nvl(sum(iss.amount),0) as amount, job.fcname as jobname, iss.reference as reference, iss.remarks as remarks, iss.linenohist" & _
		 " FROM issuedetails iss, issue issue, stock stock, job job, STOCKGROUP sub" & _
		 " WHERE (((iss.sivcode = issue.sivcode) and (iss.fcBA = issue.fcBA))" & _
		 " AND ((iss.stockcode = stock.fccode) and (iss.fcBA= stock.fcBa)) " & _
		 " AND ((sub.fccode= stock.stockgroupcode) and (iss.fcBA= sub.fcBa)) " & _
		 " AND ((iss.jobcode = job.fccode) and (iss.fcBA = job.fcBA)) )" & _
		 " AND iss.fcba='" & businessunit & "'"

if  len(trim(v_txtSIVDATE_From)) > 0 then
    txtSQL = txtSQL & " and issue.sivdate >=to_Date('" & v_txtSIVDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtSIVDATE_From,1)
end if
if  len(trim(v_txtSIVDATE_To)) > 0 then
    txtSQL =txtSQL &  " and issue.sivdate <=to_Date('" & v_txtSIVDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  To " & FormatDateTime(v_txtSIVDATE_To,1)
end if
if  len(trim(v_txtSIVCODE)) > 0 then
    txtSQL = txtSQL & " and iss.sivcode like '%" & v_txtSIVCODE & "%'"
	txtSelect = txtSelect & "  SIV Code " & v_txtSIVCODE
end if
if  len(trim(v_txtSTOCKCODE)) > 0 then
    txtSQL =txtSQL &  " and iss.stockcode='" & v_txtSTOCKCODE & "'"
	txtSelect = txtSelect & "  Stock Code " & v_txtSTOCKCODE
end if 
if  len(trim(v_txtJOBCODE)) > 0 then
    txtSQL =txtSQL &  " and iss.jobcode='" & v_txtJOBCODE & "'"
	txtSelect = txtSelect & "  Job Code " & v_txtJOBCODE
end if 
if  len(trim(v_txtLOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and iss.locationtype='" & v_txtLOCATIONTYPE & "'"
	txtSelect = txtSelect & "  Target Type " & v_txtLOCATIONTYPE
end if 
if  len(trim(v_txtLOCATIONCODE)) > 0 then
    txtSQL =txtSQL &  " and iss.locationcode='" & v_txtLOCATIONCODE & "'"
	txtSelect = txtSelect & "  Target Code " & v_txtLOCATIONCODE
end if 

txtsql=(txtsql) & 	" group by issue.fcba,issue.postingflg,sub.FCNAME,issue.storecode,issue.sivdate,iss.sivcode,issue.rowstate, iss.stockcode, stock.fcname, iss.jobcode," & _
					" iss.locationcode, iss.locationtype, job.fcname, iss.reference, iss.remarks, iss.linenohist  order by issue.sivdate,iss.sivcode,iss.stockcode"
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
         <th colspan="19" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=FcbaName%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>ISSUE DETAILS</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="19" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
							<tr>
							<%	
							Set rsCat = connStr.Execute(txtSQL)
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
							</tr>
							<tr>
								<th colspan="3"></th>
								<th>Date</th>
								<th>Issue Code</th>
								<th>Stock Group</th>
								<th>Stock Code</th>
								<th>Stock Desc.</th>
								<th>Jobcode</th>
								<th>Job Description</th>
								<th>Target Type</th>
								<th>Target Code</th>
								<th>Quantity</th>
								<th>Price</th>
								<th>Amount</th>
								<th>Line No Hist</th>
								<th>Reference</th>
								<th>Remarks</th>
								<th>Rowstate</th>
								<th>Posting</th>
								
							</tr>
                            </thead>
                            <tbody>
							<%
								Dim v_Quantity,v_price,v_amount
								v_Quantity=0
								v_price=0
								v_amount=0
								While Not rsCat.EOF
	    
								if cdbl(rscat("quantity")) <>0 and not isnull(rscat("quantity")) then
								   v_price=cdbl(rscat("amount")) / cdbl(rscat("quantity"))
								else
								   v_price=0   
								end if   
							%>
							<tr>
								<td align="Left" style="<%=cmdstyle%>">
									<a href="IssueDetails.Asp?fcCode=<%=rscat("sivcode")%>&fcba=<%=rscat("fcba")%>&xCode=<%=rscat("storecode")%>&txtJenis=Edit">
										<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
									</a>
                                </td>								
								<td>
									<a href="javascript:launchwin('PrintSIV.Asp?fcCode=<%=rscat("sivcode")%>&fcba=<%=rscat("fcba")%>&Mode=Print','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
										<img src="IMAGES/print.gif" alt="Print Data : <%=rscat(1)%>" border="0">
									</a>									
                                </td>
								<td>
									<% if Session("CodeName") = "MUKTI_ACCT_01" Or Session("CodeName") = "BONITA_ACC" Or Session("CodeName") = "CORNELIUS_ACCT" Or Session("CodeName") = "RENRI_ACCT" Or Session("CodeName") = "STAFF_OFFICER_ACCT" Or Session("CodeName") = "SR_OFFICER_ACCT_01" Then%>
										<%if rscat("rowstate") = "Approved" Or rscat("rowstate") = "Planned" then%>
											<a href="javascript:if(confirm('Apakah Anda yakin ingin Replanned posting untuk SIV <%=rscat("sivcode")%>?')) { launchwin('ResetSIV.Asp?fcCode=<%=rscat("sivcode")%>&fcba=<%=rscat("fcba")%>&tgl=<%=rscat("sivdate")%>','newwindow','width='+(400)+',height='+(200)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=0'); }">
												<img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ2X7OYNwyKXW8HKDtL3yBGG972_OEYpiY89g&s" alt="Reset Posting : <%=rscat("sivcode")%>" border="0" width="20" height="20">
											</a>		
										<%Else%>
											<span>-</span>
										<%End If%>
									<%Else%>
										<span>-</span>
									<%End If%>
                                </td>	
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sivdate")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sivcode")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("kodesub")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stockcode")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobname")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationtype")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationcode")%></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("quantity"),2)%></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_price,2)%></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount"),2)%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("linenohist")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("reference")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rowstate")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%if rscat("postingflg")="1" then response.write "Ya" else response.write "Tidak" end if%></td>
							</tr>
							  <%
							  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
							  v_Quantity  =v_Quantity+cdbl(rsCat("quantity"))
							  v_amount =v_amount+cdbl(rsCat("amount"))
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
							<td align="Center" style="<%=cmdstyleLeft%>" colspan="11"> <font face="Verdana" size="1">T o t a l</td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amount,2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
							</tr>
                            </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
			
		</DIV>
	</BODY>
</HTML>