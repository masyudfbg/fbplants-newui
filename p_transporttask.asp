<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Transport Task</title>
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

v_txtORDERDATE_From=Trim(REQUEST("ORDERDATE_From"))
v_txtORDERDATE_To=Trim(REQUEST("ORDERDATE_To"))
v_txtORDERNO=Trim(REQUEST("ORDERNO"))
v_txtSTOCKCODE=Trim(REQUEST("STOCKCODE"))
v_txtJOBCODE=Trim(REQUEST("JOBCODE"))
businessunit = request("businessunit")
fcbadistination = request("fcbadistination")

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

txtSQL = " SELECT TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) ORDERDATE" & _
         "        , ID.ORDERNO" & _
		 "        , TT.ROWSTATEID" & _
         "        , ID.STOCKCODE" & _
         "        , STOCK_API.get_stockName(ID.STOCKCODE,ID.Fcba) StockNAME" & _
         "        , ID.JOBCODE" & _
         "        , Job_Api.get_jobcode_name(ID.Jobcode,ID.Fcba) JOBNAME" & _
		 "        , STG.FCNAME AS Stockgroup" & _
         "        , ID.LOCATIONCODE" & _
         "        , ID.LOCATIONTYPE" & _
         "        , NVL(ID.QUANTITYTRANSFERED,0) AS QUANTITYTRANSFERED" & _
         "        , NVL(ID.PRICE,0) AS PRICE" & _
         "        , NVL(ID.AMOUNT,0) AS AMOUNT" & _
         "        , ID.REMARKS" & _
         "        , ID.LINENOHIST" & _
         "        , ID.PRICEBY" & _
         "        , ID.DOCUMENTNO" & _
         "        , ID.LINEBY" & _
         "        , STORECODE" & _
         "        , FCBADISTINATION" & _
         "        ,RECEIVEDTASKDTL_API.GET_ORDERNO(ID.FCBA,ID.ORDERNO,ID.LINENO) orderdistinantion" & _
		 "        , QUANTITYRECEIVED" & _
		 "        , TT.POSTINGFLG" & _
		 "        , TT.FCBA" & _ 
         " FROM TRANSPORTTASKDTL_TAB ID" & _
		 " JOIN TRANSPORTTASK_TAB TT ON (TT.ORDERNO = ID.ORDERNO AND TT.FCBA = ID.FCBA)" & _
		 " JOIN STOCK ST ON (ID.STOCKCODE=ST.FCCODE AND ST.FCBA=TT.FCBA)" & _
		 " JOIN STOCKGROUP STG on (TT.FCBA = STG.FCBA AND STG.FCCODE=ST.STOCKGROUPCODE)" & _
         " WHERE  ID.FCBA = '" & businessunit & "'"

if  len(trim(v_txtORDERDATE_From)) > 0 then
    txtSQL = txtSQL & " AND TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) >=to_Date('" & v_txtORDERDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtORDERDATE_From,1)
end if
if  len(trim(v_txtORDERDATE_To)) > 0 then
    txtSQL =txtSQL &  " AND TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) <=to_Date('" & v_txtORDERDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  To " & FormatDateTime(v_txtORDERDATE_To,1)
end if
if  len(trim(v_txtORDERNO)) > 0 then
    txtSQL = txtSQL & " AND ID.ORDERNO='" & v_txtORDERNO & "'"
	txtSelect = txtSelect & "  ORDER No " & v_txtORDERNO
end if
if  len(trim(v_txtSTOCKCODE)) > 0 then
    txtSQL =txtSQL &  " AND ID.stockcode='" & v_txtSTOCKCODE & "'"
	txtSelect = txtSelect & "  Stock Code " & v_txtSTOCKCODE
end if 
if  len(trim(v_txtJOBCODE)) > 0 then
    txtSQL =txtSQL &  " AND ID.jobcode='" & v_txtJOBCODE & "'"
	txtSelect = txtSelect & "  Job Code " & v_txtJOBCODE
end if 
if  len(trim(fcbadistination)) > 0 then
    txtSQL =txtSQL &  " AND ID.fcbadistination='" & fcbadistination & "'"
	txtSelect = txtSelect & "  Business Unit Destination " & fcbadistination
end if 

txtsql=(txtsql) & 	" ORDER BY ID.ORDERNO,ID.LINENO, TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO)"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="25" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>TRANSPORT TASK DETAILS</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="25" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
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
								<th>Order No</th>
								<th>Stock Code</th>
								<th>Stock Desc.</th>
								<th>Stock Group</th>
								<th>Jobcode</th>
								<th>Job Description</th>
								<th>Target Type</th>
								<th>Target Code</th>
								<th>Remarks</th>
								<th>Quantity Transfer</th>
								<th>Price By</th>
								<th>Document No</th>
								<th>Line By</th>
								<th>Price</th>
								<th>Amount</th>
								<th>Line No Hist</th>
								<th>Fcba Distination</th>
								<th>Order Distinantion</th>
								<th>Store</th>
								<th>Quantity Received</th>
								<th>State</th>
								<th>Posting</th>
							</tr>
                            </thead>
                            <tbody>
							<%
								Dim v_Quantity,v_price,v_amount,posting
								v_Quantity=0
								v_price=0
								v_amount=0
								While Not rsCat.EOF
	    
								' if cdbl(rscat("QUANTITYTRANSFERED")) <>0 and not isnull(rscat("QUANTITYTRANSFERED")) then
								'    v_price=cdbl(rscat("amount")) / cdbl(rscat("QUANTITYTRANSFERED"))
								' else
								'    v_price=0   
								' end if   
								posting = rscat("postingflg")
								if posting = "1" then
									posting = "YA"
								else 
									posting = "Tidak"
								end if
							%>
							<tr>
								<td align="Left" style="<%=cmdstyle%>">
									<a href="TransportTask.Asp?OrderNo=<%=rscat("ORDERNO")%>&fcba=<%=rscat("fcba")%>&txtJenis=Edit">
										<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat("ORDERNO")%>" border="0">
									</a>
                                </td>
								<td>
									<a href="javascript:launchwin('Print_transporttask.Asp?ORDERNO=<%=rscat("ORDERNO")%>&Mode=Print','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">
										<img src="IMAGES/print.gif" alt="Print Data : <%=rscat("ORDERNO")%>" border="0">
									</a>
                                </td>
								<td>
									<% if Session("CodeName") = "MUKTI_ACCT_01" Or Session("CodeName") = "CORNELIUS_ACCT" Or Session("CodeName") = "RENRI_ACCT" Or Session("CodeName") = "STAFF_OFFICER_ACCT" Or Session("CodeName") = "SR_OFFICER_ACCT_01" Then%>
										<%if rscat("rowstateid") = "Approved" then%>
											<a href="javascript:if(confirm('Apakah Anda yakin ingin melakukan Replanned untuk Order No <%=rscat("ORDERNO")%>?')) { launchwin('ResetTransportTask.Asp?OrderNo=<%=rscat("ORDERNO")%>&fcba=<%=rscat("fcba")%>&tgl=<%=rscat("ORDERDATE")%>','newwindow','width='+(400)+',height='+(200)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=0'); }">
												<img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ2X7OYNwyKXW8HKDtL3yBGG972_OEYpiY89g&s" alt="Reset Posting : <%=rscat("ORDERNO")%>" border="0" width="20" height="20">
											</a>		
										<%Else%>
											<span>-</span>
										<%End If%>
									<%Else%>
										<span>-</span>
									<%End If%>
								</td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ORDERDATE")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ORDERNO")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stockcode")%></font></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("StockNAME")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Stockgroup")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%></font></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobname")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationtype")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationcode")%></font></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></font></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QUANTITYTRANSFERED"),2)%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRICEBY")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOCUMENTNO")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LINEBY")%></font></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("PRICE"),2)%></font></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount"),2)%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("linenohist")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcbadistination")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("orderdistinantion")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("storecode")%></font></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("quantityreceived"),2)%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rowstateid")%></font></td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=posting%></font></td>
							</tr>
							  <%
							  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
							  v_Quantity  =v_Quantity+cdbl(rsCat("QUANTITYTRANSFERED"))
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
							<td align="Center" style="<%=cmdstyleLeft%>" colspan="11"> <font face="Verdana" size="1"><b>T o t a l</b></font></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><b><%=FormatNumber(v_Quantity,2)%></b></font></td>
							<td align="Right" colspan = "4" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></font></td> 
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><b><%=FormatNumber(v_amount,2)%></b></font></td>
							<td align="Right" colspan="6" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></font></td> 
							</tr>
                            </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
			
		</DIV>
	</BODY>
</HTML>