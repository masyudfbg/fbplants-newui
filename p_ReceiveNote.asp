<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Receive Note</title>
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


v_txtRNDATE_From=trim(request("RNDATE_From"))
v_txtRNDATE_To=trim(request("RNDATE_To"))
v_txtRECEIVENOTECODE=trim(request("RECEIVENOTECODE"))
v_txtPURCHASEITEMCODE=trim(request("PURCHASEITEMCODE"))
v_txtSUPPLIERCODE=trim(request("SUPPLIERCODE"))
v_txtJenis = trim(request("v_txtJenis"))
Dim v_txtJobCode : v_txtJobCode = trim(request("jobcode"))

txtSQL =	" SELECT RN.FCBA,RN.Postingflg,RN.rowstateid, RN.RECEIVENOTECODE" & _
			"        , RN.SUPPLIERCODE" & _
			"        , supplier_api.Get_Supp_Name(RN.SUPPLIERCODE, RN.FCBA) SUPPLIERNAME" & _
			"        , RN.POCODE" & _
			"		 , RN.STOCKORDERNO" & _
            "   	 , RN.DELIVERYORDERCODE " & _
			"        , RN.RNDATE" & _
			"        , RND.JOBCODE" & _
			"        , job_api.get_jobcode_name(RND.JOBCODE, RND.FCBA) JOBNAME" & _
			"        , RND.LOCATIONTYPE" & _
			"        , RND.LOCATIONCODE" & _
			"        , RND.PURCHASEITEMCODE" & _
			"        , RND.PURCHASELINENOKEY" & _
			"        , RND.LINENOHIST" & _
			"        , STOCK.FCNAME" & _
			"        , STOCK.UNITOFMEASUREMENT" & _
			"        , NVL(SUM(NVL(RND.QUANTITY, 0)),0) AS QUANTITY" & _
			"        , NVL(SUM(NVL(RND.QTYRETURN, 0)),0) AS QTYRETURN" & _
			"        , NVL(SUM(NVL(RND.AMOUNTRECEIVED, 0)),0) AS AMOUNTRECEIVED" & _
			"        , NVL(SUM(NVL(RND.AMOUNTRETURN, 0)),0) AS AMOUNTRETURN" & _
			" FROM RECEIVENOTE RN" & _
			"      , RECEIVENOTEDETAIL RND" & _
			"      , STOCK STOCK" & _
			" WHERE ( ((RN.RECEIVENOTECODE = RND.RECEIVENOTECODE) AND (RN.FCBA = RND.FCBA))" & _
			"         AND ((STOCK.FCCODE = RND.PURCHASEITEMCODE) AND (STOCK.FCBA = RND.FCBA))  )" & _
			"         AND RN.FCBA = '" & session("fcBA") & "'"

if  len(trim(v_txtRNDATE_From)) > 0 then
    txtSQL = txtSQL & " and RN.rndate >=to_Date('" & v_txtRNDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtRNDATE_From,1)
end if
if  len(trim(v_txtRNDATE_To)) > 0 then
    txtSQL =txtSQL &  " and RN.rndate <=to_Date('" & v_txtRNDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect &" To " & FormatDateTime(v_txtRNDATE_To,1)
end if
if  len(trim(v_txtRECEIVENOTECODE)) > 0 then
    txtSQL = txtSQL & " and RN.receivenotecode='" & v_txtRECEIVENOTECODE & "'"
	txtSelect = txtSelect &" RN Code " & v_txtRECEIVENOTECODE
end if
if  len(trim(v_txtPURCHASEITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and RND.PURCHASEITEMCODE='" & v_txtPURCHASEITEMCODE & "'"
	txtSelect = txtSelect &" PO Code " & v_txtPURCHASEITEMCODE
end if 
if  len(trim(v_txtSUPPLIERCODE)) > 0 then
    txtSQL =txtSQL &  " and RN.SUPPLIERCODE='" & v_txtSUPPLIERCODE & "'"
	txtSelect = txtSelect & " Supplier Code " & v_txtSUPPLIERCODE
end if 

if  len(trim(v_txtJobCode)) > 0 then
    txtSQL = txtSQL &  " and RND.JOBCODE = '" & v_txtJobCode & "'"
	txtSelect = txtSelect & " Jobcode " & v_txtJobCode
end if 

txtsql	=	txtsql &	" GROUP BY RN.FCBA,RN.rowstateid,RN.Postingflg,RN.RECEIVENOTECODE" & _
						"          , RN.SUPPLIERCODE" & _
						"          , supplier_api.Get_Supp_Name(RN.SUPPLIERCODE, RN.FCBA)" & _
						"          , RN.POCODE" & _
						"		   , RN.STOCKORDERNO" & _
                        "		   , RN.DELIVERYORDERCODE " & _
						"          , RN.RNDATE" & _
						"          , RND.JOBCODE" & _
						"          , job_api.get_jobcode_name(RND.JOBCODE, RND.FCBA)" & _
						"          , RND.LOCATIONTYPE" & _
						"          , RND.LOCATIONCODE" & _
						"          , RND.PURCHASEITEMCODE" & _
						"          , RND.PURCHASELINENOKEY" & _
						"          , RND.LINENOHIST" & _
						"          , STOCK.FCNAME" & _
						"          , STOCK.UNITOFMEASUREMENT" & _
						" ORDER BY RN.RECEIVENOTECODE" & _
						"          , RN.POCODE,RND.PURCHASELINENOKEY"

'response.Write(txtsql)
response.Flush()

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " SELECT 'Y' istrue FROM users_tab " & _
			" WHERE fccode = '" & Session("CodeName") & "' AND MENU_GRN = 'TRUE'" 
' response.write oSQLQuery
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if

%>
 <body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;">
					<h2>RECEIVE NOTE</h2>
				</div>
			</div>
			</th>
			</tr>
				<!--<div class="printhide">-->
						<tr>
							<th colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
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
							<th>GRN No</th>
							<th>GRN Date</th>
							<th>Supplier Code</th>
							<th>Supplier Description</th>
							<th>PO No</th>
							
							<th>STO No</th>
							<th>DO No</th>
						
							<th>Job Code</th>
							<th>Job Description</th>
							<th>Target Type</th>
							<th>Target Code</th>
							<th>Line No History</th>
							<th>Stock Code</th>
							<th>Stock Description</th>
							<th>UOM</th>
							<th>Quantity</th>
							<%If val_IsTrue = "Y" Then%>
							<th>Amount</th>
							<%End If%>
							<th>Rowstate</th>
							<th>Posting</th>
						</tr>
                        </thead>
                        <tbody>
							<%
							Dim v_Quantity,v_qtyretrun,v_amountreceived,v_amountreturn
							v_Quantity=0
							v_qtyretrun=0
							v_amountreceived=0
							v_amountreturn=0
							While Not rsCat.EOF%>
							<tr>
								<td align="Left" style="<%=cmdstyle%>">
									<a href="ReceiveNote.Asp?fcCode=<%=rscat("receivenotecode")%>&fcba=<%=rscat("fcba")%>&xCode=<%=rscat("pocode")%>&txtJenis=Edit">
										<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
									</a>
								</td>
								<td>
									<a href="javascript:launchwin('goods_received_note.asp?grnno=<%=rscat("receivenotecode")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
										<img src="IMAGES/print.gif" alt="Print Data : <%=rscat(1)%>" border="0">
									</a>
								</td>
								<td>
									<% if Session("CodeName") = "MUKTI_ACCT_01" Or Session("CodeName") = "CORNELIUS_ACCT" Or Session("CodeName") = "RENRI_ACCT" Or Session("CodeName") = "STAFF_OFFICER_ACCT" Or Session("CodeName") = "SR_OFFICER_ACCT_01" Then%>
										<%if rscat("rowstateid") = "Approved" Or rscat("rowstateid") = "Planned" then%>
											<a href="javascript:if(confirm('Apakah Anda yakin ingin Replanned posting untuk GRN <%=rscat("receivenotecode")%>?')) { launchwin('ResetGRN.Asp?fcCode=<%=rscat("receivenotecode")%>&fcba=<%=rscat("fcba")%>&tgl=<%=rscat("rndate")%>','newwindow','width='+(400)+',height='+(200)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=0'); }">
												<img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ2X7OYNwyKXW8HKDtL3yBGG972_OEYpiY89g&s" alt="Reset Posting : <%=rscat("receivenotecode")%>" border="0" width="20" height="20">
											</a>		
										<%Else%>
											<span>-</span>
										<%End If%>
									<%Else%>
										<span>-</span>
									<%End If%>
                                </td>							
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("receivenotecode")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rndate")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("suppliercode")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("suppliername")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pocode")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STOCKORDERNO")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DELIVERYORDERCODE")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobname")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationtype")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationcode")%></td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LINENOHIST")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("purchaseitemcode")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("unitofmeasurement")%></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("quantity"),2)%></td>
								<%If val_IsTrue = "Y" Then%>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amountreceived"),2)%></td>
								<%End If%>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rowstateid")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%if rscat("postingflg")="1" then response.write "Ya" else response.write "Tidak" end if%></td>
							</tr>
						  <%
						  v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
						  'v_qtyretrun  =v_qtyretrun+Trim(cdbl(rsCat("qtyreturn")))
						  v_amountreceived =v_amountreceived+Trim(cdbl(rsCat("amountreceived")))
						  'v_amountreturn  =v_amountreturn+trim(cdbl(rsCat("amountreturn")))
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
							<td align="Center" style="<%=cmdstyleLeft%>" colspan="16"> <font face="Verdana" size="1"><strong>T o t a l</strong></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_Quantity,2)%></strong></td>
							<%If val_IsTrue = "Y" Then%>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_amountreceived,2)%></strong></td>
							<%End If%>	
							</tr>
                         </tbody>   
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
	<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
	<script>
		function UpdateData(no, fcba){
			var konfirmasi = confirm("Apakah Anda yakin ingin membatalkan GRN "+no+" ini?");

			if (konfirmasi) {
				$.post({
					url: "bkm_viewdata.asp?idview=41&fccode=" + no + "&fcba=" + fcba,
					dataType: "json", // ini penting
					success: function(data) {
						console.log(data.status);  // "OK"
						console.log(data.message); // "Data berhasil diupdate"
						alert(data.message);
						window.location.reload();
					},
					error: function(xhr, status, err) {
						console.error("Error:", err);
					}
				});
			} else {
				// Jika user pilih "No" / Cancel
				alert("Update dibatalkan.");
			}
		}
	</script>
</HTML>	
