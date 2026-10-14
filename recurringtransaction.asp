<%
	Option Explicit
	response.buffer = True
	validatesession
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include file="adovbs.inc"-->
<html>
<head>
	<title>ePlaS - Recurring Transaction</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<link rel="stylesheet" href="css/jquery.dataTables.min.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<script type='text/javascript'>
    //Replaces all non numeric characters
    function checkInput(element){ element.value = element.value.replace(/[^0-9\.]/g,'');}
</script>
</head>
<body>
	<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		<% if Session("fcBa") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;">
<%
'--Awal--
Dim v_txtagreementcode, v_txtmonth, v_txtyear, MinYr, MaxYr, mon, Yr, n, SudahClose, v_txtagreementfinalno
Dim xCode,xstoreCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,v_CrossCekPO
Dim Remarks
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
Dim Jobcode, LocationType, LocationCode, Reference, Remark
Dim qLineMax, rsqLineMax, maxagreementfinalno
Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
Dim sqlSQLCheck, rsSQL, qtyRemained
Dim qCheckStore, rsqCheckStore
Dim txtLocationType	
Dim rsCheckRate, checkRate, varRate
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
Dim wipdate : wipdate = ""

Dim stringduedate : stringduedate = 0

Dim IsOverDuedate, varRemainDueDate, varDueDate, CompleteStringError, varErrorDueDate, stringremainqty, varRemainQty, IsOverRemainQty, varErrorRemainQty, varEditedQty

Dim qtyplanedit : qtyplanedit = Request("qtyplanedit")
Dim unitplanedit : unitplanedit = Request("unitplanedit")
Dim outputplanedit : outputplanedit = Request("outputplanedit")
Dim qtyfinaledit : qtyfinaledit = Request("qtyfinaledit")
Dim unitfinaledit : unitfinaledit = Request("unitfinaledit")
Dim outputfinaledit : outputfinaledit = Request("outputfinaledit")
Dim txtagreementlineno : txtagreementlineno = Request("txtagreementlineno")

Dim ComparetoAcc : ComparetoAcc = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0
Dim payroll_executed
payroll_executed = "FALSE"

Dim documentno, transactionno, fcba__, documentno__, transactionno__, executedate__, fcentry__, fcedit__, fcip__, ResultCommand__, comm
Dim prmfcba, prmdocumentno,prmtransactionno, prmexecutedate, prmfcentry, prmfcedit, prmfcip, prmresult, rsrecurring
documentno = Request("documentno")
transactionno = Request("transactionno")

txtTable = Request("txtTable")
txtModule = Request("txtModule")
txtSearch = Request("txtSearch")
txtField = Request("txtField")

v_txtmonth = Request("txtmonth")
v_txtyear = Request("txtyear")

Yr = Year(Session("v_Bank"))
mon = Month(Session("v_Bank"))
MinYr = Yr - 2
MaxYr = Yr  + 2

If len(v_txtyear)= 0 then v_txtyear = Yr
If len(v_txtmonth) = 0 then v_txtmonth = mon

txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
			
			'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")="Posting" Then
	if Request("txtJenis")= "Header" then
		
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		
		cmdView	= "True"					
	elseif Request("txtJenis")="New" then
		
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		
		cmdView	= "True"
		
	elseif Request("txtJenis")="Posting" then
		oSql = " SELECT NVL(rec - curr, '-100') AS result" & _
               " FROM (" & _
               "     SELECT TO_NUMBER(TO_CHAR(currentdate, 'yyyymm')) curr, " & v_txtyear & "||LPAD(" & v_txtmonth & ", 2, 0) rec" & _
               "     FROM periodcontrol" & _
               "     WHERE fcba = '" & Session("fcBA") & "'" & _
               "           AND UPPER(SYSTEM) = UPPER('costbook')" & _
               " )"

		set oRs = server.CreateObject("Adodb.recordset")
		oRs.Open oSql, connstr
		IF NOT oRs.EOF THEN
			ComparetoAcc = Cdbl(oRs("result"))
		END IF
		oRs.Close
		Set oRs = Nothing
		
		IF ComparetoAcc > 0 THEN
			CompleteStringError = "Periode data yang di proses melebihi periode accounting"
		ELSEIF ComparetoAcc < 0 THEN
			CompleteStringError = "Periode data yang di proses, dibawah periode accounting (sudah closing!)"
		ELSEIF ComparetoAcc = 0 THEN
		
			oSql = " SELECT NVL(payroll_executed, 'FALSE') payroll_executed" & _
                   " FROM empalldedstatus" & _
                   " WHERE fcba = '" & Session("fcBA") & "'" & _
                   "       AND year = '" & v_txtyear & "'" & _
                   "       AND month = '" & v_txtmonth & "'"
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open oSql, connstr

			if not oRs.eof then
				payroll_executed = oRs("payroll_executed")
			else
				payroll_executed = "FALSE"
			end if
			oRs.close
			set oRs = Nothing
			
			IF payroll_executed = "FALSE" THEN
				CompleteStringError = "Payroll belum di closing, <b>Silahkan finalkan dahulu payroll nya</b>"
			ELSEIF payroll_executed = "TRUE" THEN
	
				fcba__ = Session("fcBA")
				documentno__ = documentno
				transactionno__ = transactionno
				executedate__ = textDate(v_txtmonth & "/" & GetLastDay(v_txtmonth, v_txtyear) & "/" & v_txtyear)
				' response.write executedate__
				fcentry__ = Session("CodeName")
				fcedit__ = Session("CodeName")
				fcip__ = Request.ServerVariables("REMOTE_ADDR")
				
				Set comm = Server.CreateObject("ADODB.Command")
				With comm 
					.CommandText = "{ CALL RECURRING_API.ACCRUAL_PROCESS (?,?,?,?,?,?,?,?) }"
					.CommandType = adCmdText
					.ActiveConnection = connStr
					.Prepared = True
					Set prmfcba = .CreateParameter("",adVarchar,adParamInput,255)
					Set prmdocumentno = .CreateParameter("",adVarchar,adParamInput,255)
					Set prmtransactionno = .CreateParameter("",adVarchar,adParamInput,255)
					Set prmexecutedate = .CreateParameter("",adDate,adParamInput)
					Set prmfcentry = .CreateParameter("",adVarchar,adParamInput,255)
					Set prmfcedit = .CreateParameter("",adVarchar,adParamInput,255)
					Set prmfcip = .CreateParameter("",adVarchar,adParamInput,255)
					Set prmresult = .CreateParameter("",adLongVarChar,adParamOutput,1000)
					.Parameters.Append(prmfcba)
					.Parameters.Append(prmdocumentno)
					.Parameters.Append(prmtransactionno)
					.Parameters.Append(prmexecutedate)
					.Parameters.Append(prmfcentry)
					.Parameters.Append(prmfcedit)
					.Parameters.Append(prmfcip)
					.Parameters.Append(prmresult)
				End With

				With comm 
					.Parameters(0).Value = fcba__
					.Parameters(1).Value = documentno__
					.Parameters(2).Value = transactionno__
					.Parameters(3).Value = executedate__
					.Parameters(4).Value = fcentry__
					.Parameters(5).Value = fcedit__
					.Parameters(6).Value = fcip__
					ResultCommand__ = .Parameters(7).Value
					Set rsrecurring  = .Execute
					' Response.Write(ResultCommand_)
				End With
				
			END IF
		END IF
		cmdView	= "True"
    else
	
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")		    
End if


If Len(Trim(CompleteStringError)) > 0 Then
	response.Write  "<div class = ""error"">" & CompleteStringError & "</div>"
End If

 %>
<form name="frmEntry" method="post" action="recurringtransaction.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Year/ Tahun</td>
			<td align="left">
				<select name="txtyear">
				<%For i = MinYr to MaxYr%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtyear) then%> Selected <%End If%>><%=i%></option>
				<%Next%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Month/ Bulan</td>
			<td align="left">
				<select name="txtmonth">
				<%For i = 1 to 12%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtmonth) then%> Selected <%End If%>><%=textMonth(i)%></option>
				<%Next%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left"> 
				<input Type="Hidden" Value="True" Name="cmdView">  
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">  
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">   	 
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;  
				<a href="print_recurringtransaction.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
			</td> 
		</tr> 
	</table>    
</form>
<%
If cmdView = "True" Then
	Dim intRecordsPerPage, cColor
	'response.write " v_txtyear : " & v_txtyear
%>
	<form action="recurringtransaction.asp" method="post" name="eForm">
	<table id="posting_data" class="cell-border stripe hover" cellspacing="0" width="100%">
        <thead>
            <tr>
                <th class="center"></th>
                <th class="center">Document No</th>
                <th class="center">Transaction No</th>
                <th class="center">Currency</th>
                <th class="center">Total Amount</th>
                <th class="center">Status</th>
            </tr>
        </thead>
    </table>
	</form>
<%End if%>
<!---------------akhir!--------------->
</div>
</div>
<!--#include file="bottom.jsp"-->
<script src="/js/jquery-1.12.3.js"></script>
<script src="/js/jquery.dataTables.min.js"></script>
<script src="/js/jquery-ui.js"></script>
<script>
/* Formatting function for row details - modify as you need */
function format ( d ) {	
	var jml = d.detail.length;
	var td = "";
	for (i=0;i<jml;i++) {
         td += '<tr>'+
            '<td>'+d.detail[i].debetkredit+'</td>'+
            '<td>'+d.detail[i].jobcode+'</td>'+
            '<td>'+d.detail[i].targettype+'</td>'+
            '<td align="right" >'+accounting.formatNumber(d.detail[i].amount);+'</td>'+
        '</tr>'
		}
    // `d` is the original data object for the row
    return '<table cellpadding="5" cellspacing="0" border="0" style="padding-left:50px;">'+
        '<tr>'+
            '<th>Debet/Kredit</th>'+
			'<th>Jobcode</th>'+
			'<th>Target Type</th>'+
			'<th>Amount</th>'+
        '</tr>'+ td +
    '</table>';
}
 
$(document).ready(function() {
    var table = $('#posting_data').DataTable( {
        "ajax": "/ajax/data_recurringtransaction.asp?accountingyear=<%=v_txtyear%>&accountingperiod=<%=v_txtmonth%>",
        "columns": [
            {
                "className": 'details-control',
                "orderable": false,
                "data": null,
                "defaultContent": ''
            },
            { "data": "documentno" },
            { "data": "transactionno" },
            { "data": "currency" },
            { "data": "totalamount", render: $.fn.dataTable.render.number( ',', '.', 0, '' ), "sClass": "right" },
            { "data": "status", render: function ( data, type, full ) {
									if (data.toLowerCase() == 'opened'||data.toLowerCase() == 'open') {
										return '<a href="recurringtransaction.asp?txtTable=<%=txtTable%>&txtModule=<%=txtModule%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&documentno='+full.documentno+'&transactionno='+full.transactionno+'&txtJenis=Posting">Process for posting</a>';
									} else {
										return data;
									}
								}  , "sClass": "center"
			}
        ],
        "order": [[1, 'asc']]
    } );
	
    // Add event listener for opening and closing details
    $('#posting_data tbody').on('click', 'td.details-control', function () {
        var tr = $(this).closest('tr');
        var row = table.row( tr );
 
        if ( row.child.isShown() ) {
            // This row is already open - close it
            row.child.hide();
            tr.removeClass('shown');
        }
        else {
            // Open this row
            row.child( format(row.data()) ).show();
            tr.addClass('shown');
        }
    } );
} );
</script>
<script src="js/accounting.min.js"></script>
<style>.right{text-align:right;}.left{text-align:left;}.center{text-align:center;}</style>
</body>
</html>
