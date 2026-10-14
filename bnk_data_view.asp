<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim bankcode,bulan,tahun,jobcode,fcba,q,fcbaname,CompanyName ,txtFileReportName,txtControlJob,txtSQL1,txtSQL2,txtsql,txtSelect1,txtSelect,v_Amount,v_Saldo,cmdstyleLeft,cmdstyle
dim ttldebet,ttlcredit,txtSQL3
bankcode= trim(request("bankcode"))
bulan= trim(request("month"))
tahun= trim(request("tahun"))
jobcode= trim(request("jobcode"))
fcba = trim(request("fcba"))

q = "BEGIN fac_journal_fac_bnk_api.FAC_BNK_STATUSPOST(" & tahun & "," & bulan & ",'" & fcba & "','System'); END;" 
connStr.Execute(q)

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME") 
end if

' get data header  to_date('" & IssueDate & "','mm/dd/yyyy')
dim sqlheader,rsHeader,stockcode,stockname 
sqlheader = "select fccode, fcname,CONTROLJOB from bank where fccode='" & bankcode & "' AND fcba='" & fcba & "' "
set rsHeader = connStr.Execute(sqlheader)
'response.write sqlheader
if not rsHeader.eof then 
	stockcode = rsHeader("fccode")
	stockname = rsHeader("fcname")
	txtControlJob = rsHeader("CONTROLJOB")
else
	txtControlJob="" 

end if


%>

<html>
	<head>
		<title>FBPlants - BANK VIEW</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
	</head>
<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: -10px;
	}
	.posisiheader{
		position: relative;
		left: 0px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarabsen{
		width: 60px;
	}
	.lebaremployee{
		width: 130px;
	}
	.lebaractivity{
		width: 250px;
	}
	.lebaracllowedcode{
		width: 210px;
	}
	.lebarupah{
		width: 120px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakiri{
		text-align: left;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}

	
</style>

<script>

</script>

<body>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>

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

<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BANK</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">	
<div id="wrapper">		
<form name="frmEntry" >
	<table class="data posisiheader" border="0" width="50%" id="product-table" class="xthf-site"> 			
		<tr>
			<td align="left" >BANK CODE</td>
			<td align="left">
				<input type="text" name="stockcode" readonly id="stockcode" size="20" value="<%=stockcode%>" >
			</td>
			<td align="left" >PERIODE</td>
			<td align="left">
				<%=bulan%>/<%=tahun%> 
			</td>
			
		</tr>
		
		<tr>
			<td align="left" >BANK NAME</td>
			<td align="left" colspan="">
				<input type="text" name="stockname" readonly id="stockname" size="50" value="<%=stockname%>" >				
			</td>						
		</tr>
		
		<tr>
		<td colspan="4" align="left">
		<a class="button_menu_top" onclick="createData('PYV')" ><img src="ICON/Bank.png" border="0" ><span> New Payment </span></a>&nbsp;
		<a class="button_menu_top" onclick="createData('RCV')" ><img src="ICON/Bank.png" border="0" ><span> New Receive </span></a>&nbsp;
		<a class="button_menu_top" onclick="getExport()" ><img  src="Images/butt_export_xls.png" border="0" ><span> Export XLS</span></a>&nbsp;
		<input type=button onclick="window.close();" name="btnclose" value="CLOSE">
		
		
		</td>
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='800'  id="product-table" class="xthf-site">
		<thead>
				<tr>
					<th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
						  
							<div style="clear:both;"></div>
						</div>
					</th>
				</tr> 
		</thead>
		<!--<div class="printhide">--> 
			
		<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">Payment</li>
						<li class="TabbedPanelsTab">Receive</li>
						<li class="TabbedPanelsTab">Cash Flow</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

<!-- Payment -->
						<div class="TabbedPanelsContent">
						 	<table align=center border="0" width="600" class='data' id='table2'>
								<tr id="trheader2">									
									<th>No.Voucher</th>
									<th>Date</th>
									<th>Period</th>
									<th>Remark Header</th>
									<th>Cek No</th>
									<th>Classification</th>
									<th>Paid To</th>
									<th>Keterangan</th>
									<th>Akun</th>
									<th>Amount</th>
									<th>Status</th>
									<th>Posting</th>
									<th >FCENTRY</th>
									<th >FCEDIT</th>
									<th>Action</th>	
									
								</tr>

								<% 
								
								
								'get data detail upah
								'dim sqldtlupah,rsdtlupah,no,totalamount
								sqldtlupah = " SELECT a.vouchercode,to_char(datecreated,'dd/mm/yyyy')datecreated,referenceno,cekno,classification_api.get_fcname(document_classification) class, " & _ 
											 " decode(nvl(b.bano,''),'',b.remarks,b.remarks||' ('||b.bano||')') remarksheader, a.remarks,nvl(a.amount,0) totalamount,rowstateid,transtype,bankcode,a.fcba,nvl(b.postingflg,'0') postingflg" & _
											 " ,b.accountingperiod,b.fcentry,b.fcedit,b.accountingyear,a.lineno,a.jobcode||'-'||a.locationcode Akun " & _
											 " FROM paymentvoucherdetail a " & _
											 " INNER JOIN paymentvoucher b ON (a.vouchercode=b.vouchercode AND a.fcba=b.fcba)  " & _
											 " LEFT JOIN bkmjournalhdr c ON (a.vouchercode=c.docno AND a.fcba=c.fcba) " & _
											 " where a.fcba='" & fcba & "'" & _ 
											 " AND bankcode='" & bankcode & "'" & _ 
											 " AND to_number(to_char(datecreated,'YYYY'))="& tahun & " " & _ 
											 " AND to_number(to_char(datecreated,'MM'))="& bulan & " " & _
											 " AND transtype='02-PAYMENT' " & _
											 " ORDER BY datecreated DESC, a.vouchercode DESC "
                           	
								set rsdtlupah = connStr.Execute(sqldtlupah)
							
																
								no = 0
								totalamount = 0
								While Not rsdtlupah.EOF 
									postingflg = rsdtlupah("postingflg")
									if postingflg = "0" then 
										postingflg="Tidak"
									else 
										postingflg="YA"
									end if
								%>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("vouchercode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("datecreated")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("accountingperiod")%>-<%=rsdtlupah("accountingyear")%>
									</td>
									
									<td class="ratakiri lebaremployee">
										<%=rsdtlupah("referenceno")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("cekno")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("class")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("remarksheader")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("remarks")%>
									</td>

									<td class="ratakiri">
										<%=rsdtlupah("Akun")%>
									</td>
									<td>
										<%=formatnumber(rsdtlupah("totalamount"),0)%>
									</td>
									
									<td>
										<%=rsdtlupah("rowstateid")%>
									</td>
									<td class="ratatengah">
										<%=postingflg%>
									</td>
									
									<td>
										<%=rsdtlupah("fcentry")%>
									</td>

									<td>
										<%=rsdtlupah("fcedit")%>
									</td>

									<td colspan=2 class="ratatengah">
									<a href="PaymentVoucherDetail.Asp?fcCode=<%=rsdtlupah("VoucherCode")%>&transtype=<%=rsdtlupah("transtype")%>&txtlineno=<%=rsdtlupah("lineno")%>&xCode=<%=rsdtlupah("bankcode")%>&fcba=<%=rsdtlupah("fcba")%>&txtJenis=Edit">
									<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(1)%>" border="0">
									</a>
									<%  
									if  (rsdtlupah("rowstateid") = "Confirmed" or rsdtlupah("rowstateid") = "Planned")  then 
										txtFileReportName = "printvoucherfin"
										else 
										txtFileReportName = "printvoucherfinapprove"
									end if
									%>
									<a href="<%=txtFileReportName%>.asp?fcba=<%=fcba%>&vouchercode=<%=rsdtlupah("VoucherCode")%>&transtype=<%=rsdtlupah("transtype")%>" target="_blank">
									<img src="IMAGES/Print.gif" border="0" title="Print Preview">
									</a>
									<a href="bkm_jurnal_view.Asp?bkmno=<%=rsdtlupah("VoucherCode")%>&tglbkm=<%=rsdtlupah("datecreated")%>&xCode=<%=rsdtlupah("bankcode")%>&fcba=<%=rsdtlupah("fcba")%>&tipetran=FAC-BNK" target="_blank">
									<img src="IMAGES/view.png" target="_blank"  alt="Posting Data : <%=rsdtlupah("VoucherCode")%>" border="0">
									</a>
									</td>
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("totalamount"))
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
										<td colspan="9" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
								</tr>
							</table>
						</div>
<!-- Receive -->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table2'>
								<tr id="trheader2">									
									<th>No.Voucher</th>
									<th>Date</th>
									<th>Period</th>
									<th>Remark Header</th>
									<th>Cek No</th>
									<th>Classification</th>
									<th>Receive By</th>
									<th>Keterangan</th>
									<th>Amount</th>
									<th>Status</th>
									<th>Posting</th>
									<th >FCENTRY</th>
									<th >FCEDIT</th>
									<th>Action</th>	
								</tr>

								<% 
								
								
								'get data detail upah
								'dim sqldtlupah,rsdtlupah,no,totalamount
								sqldtlupah = " SELECT a.vouchercode,to_char(datecreated,'dd/mm/yyyy')datecreated,referenceno,cekno,classification_api.get_fcname(document_classification) class," & _ 
								             " a.remarks,b.remarks remarksheader, nvl(a.amount,0) totalamount,rowstateid,transtype,bankcode,a.fcba,nvl(b.postingflg,'0') postingflg" & _
											 " ,b.accountingperiod,b.accountingyear,b.fcentry,b.fcedit,a.lineno,nvl(c.docno,a.vouchercode) docno " & _
											 " FROM paymentvoucherdetail a " & _
											 " INNER JOIN paymentvoucher b ON (a.vouchercode=b.vouchercode AND a.fcba=b.fcba)  " & _
											 " LEFT JOIN bkmjournalhdr c ON (a.vouchercode=c.docno AND a.fcba=c.fcba) " & _
											 " where a.fcba='" & fcba & "'" & _ 
											 " AND bankcode='" & bankcode & "'" & _ 
											 " AND to_number(to_char(datecreated,'YYYY'))="& tahun & " " & _ 
											 " AND to_number(to_char(datecreated,'MM'))="& bulan & " " & _
											 " AND transtype='01-RECEIVE' " & _
											 " UNION " & _
											 " SELECT a.vouchercode,to_char(datecreated,'dd/mm/yyyy')datecreated,referenceno,cekno,classification_api.get_fcname(document_classification) class," & _ 
											 " a.remarks,b.remarks remarksheader,nvl(a.amount,0) totalamount,rowstateid,transtype,bankcode,a.fcba,nvl(b.postingflg,'0') postingflg" & _
											 " ,b.accountingperiod,b.accountingyear,b.fcentry,b.fcedit,a.lineno,nvl(c.docno,Trim(a.vouchercode)||'-INT') docno " & _
											 " FROM paymentvoucherdetail a " & _
											 " INNER JOIN paymentvoucher b ON (a.vouchercode=b.vouchercode AND a.fcba=b.fcba)  " & _
											 " INNER JOIN bkmjournalhdr c ON (Trim(a.vouchercode)||'-INT'=Trim(c.docno)) " & _
											 " where a.chargeunit = '" & fcba & "'" & _ 
											 " AND chargecode='" & bankcode & "'" & _ 
											 " AND to_number(to_char(datecreated,'YYYY'))="& tahun & " " & _ 
											 " AND to_number(to_char(datecreated,'MM'))="& bulan & " " & _
											 " AND transtype='02-PAYMENT' " & _
											 " ORDER BY datecreated DESC, vouchercode DESC "
								'response.write sqldtlupah
								set rsdtlupah = connStr.Execute(sqldtlupah)
							
																
								no = 0
								totalamount = 0
								While Not rsdtlupah.EOF 
									postingflg = rsdtlupah("postingflg")
									if postingflg = "0" then 
										postingflg="Tidak"
									else 
										postingflg="YA"
									end if
								%>
								<tr id="tru<%=no%>">
									
									<td class="ratakiri">
										<%=rsdtlupah("vouchercode")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("datecreated")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("accountingperiod")%>-<%=rsdtlupah("accountingyear")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("referenceno")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("cekno")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("class")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("remarksheader")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("remarks")%>
									</td>
									
									<td>
										<%=formatnumber(rsdtlupah("totalamount"),0)%>
									</td>
									
									<td>
										<%=rsdtlupah("rowstateid")%>
									</td>
									<td class="ratatengah">
										<%=postingflg%>
									</td>

									<td class="ratakiri">
										<%=rsdtlupah("fcentry")%>
									</td>

									<td class="ratakiri">
										<%=rsdtlupah("fcedit")%>
									</td>
									
									<td colspan=2 class="ratatengah">
									<a href="PaymentVoucherDetail.Asp?fcCode=<%=rsdtlupah("VoucherCode")%>&transtype=<%=rsdtlupah("transtype")%>&txtlineno=<%=rsdtlupah("lineno")%>&xCode=<%=rsdtlupah("bankcode")%>&fcba=<%=rsdtlupah("fcba")%>&txtJenis=Edit">
									<img src="IMAGES/edit.gif"  alt="Edit Data : <%=rsdtlupah("VoucherCode")%>" border="0">
									</a>
									<%  
									if  (rsdtlupah("rowstateid") = "Confirmed" or rsdtlupah("rowstateid") = "Planned")  then 
										txtFileReportName = "printvoucherfin"
										else 
										txtFileReportName = "printvoucherinfinapprove"
									end if
									%>
									<a href="<%=txtFileReportName%>.asp?fcba=<%=fcba%>&vouchercode=<%=rsdtlupah("VoucherCode")%>&transtype=<%=rsdtlupah("transtype")%>" target="_blank">
									<img src="IMAGES/Print.gif" border="0" title="Print Preview">
									</a>
									<a href="bkm_jurnal_view.Asp?bkmno=<%=rsdtlupah("docno")%>&tglbkm=<%=rsdtlupah("datecreated")%>&xCode=<%=rsdtlupah("bankcode")%>&fcba=<%=rsdtlupah("fcba")%>&tipetran=FAC-BNK" target="_blank">									
									<img src="IMAGES/view.png" alt="Posting Data : <%=rsdtlupah("VoucherCode")%>" border="0">
									</a>

									
									</td>
									
								</tr>
								<%

									no = no +1
									totalamount = cdbl(totalamount) + cdbl(rsdtlupah("totalamount"))
									rsdtlupah.MoveNext
									Wend
								%>
								<tr class="ratakanan">
										<td colspan="8" class="tebal">Total</td>
										<td class="tebal"><input class="ratakanan lebarupah" type="text" readonly  style="background-color:#eee;" name="amount" id="amount<%=no%>" value="<%=FormatNumber(totalamount,0)%>"/></td>
								</tr>
							</table>
						</div>
						
<!-- Cash Flow -->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data'  id="product-table" class="xthf-site">
								<tr id="trheader2">									
									<th>Date</th>
									<th>No.Voucher</th>
									<th>Paid For / Receie From</th>
									<th>Reference</th>
									<th>Remark</th>
									<th>Debet</th>
									<th>Credit</th>
									<th>Saldo</th>
								
									
								</tr>

								<% 
								
								
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,postingflg,rsCARI,nOpening

								q="SELECT nvl(sum(OPENING_BALANCE),0) as OPENING_BALANCE  FROM fac_trialbalance WHERE FCCODE ='" & txtControlJob & _
								"' AND  FCBA= '" & fcba & "'" & _
								"  AND PERIOD_MONTH = '" & bulan & "'" & _
								"  AND PERIOD_year = '" & tahun & "' "
								Set rsCARI = Server.CreateObject("ADODB.Recordset")
								rsCARI.CursorType = 3
								rsCARI.CursorLocation = 3
								' response.Write q
								response.Flush()
								rsCARI.Open q, connstr	 
								if Not rsCARI.eof Then
								nOpening=rsCARI("OPENING_BALANCE")
								else
								nOpening=0
								End if
								rsCari.Close
								set rsCari=nothing

								' Amount diminuskan
								' Karena kalau debet adalah kas bekurang
								' Kalau Kredit maka kas bertambah
								'-------------------------------- Control Job  -----------------------------------------------------------------------------------
								txtSQL1 =  	 "  SELECT p.datecreated," & _ 
											"                    p.bankcode," & _ 
											"                    p.transtype," & _ 
											"                    pay.vouchercode," & _
											"                    p.remarks orang," & _			 
											"                    pay.jobcode," & _ 
											"                    job.fcname jobname," & _ 
											"                    pay.locationtype," & _ 
											"                    pay.locationcode," & _ 
											"                    NVL (SUM (pay.amount), 0) AS debet," & _ 
											"                    0 as credit," & _ 
											"                    pay.REFERENCE," & _ 
											"                    pay.remarks" & _ 
											"               FROM paymentvoucherdetail pay " & _ 
											" 				INNER JOIN paymentvoucher p ON pay.vouchercode=p.vouchercode and pay.fcba=p.fcba " & _ 
											"				inner join jobmaster job on pay.jobcode = job.fccode " & _ 
											" 				LEFT JOIN bkmjournalhdr c ON (p.vouchercode=c.docno AND p.fcba=c.fcba) " & _
											"              WHERE pay.fcba = '" & fcba & "'" & _ 
											"                    AND p.transtype = '01-RECEIVE'" & _ 
											"                    AND p.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _ 
											"				and TO_NUMBER(TO_CHAR(p.datecreated,'mm'))='" & bulan & "' " & _ 
											"				and TO_NUMBER(TO_CHAR(p.datecreated,'yyyy'))='" & tahun & "' " & _
											" 				and (p.BANKCODE='" & bankcode & "')  "

								txtsql1=(txtsql1) &  "           GROUP BY p.datecreated," & _ 
													"                    p.bankcode," & _ 
													"                    p.transtype," & _ 
													"                    pay.vouchercode," & _
													"                    p.remarks," & _	
													"                    pay.jobcode," & _ 
													"                    job.fcname," & _ 
													"                    pay.locationtype," & _ 
													"                    pay.locationcode," & _ 
													"                    pay.REFERENCE," & _ 
													"                    pay.remarks" 

								txtSQL2 =  	 "  SELECT p.datecreated," & _ 
											"                    p.bankcode," & _ 
											"                    p.transtype," & _ 
											"                    pay.vouchercode," & _
											"                    p.remarks orang," & _			 
											"                    pay.jobcode," & _ 
											"                    job.fcname jobname," & _ 
											"                    pay.locationtype," & _ 
											"                    pay.locationcode," & _ 
											"                    NVL (SUM (pay.amount), 0) AS debet," & _ 
											"                    0 as credit," & _ 
											"                    pay.REFERENCE," & _ 
											"                    pay.remarks" & _ 
											"               FROM paymentvoucherdetail pay " & _ 
											" 				INNER JOIN paymentvoucher p ON pay.vouchercode=p.vouchercode and pay.fcba=p.fcba " & _ 
											"				inner join jobmaster job on pay.jobcode = job.fccode " & _ 
											" 				LEFT JOIN bkmjournalhdr c ON (p.vouchercode=c.docno AND p.fcba=c.fcba) " & _
											"              WHERE pay.chargeunit = '" & fcba & "'" & _ 
											"                    AND p.transtype = '02-PAYMENT'" & _ 
											"                    AND p.rowstateid in('Confirmed','AuthorizeOnProgress','Approved') " & _ 
											"				and TO_NUMBER(TO_CHAR(p.datecreated,'mm'))='" & bulan & "' " & _ 
											"				and TO_NUMBER(TO_CHAR(p.datecreated,'yyyy'))='" & tahun & "' " & _
											" 				and (pay.chargecode='" & bankcode & "')"

								txtsql2=(txtsql2) &  "           GROUP BY p.datecreated," & _ 
													"                    p.bankcode," & _ 
													"                    p.transtype," & _ 
													"                    pay.vouchercode," & _
													"                    p.remarks," & _	
													"                    pay.jobcode," & _ 
													"                    job.fcname," & _ 
													"                    pay.locationtype," & _ 
													"                    pay.locationcode," & _ 
													"                    pay.REFERENCE," & _ 
													"                    pay.remarks" 


								'-------------------------------- Jurnal ----------------------------------------------------------------------------------------

								txtSQL3 =  	 "             SELECT p.datecreated," & _ 
											"                    p.bankcode," & _ 
											"                    p.transtype," & _ 
											"                    pay.vouchercode," & _ 
											"                    p.remarks orang," & _	
											"                    pay.jobcode," & _ 
											"                    job.fcname jobname," & _ 
											"                    pay.locationtype," & _ 
											"                    pay.locationcode," & _ 
											"                    0 as debet," & _ 
											"                    NVL (SUM (pay.amount), 0) AS credit," & _ 
											"                    pay.REFERENCE," & _ 
											"                    pay.remarks" & _ 
											"               FROM paymentvoucherdetail pay " & _ 
											" 				INNER JOIN paymentvoucher p ON pay.vouchercode=p.vouchercode and pay.fcba=p.fcba " & _ 
											"				inner join jobmaster job on pay.jobcode = job.fccode " & _ 
											" 				LEFT JOIN bkmjournalhdr c ON (p.vouchercode=c.docno AND p.fcba=c.fcba) " & _
											"               WHERE pay.fcba = '" & fcba & "'" & _ 
											"                    AND p.transtype = '02-PAYMENT'" & _ 
											"                    AND p.rowstateid in('Confirmed','AuthorizeOnProgress','Approved')" & _ 
											"				and TO_NUMBER(TO_CHAR(p.datecreated,'MM'))='" & bulan & "' " & _ 
											"				and TO_NUMBER(TO_CHAR(p.datecreated,'yyyy'))='" & tahun & "' " & _
											" 				and (p.BANKCODE='" & bankcode & "')"
											
								txtsql3=(txtsql3) &  "           GROUP BY p.datecreated," & _ 
													"                    p.bankcode," & _ 
													"                    p.transtype," & _ 
													"                    pay.vouchercode," & _ 
													"                    p.remarks," & _	
													"                    pay.jobcode," & _ 
													"                    job.fcname," & _ 
													"                    pay.locationtype," & _ 
													"                    pay.locationcode," & _ 
													"                    pay.REFERENCE," & _ 
													"                    pay.remarks"  

								txtsql = "   SELECT vouchercode," & _ 
										"          datecreated," & _ 
										"          orang," & _ 
										"          bankcode," & _ 
										"          jobcode," & _ 
										"          jobname," & _ 
										"          transtype," & _ 
										"          locationtype," & _ 
										"          locationcode," & _ 
										"          reference," & _ 
										"          remarks," & _ 
										"          SUM (NVL (debet, 0)) debet," & _ 
										"          SUM (NVL (credit, 0)) credit" & _ 
										"     FROM ( "& txtsql1 &" union all "& txtsql2 &" union all "& txtsql3 &")" & _
										" GROUP BY vouchercode," & _ 
										"          datecreated," & _ 
										"          bankcode," & _ 
										"          transtype," & _ 
										"          jobcode," & _ 
										"          orang," & _ 
										"          jobname," & _ 
										"          locationtype," & _ 
										"          locationcode," & _ 
										"          reference," & _ 
										"          remarks ORDER BY datecreated, transtype, vouchercode, bankcode " 
							'	Response.Write txtsql

								Set rsdtlupah = Server.CreateObject("ADODB.Recordset")
									rsdtlupah.Open txtsql, connstr, adOpenStatic 

							v_Amount = 0
							v_Saldo = 0%>
							<tr>
								<td align="left" style="<%=cmdstyleLeft%>" colspan="7"><strong><font face="Verdana" size="1">Opening Balance.........</strong></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(nOpening,2)%></strong></td> 
								
							</tr>
							<%
								v_Amount=cdbl(to_nol(nOpening))
								v_Saldo=cdbl(to_nol(nOpening))
								ttldebet = 0
								ttlcredit = 0
								While Not rsdtlupah.EOF 
									v_Saldo=v_Saldo + cdbl(to_nol(rsdtlupah("debet"))) - cdbl(to_nol(rsdtlupah("credit")))
								%>
								<tr id="tru<%=no%>">
									<td class="ratakiri">
										<%=rsdtlupah("datecreated")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("vouchercode")%>
									</td>
									
									<td class="ratakiri">
										<%=rsdtlupah("orang")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("Reference")%>
									</td>
									<td class="ratakiri">
										<%=rsdtlupah("Remarks")%>
									</td>
									
									
									<td class="ratakanan">
										<%=FormatNumber(rsdtlupah("debet"),0)%>
									</td>
									<td class="ratakanan">
										<%=FormatNumber(rsdtlupah("credit"),0)%>
									</td>
									<td class="ratakanan">
										<%=FormatNumber(v_Saldo,0)%>
									</td>
								
									
								</tr>
								<%
									 ttldebet = ttldebet + trim(cdbl(to_nol(rsdtlupah("debet"))))
									 ttlcredit= ttlcredit +  trim(cdbl(to_nol(rsdtlupah("credit"))))
									 v_Amount=v_Amount+trim(cdbl(to_nol(rsdtlupah("debet")))) - trim(cdbl(to_nol(rsdtlupah("credit"))))
								
								rsdtlupah.MoveNext
								
								no=no+1
								Wend
								rsdtlupah.close()
								rsdtlupah=null%>
								<tr class="ratakanan"> 
										<td align="left"  colspan="5"> <font face="Verdana" size="1"><strong>Ending Balance.........</strong></td>
										<td align="Right"  <font face="Verdana" size="1"><strong><%=FormatNumber(ttldebet,0)%></strong></td>
										<td align="Right"  <font face="Verdana" size="1"><strong><%=FormatNumber(ttlcredit,0)%></strong></td>
										<td align="Right"  <font face="Verdana" size="1"><strong><%=FormatNumber(v_Amount,0)%></strong></td>
								</tr>
							</table>
						</div>				

					</div>
				</div>
			</td>
		</tr>
		<tr class="ratatengah">
			
		
			 
			
		</tr>
	</table>

</form>
</div>
</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script src="/distexcel/tableToExcel.js"></script>
	<script src="/distexcel/jquery.tableToExcel.js"></script>
<script type="text/javascript">

    function createData(pilihbkm)
	{
		var pfcba 	     = $("#businessunitbnk").val();
		var ptype 	     = $("#stockcode").val();
		
		
			if(pilihbkm=="PYV"){
				window.open('paymentvoucherdetail.Asp?txtModule=Finance%20and%20Accounting&txtTable=02-paymentvoucher&transtype=02-payment&banktype=0&bankcodeurl=' + ptype,'_blank');
			}else if(pilihbkm=="RCV")
			{
			    window.open('paymentvoucherdetail.Asp?txtModule=Finance%20and%20Accounting&txtTable=01-receivevoucher&transtype=01-receive&banktype=0&bankcodeurl=' + ptype,'_blank');
			}
		
	}
</script>

<script type="text/javascript" LANGUAGE="VBScript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

$(function() {	
	$("#jurnaldate").datepicker();
	var postingflg =  $("#POSTINGFLG").val();
	if(postingflg==0){
		$("#btnsave").show();
		$("#btnunposting").hide();
	}else {
		$("#btnunposting").show();
		$("#btnsave").hide();
	}
	
		
});

function getExport()
	{
		
		var bankcode 	= "<%=stockcode%>";
		var banknm 	= "<%=stockname%>";
		var nm = 'CashFlow_' + bankcode + ' - ' + banknm + ".xlsx";  
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "CashFlow"
            }
          });

	}


function numformat(val) {
	var nStr = val.toString();
	nStr = nStr.replace(/\,/g, "");
	var x = nStr.split('.');
	var x1 = x[0];
	var x2 = x.length > 1 ? '.' + x[1] : '';
	var rgx = /(\d+)(\d{3})/;
	while (rgx.test(x1)) {
		x1 = x1.replace(rgx, '$1' + '.' + '$2');
	}
	return x1 + x2;
}


function actionBtn()
	{
		var bkmfrom = $("#bkmfrom").val();
		var bkmto 	= $("#bkmto").val();
		var jumlinv = parseInt($('#jumldata').val());
		
		if (bkmfrom != "") { 
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?bkmfrom=" + bkmfrom + "&bkmto=" + bkmto,function(data) {
				$('#bodytable').empty();
				var json  = JSON.parse(data);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					var x = json.results[i].fccode
					//alert(x);
					var docketlist = '';
						docketlist += '<tr id="tru'+i+'">';
						docketlist += '<td align="left">' + json.results[i].bkmno + '</td>';
						docketlist += '<td align="left">' + json.results[i].createdate + '</td>';
						docketlist += '<td align="right">' + numformat(json.results[i].totalbkmupah) + '</td>';
						docketlist += '<td align="right">' + json.results[i].totalbkmmaterial + '</td>';				
						docketlist += '<td align="right">' + json.results[i].totalbkmvehicle + '</td>';
						docketlist += '<td align="right">' + json.results[i].totalbkmwip + '</td>';
						docketlist += '<td align="center"><a target=_blank href="bkm_operasional_edit.Asp?bkmno='+ json.results[i].bkmno + '"><img src="IMAGES/edit.gif" alt="Edit Data border="0"></a></td>';
						docketlist += '<td align="center"><a target=_blank href="bkm_jurnal_input.Asp?tglbkm='+ json.results[i].createdate + '"><img src="IMAGES/view.png" alt="View Data border="0"></a></td>';
						docketlist += '</tr>';
						$('#bodytable').append(docketlist);
					no++;
					$("#jumldata").val(jumlview);
					
				}
				$('#loadingjudul').hide();
				$('#loading').hide();
				
			});
		 
		} else {
			$('#bodytable').empty();
		}
	}




</script>