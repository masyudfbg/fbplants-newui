<%'Option Explicit
%>
<%'On Error GOTO CmdErrorlog 
%>
<%response.buffer=True%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<%
 	Dim txtNoinvoice,txtnoref,txtnogiro,txtdate,txtpaidfor,q,i
	Dim rsPaymentfin
 	


    Set rsPaymentfin = Server.CreateObject("ADODB.Recordset")
 	q = "select vouchercode,referenceno,cekno,datecreated,remarks " & _
	" from paymentvoucher " & _
	" where fcba='" & session("FCBA") & "' AND vouchercode='" & request("vouchercode") & "' "
   'response.Write q

  rsPaymentfin.Open q, connstr

   if not rsPaymentfin.eof then
      txtNoinvoice 		= rsPaymentfin("vouchercode")
	  txtnoref 			= rsPaymentfin("referenceno")
	  txtnogiro 		= rsPaymentfin("cekno")
	  txtdate 			= rsPaymentfin("datecreated")
	  txtpaidfor 		= rsPaymentfin("remarks")
   End if
   rsPaymentfin.close
   set rsPaymentfin = Nothing
   
%>

<style type="text/css">
		.ratakanan{
				text-align:right
			}
			.ratatengah{
				text-align:center
			}
			.ratakiri{
				text-align:left
			}
			.posisibaris{
				padding: 5px;
			}
			.tebal{
				font-weight: bold;
			}
			.jarakjudulkanan{
				padding-right: 105px;
			}

			.jaraktandabawah{
				padding-right: 60px;
			}
			.jarakjobcode{
				width: 100px;
			}
			.jarakjobname{
				width: 180px;
			}
			.jarakremark{
				width: 400px;
			}
			.jarakamount{
				width: 120px;
			}
			.jarakno{
				width: 20;
			}
</style>
</head>
PT. <%=session("company")%> - Unit. <%=session("FCBA")%><br><br>
						<center><strong>Permohonan Bank Keluar / Klaim</strong></center><br><br>
<div class="ratakiri">Dibayarkan kepada	: <%=txtpaidfor%></div>										<div class="ratakanan jarakjudulkanan">No. Bukti     : <%=txtNoinvoice%></div>
<div class="ratakiri">No. Rek</div>