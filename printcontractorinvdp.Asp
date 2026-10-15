<%'Option Explicit
%>
<%'On Error GOTO CmdErrorlog 
%>
<%response.buffer=True%>
<%

%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Print Contractor Invoice DP</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
        <script language="JavaScript">
        xAddEventListener(window, 'load',
          function() {
            new xTableHeaderFixed('xthf-site', window);
          }, false
        );
        </script>
		<script language="JavaScript" src="/js/jquery.printPage.js"></script> 
	<script>
					function toggle(id){
						var e = document.getElementById(id);
						if(e.style.display == 'inline-block'){
							e.style.display = 'none';
						}else if(e.style.display == 'none'){
							e.style.display = 'inline-block';
						}
					}
	</script>
	<script>
	 function printDiv(divName) {
    		window.print();
	}
</script>

<style type="text/css">
@media all {
#product-table td{padding:5px;padding-bottom: 0; border-top:0px;border-bottom:0px;
    padding-top: 0;}
.page-break	{ display: block; }
a.button_print { display:block; }	
#wrapper{background: #e2e2e2; /* Old browsers */
background: -moz-linear-gradient(-45deg,  #e2e2e2 0%, #dbdbdb 50%, #d1d1d1 51%, #fefefe 100%); /* FF3.6+ */
background: -webkit-gradient(linear, left top, right bottom, color-stop(0%,#e2e2e2), color-stop(50%,#dbdbdb), color-stop(51%,#d1d1d1), color-stop(100%,#fefefe)); /* Chrome,Safari4+ */
background: -webkit-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Chrome10+,Safari5.1+ */
background: -o-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Opera 11.10+ */
background: -ms-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* IE10+ */
background: linear-gradient(135deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* W3C */
filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#e2e2e2', endColorstr='#fefefe',GradientType=1 ); /* IE6-9 fallback on horizontal gradient */
}
a.button_print {display: inline-block;}
.letter { 
background-color:#fff; 
box-shadow:0px 0px 4px rgba(0,0,0,0.4); 
border-radius:4px;
margin: auto;
width: 98%;
padding-top:30px;
padding-bottom:30px;
padding-left:1%;
padding-right:1%;
}
#header-wrapper {display:block}
.navi_to_print {
display:block;
margin: auto;
width: 98%;
padding-top:30px;
}
}
@media print
{
table td{padding:3px;}
#header-wrapper {display:none}
.button_print {display: none;}
.printhide{display:none;}
.page-break	{ display: block; page-break-before: always; }
.product-table{width: 100%;}
#wrapper{background:#fff;padding:none;}
#product-table th {font-size:14px; }
#product-table td {font-size:14px; padding:2px;padding-bottom: 0;
    padding-top: 0;order-top:0px;border-bottom:0px;}
a.button_print {display:none;}
.letter { 
background-color:#fff; 
box-shadow:none; 
border-radius:0px;
width: 100%;
margin: none;
padding-top:0px;
padding-bottom:0px;
padding-left:0px;
padding-right:0px;
}
.navi_to_print {
display:none;
margin: auto;
width: 100%;
padding-top:30px;
}
}
</style>

</head>

<body>
<%
 Dim invoicecode
 invoicecode = Request("invoicecode")
 businessunit = Request("fcba")

%>

<%

	Dim sqlview,rsview
	Dim txtinvoicedate,txtinvoiceno,txtpocode,txtsupplier,txtamount,txtttlppn,txtothercharge,txttotalamount 
	sqlview = "select h.invoicecode,h.transdate,h.agreementcode,h.contractorcode,c.fcname,h.qty,h.rate,h.amount,h.dp,h.remain " & _
" from CONTRACTORINVOICEDP h,CONTRACTOR c " & _
" where h.CONTRACTORCODE=c.FCCODE and h.FCBA = c.FCBA and h.dp-h.payment>0 " & _
" and h.invoicecode='" & invoicecode & "' " & _
" and h.fcba='" & businessunit & "'"
	set rsview = connStr.Execute(sqlview)

   if not rsview.eof then
   		txtinvoicedate 		= rsview("transdate")
		   txtinvoiceno 	= rsview("invoicecode")
		   agreementcode 	= rsview("agreementcode")
		   fcname 			= rsview("fcname")
		   qty 				= rsview("qty")
		   rate 			= rsview("rate")
		   amount 			= rsview("amount")
		   dp 				= rsview("dp")
		   remain 			= rsview("remain")
   End if
   rsview.close
   set rsview = Nothing


   
%>

<div id="wrapper">
	<div class="navi_to_print">
	<input type="button" value="Print" onclick="printDiv('note')"/>
</div>

<div class="letter">
	<div id="note">
		
		<header id="info_note" >
			<div class="posisibaris ratakiri fontcourir" style="font-size:10px; ">PT   : <%=session("company")%>  -  Unit : <%=session("FCBA")%> - <%=Session("BA")%> <span class="ratakanan" style="padding-left: 450px;"><%=bankperusahaan%></span></div>
				<div class="ratatengah fontcourir" style="font-size:12px;"><strong>INVOICE CONTRACTOR DP</strong></div>
		</header><br>
		<div class="clear"></div>
	        <style>
			.group_checkbox {margin-bottom:10px;margin-top:25px;}
			.group_checkbox ul {width:100%;}
			.group_checkbox ul li { 
				list-style-type:none;
				float:left;
			}
			.group_checkbox ul li p{
				font-weight:bold;
				font-size:12px;
				float:left;
				display:initial;
			}
			.group_checkbox ul li span{
				float:left;
				margin-top:0px;
				margin-left:10px;
			}
			.group_checkbox ul li span img{
				width:21px;
			}
			.ratakanan{
				text-align:right
			}
			.ratatengah{
				text-align:center
			}
			.ratakiri{
					text-align: left;
				}
			.posisibaris{
				padding: 1px;
			}
			.tebal{
				font-weight: bold;
			}
			.jaraktanda{
				padding-right: 55px;
			}

			.jaraktandabawah{
				padding-right: 45px;
			}
			.font10px{
				font-size:12px;
			}
			.rowtinggi{
				height: 3px;
			}
			
			
		</style>

		<table border="0" width="100%" style="font-size:12px; line-height:18px;" class="fontcourir">
			
			<tr>
				<td>Contractor</td>
				<td>:</td>
				<td><%=fcname%></td>
				<td>No Invoice</td>
				<td>:</td>
				<td><%=txtinvoiceno%></td>
			</tr>
			<tr>
				<td>Agreement Code</td>
				<td>:</td>
				<td><%=agreementcode%></td>
				<td>Tgl Invoice</td>
				<td>:</td>
				<td><%=day(txtinvoicedate) & " " & monthname(month(txtinvoicedate),true) & " " & year(txtinvoicedate)%></td>
			</tr>
			<tr>
				<td>Qty</td>
				<td>:</td>
				<td><%=qty%></td>
				<td>Rate</td>
				<td>:</td>
				<td><%=FormatNumber(rate,2)%></td>
			</tr>
			<tr>
				<td></td>
				<td></td>
				<td></td>
				
				<td>Amount</td>
				<td>:</td>
				<td><%=FormatNumber(amount,2)%></td>
				
			</tr>
			<tr>
				<td></td>
				<td></td>
				<td></td>
				
				<td>DP</td>
				<td>:</td>
				<td><%=FormatNumber(dp,2)%></td>
				
			</tr>
			<tr>
				<td></td>
				<td></td>
				<td></td>
				
				<td>Remain</td>
				<td>:</td>
				<td><%=FormatNumber(remain,2)%></td>
				
			</tr>

		</table>
		
		
	<span class="jaraktanda font10px fontcourir">Dibuat Oleh,</span><span class="jaraktanda font10px fontcourir">Diperiksa Oleh,</span><span class="jaraktanda font10px fontcourir">Disetujui Oleh,</span><br><br><br>

	<span class="jaraktandabawah font10px fontcourir">--------------</span>				<span class="jaraktandabawah font10px fontcourir">----------------</span> <span class="jaraktanda font10px fontcourir">----------------</span>
	
	<br/>
</body>
</html>
