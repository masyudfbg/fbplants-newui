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
<title>ePlaS - Print STO</title>
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
 invoicecode = Request("stono")
 businessunit = Request("fcba")

%>

<%

	Dim sqlview,rsview
	Dim txtdate,txtstono,txtpocode,txttransportter

	sqlview = "SELECT sth.company,sth.fcba,sth.stodate,sth.stono,sth.pocode,tr.fcname "  & _
			"  FROM stocktransferorderhdr sth"  & _
			"  INNER JOIN transporter tr on (tr.fccode=sth.TRANSPORTERCODE and tr.FCBA=sth.FCBA) AND sth.stono='" & invoicecode & "'"
	set rsview = connStr.Execute(sqlview)

   if not rsview.eof then
   		   txtdate 			= rsview("stodate")
		   txtstono 		= rsview("stono")
		   txtpocode 		= rsview("pocode")
		   txttransportter 	= rsview("fcname")
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
				<div class="ratatengah fontcourir" style="font-size:12px;"><strong>STOCK TRANSFER ORDER</strong></div>
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
				<td>Transporter</td>
				<td>:</td>
				<td><%=txttransportter%></td>
				<td>No STO</td>
				<td>:</td>
				<td><%=txtstono%></td>
			</tr>
			<tr>
				<td>No. PO</td>
				<td>:</td>
				<td><%=txtpocode%></td>
				<td>Tgl STO</td>
				<td>:</td>
				<td><%=txtdate%></td>
			</tr>
		</table>
		
		<table border="0" width="100%"  style="border-top:1px #000000 solid; font-size:12px;line-height:20px;">
			<thead>
				<tr style="border-bottom:1px #000000 solid; font-size:12px;">
					<th class="ratatengah font10px fontcourir">No.</th>
					<th class="font10px fontcourir ratakiri">Item Code</th>
					<th class="font10px fontcourir ratakiri">Item Name</th>
					<th class="fontcourir font10px ratatengah" >Qty PO</th>					
					<th class="fontcourir font10px ratakanan">Price</th>					
					<th class="fontcourir font10px ratakanan">Amount</th>					
				</tr>
			 </thead>
			 	<%
				 Dim sqldtl,rsdtl,arr_dtl,arr_dtl_eof,xyz
				 Dim txttotalqty,txttotalall
				 Dim txtitemcode,txtitemname,txtquantity,txtprice,txtamountdtl,no,tes
				 sqldtl = "select d.itemcode,s.fcname,d.quantityorder,d.unitprice,d.amount from stocktransferorderdtl d,stock s " & _ 
				 " where d.itemcode=s.fccode and d.fcba=s.fcba and d.fcba='" & businessunit & "' AND d.stono='" & invoicecode & "'"
				 Set rsdtl = Server.CreateObject("ADODB.Recordset")
				rsdtl.Open sqldtl, connstr
				
				 no = 1	
				txttotalqty = 0
				txttotalall = 0
				 if Not rsdtl.eof then
				While Not rsdtl.EOF
						txtitemcode = rsdtl("itemcode")
						txtitemname = rsdtl("fcname")
						txtquantity = rsdtl("quantityorder")
						txtprice = rsdtl("unitprice")
						txtamountdtl = rsdtl("amount")
						
				%>
				<tr class="fontcourir">
					<td class="ratatengah fontcourir"><%=no%></td>
					<td class="fontcourir ratakiri"><%=txtitemcode%></td>
					<td class="fontcourir"><%=txtitemname%></td>
					<td class="ratakanan fontcourir ratatengah"><%=formatnumber(cdbl(txtquantity),0)%></td>
					<td class="ratakanan fontcourir"><%=formatnumber(cdbl(txtprice),0)%></td>
					<td class="ratakanan fontcourir"><%=formatnumber(cdbl(txtamountdtl),0)%></td>
				</tr>
				<%
						no=no+1
						txttotalqty = txttotalqty+cdbl(txtquantity)
						txttotalall = txttotalall+cdbl(txtamountdtl)
					
			
						rsdtl.MoveNext
						no = no+1
						Wend
					end if
					rsdtl.Close
					set rsdtl = Nothing%>
				<tr class="fontcourir" style="border-top:1px #000000 solid;">
					<td colspan="5" class="ratakanan tebal">Subtotal</td>
					
					<td class="ratakanan tebal"><%=formatnumber(cdbl(txttotalall),0)%></td>
				</tr>
			<!-- 	<tr class="fontcourir" style="">
					<td colspan="5" class="ratakanan tebal">PPN (10%)</td>
				
					<td class="ratakanan tebal"><%=formatnumber(cdbl(txtttlppn),0)%></td>
				</tr>
				<tr class="fontcourir" style="">
					<td colspan="5" class="ratakanan tebal">Grand Total</td>
				
					<td class="ratakanan tebal"><%=formatnumber(cdbl(txttotalamount),0)%></td>
				</tr>-->
			<tr>
				<td colspan="3" class="ratakiri" style="padding:0px;">
				<%
					Function Val( myString )
					' Val Function for VBScript (aka ParseInt Function in VBScript).
					' By Denis St-Pierre.
					' Natively VBScript has no function to extract numbers from a string.
					' Based shamelessly on MS' Helpfile example on RegExp object.
					' CAVEAT: Returns only the *last* match found
					'         (or, with objRE.Global = False, only the *first* match)
					
						Dim colMatches, objMatch, objRE, strPattern
					
						' Default if no numbers are found
						Val = 0
					
						strPattern = "[-+0-9]+"       ' Numbers positive and negative; use
													' "ˆ[-+0-9]+" to emulate Rexx' Value()
													' function, which returns 0 unless the
													' string starts with a number or sign.
						Set objRE = New RegExp        ' Create regular expression object.
						objRE.Pattern    = strPattern ' Set pattern.
						objRE.IgnoreCase = True       ' Set case insensitivity.
						objRE.Global     = True       ' Set global applicability:
													'   True  => return last match only,
													'   False => return first match only.
						Set colMatches = objRE.Execute( myString )  ' Execute search.
						For Each objMatch In colMatches             ' Iterate Matches collection.
							Val = objMatch.Value
						Next
						Set objRE= Nothing
					End Function
					
					Function Terbilang(Angka)
					Dim n1
						Dim nangka, Strjmlhuruf, intpecahan  , strpecahan, urai, Bil1, strtot, bil2
						Dim x, Y, z, i  
					
						For i = 1 To Len(Angka)
							If Not Mid(Angka, i, 1) = "." Then
								nangka = nangka + Mid(Angka, i, 1)
							End If
						Next
					
						If Len(LTrim(nangka)) = 0 Then
							terbit = "Nol Rupiah"
							Exit Function
						End If
						Strjmlhuruf = LTrim(RTrim((nangka)))
						intpecahan = Val(Right(Mid(Angka, 15, 2), 2))
						If (intpecahan = 0) Then
							strpecahan = ""
						End If
						x = 0
						Y = 0
						urai = ""
					
						While (x < Len(Strjmlhuruf))
							x = x + 1
							strtot = Mid(Strjmlhuruf, x, 1)
							Y = Y + Val(strtot)
							z = Len(Strjmlhuruf) - x + 1
					
							Select Case Val(strtot)
							Case 1
								If (z = 1 Or z = 7 Or z = 10 Or z = 13) Then
									Bil1 = "Satu "
								ElseIf (z = 4) Then
									If (x = 1) Then
										Bil1 = "Se"
									Else
										Bil1 = "Satu "
									End If
								ElseIf (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
									x = x + 1
									strtot = Mid(Strjmlhuruf, x, 1)
									z = Len(Strjmlhuruf) - x + 1
									bil2 = ""
					
								Select Case Val(strtot)
								Case 0
									Bil1 = "Sepuluh "
								Case 1
									Bil1 = "Sebelas "
								Case 2
									Bil1 = "Dua Belas "
								Case 3
									Bil1 = "Tiga Belas "
								Case 4
									Bil1 = "Empat Belas "
								Case 5
									Bil1 = "Lima Belas "
								Case 6
									Bil1 = "Enam Belas "
								Case 7
									Bil1 = "Tujuh Belas "
								Case 8
									Bil1 = "Delapan Belas "
								Case 9
									Bil1 = "Sembilan Belas "
								End Select
								Else
									Bil1 = "Se"
								End If
					
							Case 2
								Bil1 = "Dua "
							Case 3
								Bil1 = "Tiga "
							Case 4
								Bil1 = "Empat "
							Case 5
								Bil1 = "Lima "
							Case 6
								Bil1 = "Enam "
							Case 7
								Bil1 = "Tujuh "
							Case 8
								Bil1 = "Delapan "
							Case 9
								Bil1 = "Sembilan "
							Case Else
								Bil1 = ""
							End Select
					
						If (Val(strtot) > 0) Then
							If (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
								bil2 = "Puluh "
							ElseIf (z = 3 Or z = 6 Or z = 9 Or z = 12 Or z = 15) Then
								bil2 = "Ratus "
							Else
								bil2 = ""
							End If
						Else
							bil2 = ""
						End If
						If (Y > 0) Then
							Select Case z
							Case 4
								bil2 = bil2 + "Ribu "
								Y = 0
							Case 7
								bil2 = bil2 + "Juta "
								Y = 0
							Case 10
								bil2 = bil2 + "Milyar "
								Y = 0
							Case 13
								bil2 = bil2 + "Trilyun "
								Y = 0
							End Select
						End If
						urai = urai + Bil1 + bil2
					Wend
					
					urai = urai + strpecahan
					Terbilang = urai & ""
					Terbilang = UCase(Terbilang)
					n1 = UCase(Left(Terbilang, 1))
					Terbilang = n1 & Mid(Terbilang, 2, Len(Terbilang) - 1)
					End Function
			
				
				%>
				<div class="title_terbilang fontcourir">
				Terbilang:
				</div>
				<div class="terbilang fontcourir" >
				<%  
					Dim txtSQL,rsCats,v_terbilang
					txtSQL= "select upper(replace(terbilang_indo("& txttotalall &"),'-',' ')) terbilang from dual"
					Set rsCats = Server.CreateObject("ADODB.Recordset")
					rsCats.Open txtSQL, connstr
					if Not rsCats.eof then
						v_terbilang =  rsCats("terbilang") 
					end if
					rsCats.close
					set rsCats = Nothing
				%>
				<%=v_terbilang%> RUPIAH
				</div>
				</td>
			</tr>
		</table>
	
	<span class="jaraktanda font10px fontcourir">Dibuat Oleh,</span><span class="jaraktanda font10px fontcourir">Diperiksa Oleh,</span><span class="jaraktanda font10px fontcourir">Disetujui Oleh,</span><br><br><br>

	<span class="jaraktandabawah font10px fontcourir">--------------</span>				<span class="jaraktandabawah font10px fontcourir">----------------</span> <span class="jaraktanda font10px fontcourir">----------------</span>
	
	<br/>
</body>
</html>
