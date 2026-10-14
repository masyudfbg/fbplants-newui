<%'Option Explicit 
'On Error GOTO 0 
%>
<%response.buffer=True%>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Print PR</title>
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
	 //var printContents = document.getElementById(divName).innerHTML;
     //var originalContents = document.body.innerHTML;

     //document.body.innerHTML = printContents;

     window.print();

     //document.body.innerHTML = originalContents;
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
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
 <%
 Dim txtNoPR,txtDatePR,txtRemarks,q,nCountofField,no,i
 Dim Remarks,txtmanagedid,txtStatus,txtHeaderRemarks
 dim stock

   Dim rsPO,intRecordsPerPage,cColor,txtSupplier
   set rsPO = server.CreateObject("Adodb.recordset")
   q="select PRCODE,PRDATE,REMARKS,Rowstateid,MANAGEDID from purchaserequest  " & _
     " where PRCode='" & request("PrCode") & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsPO = Server.CreateObject("ADODB.Recordset")
   rsPO.Open q, connstr
   if not rsPO.eof then
      txtNoPR = rsPO("PRCODE")
	  txtDatePR = rsPO("PRDATE")
	  txtHeaderRemarks=rsPO("REMARKS")
	  txtmanagedid=rsPO("MANAGEDID")
	  txtStatus = rsPO("Rowstateid")
   End if
   rsPO.close
   set rsPO = Nothing %>
   
   <%
   Dim rsPurchaseOrder,nQty,nAmount,qprd
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   qprd = " select stockgroup_api.Get_ControlJob(stock_api.get_StockGroup(itemcode, prd.fcba), prd.fcba) CONTROL_JOB" & _
          "        , prd.itemcode" & _
          "        , sum(a.quantity) stock" & _
          "        , stock_api.get_stockName(prd.itemcode, prd.fcba) stockname" & _
          "        , prd.requestedqty" & _
		  "        , prd.approvedqty" & _
          "        , prd.unitofmeasurement" & _
          "        , prd.description" & _
		  "        , prd.lineid" & _
		  " from purchaserequestdetails prd" & _
		  " LEFT JOIN stockonhand a ON (prd.itemcode = a.stockcode and prd.fcba = a.fcba)" & _
		  " where prd.PrCode = '" & Request("PrCode") & "'" & _
		  "       and prd.fcBA='" & session("fcBA") & "'" & _
          " group by stockgroup_api.Get_ControlJob(stock_api.get_StockGroup(itemcode, prd.fcba), prd.fcba)" & _
          "          , prd.itemcode" & _
          "          , stock_api.get_stockName(prd.itemcode, prd.fcba)" & _
          "          , prd.requestedqty" & _
		  "          , prd.approvedqty" & _
          "          , prd.unitofmeasurement" & _
          "          , prd.description" & _
          "          , prd.lineid" & _
		  " order by prd.lineid"
   'response.Write qprd	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open qprd, connstr%>
<div id="wrapper">
	<div class="navi_to_print">
	<input type="button" value="Print" onclick="printDiv('note')"/>
</div>

<div class="letter">
	<div id="note">
		<style>
		@media all {
		td{padding:5px;}
		table.PAYSLIP td{padding:5px;}
		.page-break hr {display:block;border: 1px dashed red;}
		.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
		#note {margin-left:auto;padding:5px;}
		#info_note .judul { text-align:center; font-size:20px;}
		#info_note .left { float: left; width:35%;}
		#info_note .right {  float: right; width:35%;}
		#info_note .point { font-size:12px; font-weight:bold; width:100%;  margin-top:5px;}
		#info_note .point span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point .name { width:150px; float:left;}
		#info_note .point .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}

		#info_note .point_right { font-size:12px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
		#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point_right .name { width:100px; float:left;}
		#info_note .point_right .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}
		.clear { clear:both; content:'';}
		table .long {width:400px;}
		table .small {width:40px;}
		table th span { font-style:italic; font-weight:normal;}
		.line { height:1px; background:#000; content:' ';margin: auto;}
		.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
		#product-table th { background-color:#fff; color:#000;border:1px #000000 solid ;}
		#product-table td { border-color:#000; color:#000;padding-bottom: 0;order-top:0px;border-bottom:0px;
    padding-top: 0;}
		#product-table td.none { border: none;}
		#footer_note .block { font-size:10px; float:left; width:19%; height:150px; background-color:#fff; margin-left:1%; text-align:center;}
		.tembusan{font-size:10px;}
		input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
		hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
		hr{ margin-top: 5px;}
		}
		@media print
		{
		td{padding:3px;}
		table.PAYSLIP td{padding:3px;}
		.page-break hr {border:none;display: none;}
		.PAYSLIP td {font-family:verdana,sans-serif; font-size:7px; padding:0px;}
		#note {margin-left:auto;padding:0px;}
		#info_note .judul { text-align:center; font-size:20px;}
		#info_note .left { float: left; width:35%;}
		#info_note .right {  float: right; width:35%;}
		#info_note .point { font-size:14px; font-weight:bold; width:100%;  margin-top:5px;}
		#info_note .point span { font-size:12px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point .name { width:150px; float:left;}
		#info_note .point .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}

		#info_note .point_right { font-size:14px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
		#info_note .point_right span { font-size:12px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point_right .name { width:100px; float:left;}
		#info_note .point_right .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}
		.clear { clear:both; content:'';}
		table .long {width:400px;}
		table .small {width:40px;}
		table th span { font-style:italic; font-weight:normal;}
		.line { height:1px; background:#000; content:' ';margin: auto;}
		.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
		#product-table th { background-color:#fff; color:#000; font-size:14;}
		#product-table td { border-color:#000; color:#000;padding-bottom: 0;order-top:0px;border-bottom:0px;
    padding-top: 0;}
		#product-table td.none { border: none;}
		#footer_note .block { font-size:14px; float:left; width:19%; height:150px; background-color:#fff; margin-left:1%; text-align:center;}
		.tembusan{font-size:12px;}
		input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
		hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
		hr{ margin-top: 5px;}
		}
		</style>
		
		<header id="info_note" >
				<div class="judul"><strong>PURCHASE REQUISITION</strong></div>
					<table width="100%">
					<!-- Left -->
					<td valign="top">
					<div class="point">
							<div class="name">
							PT<span>Company</span>
							</div>
							<div class="isi">
							: <%=session("company")%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point">
							<div class="name">
							Kebun/Pabrik<span>Business Unit</span>
							</div>
							<div class="isi">
							: <%=session("FCBA")%> - <%=Session("BA")%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point">
							<div class="name">
							Ditujukan Pada<span>Mananged By</span>
							</div>
							<div class="isi">
							: <%=txtmanagedid%>
							</div>
						</div>
						<div class="clear"></div>
					<!-- right -->
					<td width="400px" valign="top">			
						<div class="point_right">
							<div class="name">
							Nomor<span>Number</span>
							</div>
							<div class="isi">
							: <%=txtNoPR%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point_right">
							<div class="name">
							Tanggal<span>Date</span>
							</div>
							<div class="isi">
							: <%=txtDatePR%>
							</div>
						</div>
						<div class="clear"></div>
						<div class="point_right">
							<div class="name">
							Status<span>Status</span>
							</div>
							<div class="isi">
							: <%=txtStatus%>
							</div>
						</div>
					</td>
				</table>
		</header>
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
		</style>
		<div class="group_checkbox" >
			<ul>
			<% 
			   Dim jbcode,txtjbcode,txtjbname,widthblock,jml_tabl,ttl_item,checked
			   set jbcode = server.CreateObject("Adodb.recordset")
			   
			   q="select fccode,fcname from job where fcba = '" & session("fcBA") & "' and  substr(fccode,1,3) = '121'"
			   'response.Write q	
			   jbcode.Open q, connstr 

			   if Not jbcode.eof then
				  While Not jbcode.EOF
					set ttl = server.CreateObject("Adodb.recordset")
					checkedSTR = "select count(*) TOTAL from (" & qprd & ") WHERE CONTROL_JOB = '" & jbcode("fccode") & "'"
					ttl.Open checkedSTR, connstr
					if Not ttl.eof then
						ttl_item = ttl ("TOTAL")
					end if
					ttl.Close
					set ttl = Nothing
					
					IF CDbl (ttl_item) > 0 then
						checked = "_checked"
						ELSE
						checked = ""
					end if
				  %>

				<li class="point_checkbox" style="margin-right:20px;">
					<p><%=jbcode("fcname")%></p>
					<span><img src="IMAGES/check_box<%=checked%>.PNG"><span>
				</li>

				 <% 
				
					  jbcode.MoveNext
				  Wend
			   else
				  response.Write("")	  
			   end if 
			   jbcode.Close
			   set jbcode = Nothing %> 
				<div class="clear"></div>
			</ul>
			<div class="clear"></div>
		</div>
		<table border="0" width="100%" id="product-table" style="border-bottom:1px #000000 solid;">
			<thead>
				<tr>
					<th class="small" style="border-bottom:1px #000000 solid;">
					Nomor<div class="line"></div><span>Item</span>
					</th>
					<th>Stock Code</th>
					<th class="long">Description</th>
					<th style="font-size:10px">Satuan<div class="line"></div><span>UOM</span></th>
					<th style="font-size:10px">Budget setahun<div class="line"></div><span>Qty</span></th>
					<th style="font-size:10px">Diperlukan<div class="line"></div><span>Qty</span></th>
					<th style="font-size:10px">Stock<div class="line"></div><span>Qty</span></th>
					<th style="font-size:10px">Diminta<div class="line"></div><span>Qty</th>
					<th style="font-size:10px">Disetujui<div class="line"></div><span>Qty</th>
					
					<!--<th colspan="3">Pembebanan</span></th>-->
					<th >Keterangan<div class="line"></div><span>Remark</span></th>
				</tr>
			 </thead>
			 <tbody>
				

			  <%
			  
			   if Not rsPurchaseOrder.eof then
					
				  nCountofField=rsPurchaseOrder.fields.count-1
				  no=1
				  nQty=0
				  nAmount=0 %>
				  <tr>
					  <td style="height:5px;"></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					</tr>
				  
				  <%
				  While Not rsPurchaseOrder.EOF
					if len(trim(rsPurchaseOrder("stock"))) > 0 then
						stock = rsPurchaseOrder("stock")
					else
						stock = "0"
					end if
					  'nQty = nQty+cdbl(rsPurchaseOrder("Quantity"))
					  'nAmount = nAmount+cdbl(rsPurchaseOrder("Amount"))
					  if no  mod 2 = 1 then
						 cColor = "#EEEEEE"
					  else
						 cColor = "#E5E5E5"
					  end if%>
					  
				 <tr>
					  <td align="center" style="<%=cmdstyleLeft%>"><%=NO%>&nbsp;</td>
					  <td align="center"  style="<%=cmdstyle%>" title="ITEMCODE" valign="middle"><%=rsPurchaseOrder("ITEMCODE")%>&nbsp;</td>
					  <td align="left"  style="<%=cmdstyle%>" title="stockname" valign="middle"><%=rsPurchaseOrder("stockname")%>&nbsp;</td>
					  <td align="center"  style="width:50px;<%=cmdstyle%>" title="unitofmeasurement" valign="middle"><%=rsPurchaseOrder("unitofmeasurement")%>&nbsp;</td>
					  <td align="CENTER"  style="width:50px;<%=cmdstyle%>" title="Budget setahun" valign="top"><font face="Verdana"><input class="isian" name ="budget_setahun" type="text" style="display:block;font-size:12px;" value=""></td>
					  <td align="CENTER"  style="width:50px;<%=cmdstyle%>" title="Budget setahun" valign="top"><font face="Verdana"><input class="isian" name ="budget_setahun" type="text" style="display:block;font-size:12px;" value=""></td>
					  <td align="CENTER"  style="width:50px;<%=cmdstyle%>" title="Budget setahun" valign="middle"><font face="Verdana"><%=stock%></td>
					  <td align="RIGHT"  style="width:50px;<%=cmdstyle%> " title="requestedqty" valign="middle"><%=FormatNumber(rsPurchaseOrder("requestedqty"),0)%>&nbsp;</td>
					  <td align="RIGHT"  style="width:50px;<%=cmdstyle%> " title="approvedqty" valign="middle"><%=FormatNumber(rsPurchaseOrder("approvedqty"),0)%>&nbsp;</td>
					  <td align="left"  style="<%=cmdstyle%>" title="description" valign="middle"><%=rsPurchaseOrder("description")%>&nbsp;</td>
				 </tr>
				
				  <%
					  rsPurchaseOrder.MoveNext
					  intRecordsPerPage = intRecordsPerPage - 1
					  no=no+1
				  Wend  %>
					<tr>
					  <td style="height:5px;"></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					</tr>  
				  
			<%	  
			   else
				  response.Write("<tr><td colspan='9'>Data is empty</td></tr>")	  
			   end if 
			   rsPurchaseOrder.Close
			   set rsPurchaseOrder = Nothing%> 
			    
			   </tbody>
		   </table>
	<div class="clear"></div>
    <div id="footer_note">
		<div class="block">Approved by,
        <br/><br/><br/><br/><br/>
        <input class="isian" type="text" value="" style="font-size:12px;">
        <hr/>
        Head Departement
        </div>
		<div class="block" >Approved by,
        <br/><br/><br/><br/><br/>
        <input class="isian" type="text" value="" style="font-size:12px;">
        <hr/>
        Budget Control HO
        </div>
        
        <div class="block" style="float:right;">Request by,
        <br/><br/><br/><br/><br/>
		<input class="isian" type="text" value="<%=Session("FCMANAGER")%>" readonly='True' style="font-size:12px;">
        <hr/>
        Manager
        </div>
        <div class="block" style="float:right;">Approved by,
        <br/><br/><br/><br/><br/>
        <input class="isian" type="text" value="" style="font-size:12px;">
        <hr/>
		<%
		dim gm,estate
		 set gm = server.CreateObject("Adodb.recordset")
			   
			  groupmanager = "select fctype from businessunit where fccode = '" & SESSION("FCBA")& "'"
			   'response.Write q	
			   gm.Open groupmanager, connstr 
	
				
			   if Not gm.eof then
			   estate = gm("fctype")
			   end if
			gm.Close
			set gm = Nothing
	if 	LEN(TRIM(estate)) > 0	then
		if estate <> "E" then 	
		%>
        Group Mill Manager
		<% else %>
		
		Group Estate Manager
		
		<% end if
	end if
		%>
        </div>
		<div class="block" style="float:right;">Acknowledge by,
        <br/><br/><br/><br/><br/>
		<input class="isian" type="text" value="" style="font-size:12px;">
        <hr/>
        Regional Head
        </div>
    </div>
	<div class="clear"></div>
</div>

</div>
<br/><br/><br/>
</body>
</html>


