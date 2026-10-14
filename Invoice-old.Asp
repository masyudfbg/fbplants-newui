<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Invoice</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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

</head>
<body>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <%
 '-------------awal-----------------------
 
 Dim xCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks,txtNoPO,v_CrossCekPO
 Dim recordTotalAmount, rsrecordTotalAmount, varTotalAmount, updateNetBalance
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""


 Dim Jobcode, LocationType, LocationCode, Reference, Remark

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 if request("fccOde")="" then
    fcCode="" '----- Nantinya Ambil dari Format Master
 else
    fcCode = request("fccOde")
 end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("DATECREATED")<>"" then
   DATECREATED   =Request("DATECREATED")
 else
   DATECREATED   =textDate(date)
 end if
 xCode = Request("xCode")
 HeaderRemarks = Request("HeaderRemarks")
 
 '---- Data Details
 JobCode=Request("JobCode")
 LocationType=Request("LocationType")
 LocationCode=Request("LocationCode")
 Reference=Request("Reference")
 Remarks=Request("Remarks")
 Amount=Request("Amount")
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	
	REMARKS=request("REMARKS")
	
	if Request("txtJenis")= "Header" then
	
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then
		
			'Simpan NO ReceiveNote nya terlebih dahulu sebelum di delete
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			rsPO.Open	" Select POCODE from INVOICEDETAIL where INVOICECODE = '" & fcCode & "'" & _
						" and fcBA = '" & Session("fcBA") & "'", connstr
			IF Not rsPO.eof then	   
				txtNoPO =  rsPO("POCODE")
			End if 
			rsPO.Close
			set rsPO = Nothing
			
			connstr.BeginTrans

			Set rsPO = Server.CreateObject("ADODB.Recordset")
			rsPO.Open	" Select ITEMCODE from INVOICEDETAIL where INVOICECODE = '" & fcCode & "'" & _
						" and fcBA = '" & Session("fcBA") & "'", connstr
			IF Not rsPO.eof then	   
				While Not rsPO.EOF
					v_CrossCekPO = cmdCrossCekReceiveNote(fcCode,txtNoPO,rsPO("ITEMCODE")) 
					rsPO.MoveNext
				Wend
			End if 
			rsPO.Close
			set rsPO = Nothing
			
			if err.number=0 then
				q = " Delete from INVOICEDETAIL where INVOICECODE = '" & fcCode & "' and fcBA = '" & Session("fcBA") & "'"
				connStr.execute(q)
			end if
			
			if err.number=0 then
				q = " Delete from INVOICE where  INVOICECODE = '" & fcCode & "' and fcBA = '" & Session("fcBA") & "'"
				connStr.execute(q)
			end if
			
			'response.Write(q)
			cmdView	= "True"

			if err.number  = 0 then
				connStr.CommitTrans
			end if
			
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		cmdView	= "True"

	elseif Request("txtJenis")= "Details" then
	
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then
			
			'Simpan NO ReceiveNote nya terlebih dahulu sebelum di delete
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			rsPO.Open 	" Select POCODE from INVOICEDETAIL where INVOICECODE = '" & fcCode & "'" & _
						" and fcBA='" & Session("fcBA") & "'", connstr

			IF Not rsPO.eof then	   
				txtNoPO =  rsPO("POCODE")
			End if 
			
			rsPO.Close
			set rsPO = Nothing
			
			connstr.BeginTrans
			v_CrossCekPO = cmdCrossCekReceiveNote(fcCode,txtNoPO,Request("ItemCode")) 

			if err.number = 0 then
				q = " Delete from INVOICEDETAIL where  INVOICECODE = '" & fcCode & "' and ITEMCODE = '" & Request("ItemCode") & _
					"' and fcBA='" & Session("fcBA") & "'"
				connStr.execute(q)
			end if

			if err.number  = 0 then
				connStr.CommitTrans
			end if
			
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
			
		cmdView	= "True"

		'response.Write(q)
	elseif Request("txtJenis")="New" then
	
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then
			
			Dim rsPO
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			rsPO.Open 	" Select INVOICECODE, suppliercode from INVOICE where INVOICECODE = '" & fcCode & "'" & _
						" and fcBA = '" & Session("fcBA") & "'", connstr
			connstr.BeginTrans

			IF rsPO.eof then	    
				Dim txtCodeSupplier,txtCodePO
				q = " Select POCODE,SUPPLIERCODE From RECEIVENOTE where RECEIVENOTECODE = '" & xCode & "'"  & _
					" and fcBA = '" & Session("fcBA") & "'"
				Dim rsLPO  
				set rsLPO = Server.CreateObject("ADODB.Recordset")
				rsLPO.Open q, connstr
				if not rsLPO.eof then
					txtCodePO=rsLPO("POCODE")
					txtCodeSupplier=rsLPO("SUPPLIERCODE")			
				End if 
			
				rsLPO.Close
				set rsLPO = Nothing
				q = " Insert into Invoice (INVOICECODE, INVOICEDATE, SUPPLIERCODE, PAID, REMARKS, NETBALANCE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					" values ('" & fcCode & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & txtCodeSupplier& "',0,'',0,'" & _
					Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
				connStr.execute(q)
			End if
			rsPO.Close	
			Dim n
			'response.Write "No " & (cdbl(request("no"))-1) & " amount : " &  request("AmountInvoice_"&TRIM(n))
			if err.number  = 0 then
				For n=1 to (cdbl(request("no"))-1) 
				if request("AmountInvoice_"&TRIM(n))<>"" then
					if err.number=0 then
						q = " Insert into  INVOICEDETAIL (INVOICECODE, ITEMCODE, QUANTITY, AMOUNT, POCODE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
							" values ('" & fcCode & "','" & request("Stock_"&TRIM(n)) & "','" & request("Quantity_"&TRIM(n)) &  _
							"','" & request("AmountInvoice_"&TRIM(n)) & "','" & txtCodePO & "'," & _
							"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
						connStr.execute(q)
						if err.number=0 then		
							q = " Update ReceiveNoteDetail  set AMOUNTINVOICE=AMOUNTINVOICE+" & request("AmountInvoice_"&TRIM(n))  & _
								", LASTUPDATE = sysdate " & _
								", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
								" where RECEIVENOTECODE='" & xCode & "' and PURCHASEITEMCODE= '" & request("Stock_"&TRIM(n)) & "'" & _
								" and fcBA='" & session("fcBA") & "'"
							'response.Write(q)
							connStr.execute(q)
							if err.number=0 then
								' Start Modification by NAF, 21-Sep-2012
								recordTotalAmount = "select sum(amount) totalAmount " & _
													"from invoicedetail " & _
													"where " & _ 
													"invoicecode = '" & fcCode & "' " & _ 
													"and fcba = '" & session("fcBA") & "'"

								' response.Write recordTotalAmount
								Set rsrecordTotalAmount = Server.CreateObject("ADODB.Recordset")
								rsrecordTotalAmount.CursorType = 3
								rsrecordTotalAmount.CursorLocation = 3 

								rsrecordTotalAmount.Open recordTotalAmount, connstr


								if Not rsrecordTotalAmount.Eof then
									varTotalAmount = rsrecordTotalAmount("totalAmount")
								Else
									varTotalAmount = 0
								End if


								updateNetBalance = 	" update invoice " & _
													" set " & _
													" netbalance = " & varTotalAmount & " - paid " & _
													", LASTUPDATE = sysdate " & _
													", LASTTIME = To_Char(sysdate,'HH24:MI') "& _												
													" where " & _
													" invoicecode = '" & fcCode & "' and " & _ 
													" fcba = '" & session("fcBA") & "'"

								connStr.execute(updateNetBalance)
								' End of Modification by NAF

								if err.number  = 0 then
									connStr.Commit
								end if

							end if
						end if
					end if
				End if	 
				Next
			end if
			
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	else
		cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
	cmdView	= request("cmdView")
End if

 Dim rsVoucher,SudahClose
 set rsVoucher = server.CreateObject("Adodb.recordset")
 q="select * from Invoice  " & _
  " where INVOICECODE='" & fcCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
 SudahClose="N"
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    DATECREATED=rsVoucher("INVOICEDATE")
	Dim rsTemp1 
    q="SELECT r.RECEIVENOTECODE as RECEIVENOTECODE FROM Invoicedetail inv,receivenote r" & _
      " where (r.POCODE = inv.POCODE and r.fcba = inv.fcba)" & _
	  " and INVOICECODE='" & fcCode & "' and" & _
      " r.fcBA='" & session("fcBA") & "'"
      response.Write q
	set rsVoucher = server.CreateObject("Adodb.recordset")
	rsVoucher.open q,connStr	
	if not rsVoucher.eof then
	   xCode=rsVoucher("RECEIVENOTECODE")
	End if
	SudahClose="Y"
 End if
 rsVoucher.Close
 Set rsVoucher = Nothing
 if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
 End if
 '-------------ISI Payment Form
 
 Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Purchase')"
	   
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
'response.Write strNurseryCulling
oRecordSet.Close
Set oRecordSet = Nothing
 %>
 <form name="frmEntry" method="post" action="Invoice.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Inv. No</td>
		<td align="left"><input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">&nbsp;*</font></td></tr><tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="DATECREATED" size="10" value="<%=DATECREATED%>"><font color="#FF0000">&nbsp;*</font>
		</td></tr>
		<tr>
		<td align="left">Receive Note</td>
		<td align="left">      
		<%if SudahClose="Y" then%>
		    <select name="xCode"><option value="<%=xCode%>" Selected><%=xCode%></option></select>
		<%else%>
		    <select name="xCode">
				<%
				Dim rsGoodsReceive, strGoodsReceive
				strGoodsReceive	=	" Select distinct(RECEIVENOTECODE) RECEIVENOTECODE from RECEIVENOTEDETAIL" & _
									" Where (AmountInvoice = 0 or  AMOUNTINVOICE IS NULL) and fcBA = '" & session("fcBA") & "'"
				
				Set rsGoodsReceive = Server.CreateObject("ADODB.Recordset")
				rsGoodsReceive.Open strGoodsReceive, connstr
				
				Response.Write strGoodsReceive
				
				if not rsGoodsReceive.eof then
				Do While Not rsGoodsReceive.eof
				if rsGoodsReceive("RECEIVENOTECODE") = xCode then%>
				<option value="<%=rsGoodsReceive("RECEIVENOTECODE")%>" Selected><%=rsGoodsReceive("RECEIVENOTECODE")%></option> 
				<%Else%>
				<option value="<%=rsGoodsReceive("RECEIVENOTECODE")%>"><%=rsGoodsReceive("RECEIVENOTECODE")%></option>  
				<%End if
					rsGoodsReceive.movenext
				loop
				end if
				rsGoodsReceive.Close
				set rsGoodsReceive = Nothing
				%>
            </select>
		<%End if%>
		<font color="#FF0000">&nbsp;*</font>
        </td>
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
	    <a href="javascript:launchwin('PrintInvoice.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<%If val_IsTrue = "Y" Then%>
			<a href="Invoice.Asp?fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0"></a>
		<%Else%>
			<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
		<%End if%>
		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Purchase"
		dataSubmodule 	= "Invoice"
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 and len(trim(xcode)) > 0 Then
   Dim rsGRN,intRecordsPerPage,cColor
   set rsGRN = server.CreateObject("Adodb.recordset")
   if SudahClose="Y" Then
      q="select sp.FCNAME as SuppDesc,invd.ITEMCODE,s.FCNAME as StockDesc,sum(invd.QUANTITY) as QUANTITY," & _
        "sum(invd.AMOUNT) as Amount From invoice inv,invoicedetail invd,stock s,supplier sp" & _
        " where (inv.INVOICECODE = invd.INVOICECODE and inv.fcba = invd.fcba) AND   (INVD.ITEMCODE   = S.FCCODE and invd.fcba = s.fcba) " & _
        " AND   (INV.SUPPLIERCODE = SP.FCCODE and inv.fcba = sp.fcba) and   (inv.INVOICECODE='" & fcCode & "')" & _
	    " and inv.fcBA='" & session("fcBA") & "'" & _
        " group by  sp.FCNAME ,invd.ITEMCODE,s.FCNAME"
   Else	 
      q="select r.POCODE,r.RECEIVENOTECODE,rd.PURCHASEITEMCODE,s.fcname,sum(rd.quantity) as Quantity,sum(rd.amountreceived)" & _
	    " as amountreceived  from receivenote r,receivenotedetail rd,stock s" & _
        " where (r.RECEIVENOTECODE = rd.RECEIVENOTECODE and r.fcba = rd.fcba) and   (rd.PURCHASEITEMCODE = s.FCCODE and rd.fcba=s.fcba)" & _
		" and r.RECEIVENOTECODE='" & xCode & "'" & _
        " and r.fcBA='" & session("fcBA") & "'" & _
        " group by r.POCODE,r.RECEIVENOTECODE,rd.PURCHASEITEMCODE,s.fcname " 
   End if
   'response.Write q	
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.CursorType = 3
   rsGRN.CursorLocation = 3
   rsGRN.Open q, connstr%>
   <form action="Invoice.Asp" method="post" name="eForm">
   <table class="data" border="1" width="100%">
   <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">PO Code</td>
	   <td align="center">GRN Code</td>
	   <td align="center">Stock Code</td>
	   <td align="center">Stock Desc.</td>
	   <td align="center">Qty Received</td>
	   <td align="center">Amount Received</td>
	   <td align="center">Amount Invoice</td>
	   </tr>
   <%ELSE%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">&nbsp</td>
	   <td align="center">Supplier Name</td>
	   <td align="center">Stock Code</td>
	   <td align="center">Stock Desc.</td>
	   <td align="center">Quantity</td>
	   <td align="center">Amount</td>
	   </tr>
   <%End if
   if Not rsGRN.eof then
      nCountofField=rsGRN.fields.count-1
	  no=1
	  While Not rsGRN.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<%IF SudahClose="Y" then%>
				<td valign="middle">
				<%If val_IsTrue = "Y" Then%>
					<a href="Invoice.Asp?fcCode=<%=fcCode%>&ItemCode=<%=rsGRN("ITEMCODE")%>&DATECREATED=<%=DATECREATED%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGRN(1)%>" border="0"></a></td>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End if%>
					<td align="left" title="Supplier" valign="top"><%=rsGRN("SuppDesc")%></td>
					<td align="left" title="Item Code" valign="top"><%=rsGRN("ITEMCODE")%></td>
					<td align="left" title="Item Name" valign="top"><%=rsGRN("StockDesc")%></td>
					<td align="right" title="Quantity" valign="top"><%=FormatNumber(rsGRN("Quantity"),0)%></td>
					<td align="right" title="Amount" valign="top"><%=FormatNumber(rsGRN("Amount"),0)%></td>
		                
			<%End if%>

			<%if SudahClose<>"Y" Then%>
                    <td align="left" title="" valign="top"><%=rsGRN("POCODE")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("RECEIVENOTECODE")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("PURCHASEITEMCODE")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("fcname")%></td>
			        <td align="right" title="" valign="top"><%=FormatNumber(rsGRN("Quantity"),0)%></td>
                    <td align="right" title="" valign="top"><%=FormatNumber(rsGRN("amountreceived"),0)%></td>
		            <td align="left" width="10"><input type="text" name="AmountInvoice_<%=no%>" width="10" value="<%=request("AmountInvoice_"&TRIM(no))%>"></td>
				 <input type="hidden" name="Stock_<%=no%>" value="<%=rsGRN("PURCHASEITEMCODE")%>">
				 <input type="hidden" name="Quantity_<%=no%>" value="<%=rsGRN("Quantity")%>">
				 </td>
			<%End if%>
		  </tr>
		  <%
		  no=no+1
		  rsGRN.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsGRN.Close
   set rsGRN = Nothing%>
	  <tr><td colspan="8" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="No" value="<%=No%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%If val_IsTrue = "Y" Then%>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Submit">
		  <%End if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </TABLE>
   </div>
   </form>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->
 
</body>
</html>
