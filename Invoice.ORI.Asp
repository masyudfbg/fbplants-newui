<%Option Explicit%>
<%'On Error Resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<html>
<head>
<title>iPS - Invoice</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
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
<table width="100%">
 <tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
 <%'-------------awal-----------------------%>
 <%
 Dim xCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks,txtNoPO,v_CrossCekPO

 Dim Jobcode, LocationType, LocationCode, Reference, Remark

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 if request("fccOde")="" then
    fcCode="INV/" '----- Nantinya Ambil dari Format Master
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

	   'Simpan NO ReceiveNote nya terlebih dahulu sebelum di delete
	   Set rsPO = Server.CreateObject("ADODB.Recordset")
	   rsPO.Open "Select POCODE from INVOICEDETAIL where INVOICECODE='" & fcCode & "'" & _
	            " and fcBA='" & Session("fcBA") & "'", connstr
	   IF Not rsPO.eof then	   
	      txtNoPO =  rsPO("POCODE")
       End if 
	   rsPO.Close
	   set rsPO = Nothing
       connstr.BeginTrans
	   Set rsPO = Server.CreateObject("ADODB.Recordset")
	   rsPO.Open "Select ITEMCODE from INVOICEDETAIL where INVOICECODE='" & fcCode & "'" & _
	            " and fcBA='" & Session("fcBA") & "'", connstr
	   IF Not rsPO.eof then	   
	      While Not rsPO.EOF
			  v_CrossCekPO = cmdCrossCekReceiveNote(fcCode,txtNoPO,rsPO("ITEMCODE")) 
			  rsPO.MoveNext
          Wend
	   End if 
	   rsPO.Close
	   set rsPO = Nothing
       if err.number=0 then
          q="Delete from INVOICEDETAIL where INVOICECODE='" & fcCode & "' and fcBA='" & Session("fcBA") & "'"
   	      connStr.execute(q)
	   end if
	   if err.number=0 then
          q="Delete from INVOICE where  INVOICECODE='" & fcCode & "' and fcBA='" & Session("fcBA") & "'"
   	      connStr.execute(q)
	   end if
	   'response.Write(q)
       cmdView	= "True"

	   if err.number  = 0 then
		  connStr.CommitTrans
	   end if
       cmdView	= "True"

	elseif Request("txtJenis")= "Details" then
	   'Simpan NO ReceiveNote nya terlebih dahulu sebelum di delete
	   Set rsPO = Server.CreateObject("ADODB.Recordset")
	   rsPO.Open "Select POCODE from INVOICEDETAIL where INVOICECODE='" & fcCode & "'" & _
	            " and fcBA='" & Session("fcBA") & "'", connstr
	   IF Not rsPO.eof then	   
	      txtNoPO =  rsPO("POCODE")
       End if 
	   rsPO.Close
	   set rsPO = Nothing
       connstr.BeginTrans
 	   v_CrossCekPO = cmdCrossCekReceiveNote(fcCode,txtNoPO,Request("ItemCode")) 
    
       if err.number = 0 then
	      q="Delete from INVOICEDETAIL where  INVOICECODE='" & fcCode & "' and ITEMCODE='" & Request("ItemCode") & _
	        "' and fcBA='" & Session("fcBA") & "'"
   	      connStr.execute(q)
	   end if
       cmdView	= "True"

	   if err.number  = 0 then
		  connStr.CommitTrans
	   end if

	   'response.Write(q)
	elseif Request("txtJenis")="New" then
	  Dim rsPO
	  Set rsPO = Server.CreateObject("ADODB.Recordset")
	  rsPO.Open "Select INVOICECODE,suppliercode from INVOICE where INVOICECODE='" & fcCode & "'" & _
	            " and fcBA='" & Session("fcBA") & "'", connstr
	  connstr.BeginTrans
	  IF rsPO.eof then	    
	     Dim txtCodeSupplier,txtCodePO
		 q="Select POCODE,SUPPLIERCODE From RECEIVENOTE where RECEIVENOTECODE ='" & xCode & "'"  & _
		   " and fcBA='" & Session("fcBA") & "'"
		 Dim rsLPO  
		 set rsLPO = Server.CreateObject("ADODB.Recordset")
		 rsLPO.Open q,connstr,1,1
		 if not rsLPO.eof then
		    txtCodePO=rsLPO("POCODE")
		    txtCodeSupplier=rsLPO("SUPPLIERCODE")			
		 End if 
		 rsLPO.Close
		 set rsLPO = Nothing
	     q="Insert into Invoice (INVOICECODE, INVOICEDATE, SUPPLIERCODE, PAID, REMARKS, NETBALANCE, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		 "values ('" & fcCode & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & txtCodeSupplier& "',0,'',0,'" & _
		 Session("CodeName") & "','" & Session("CodeName") & "','" & _
		 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
	     connStr.execute(q)
		 'response.Write(q)
	  End if
	  rsPO.Close	
	  Dim n
	  'response.Write "No " & request("no") & " amount : " &  request("AmountInvoice_"&TRIM(n))
      if err.number  = 0 then
	    For n=1 to (cdbl(request("no"))-1) 
	      if request("AmountInvoice_"&TRIM(n))<>"" then
             if err.number=0 then
			    q="Insert into  INVOICEDETAIL (INVOICECODE, ITEMCODE, QUANTITY, AMOUNT, POCODE, FCENTRY, FCEDIT, FCIP, FCBA) " & _
	            "values ('" & fcCode & "','" & request("Stock_"&TRIM(n)) & "','" & request("Quantity_"&TRIM(n)) &  _
			    "','" & request("AmountInvoice_"&TRIM(n)) & "','" & txtCodePO & "'," & _
		        "'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		        Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		        'response.Write(q)
	            connStr.execute(q)
			    if err.number=0 then		
                   q="Update ReceiveNoteDetail  set AMOUNTINVOICE=AMOUNTINVOICE+" & request("AmountInvoice_"&TRIM(n))  & _
                   " where RECEIVENOTECODE='" & xCode & "' and PURCHASEITEMCODE= '" & request("Stock_"&TRIM(n)) & "'" & _
			       " and fcBA='" & session("fcBA") & "'"
		           'response.Write(q)
	               connStr.execute(q)
				   if err.number=0 then
					  connStr.Commit
				   end if
				end if    
		     end if		
		   End if	 
	     Next	
	   end if
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
   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
 End if
 '-------------ISI Payment Form%>
 <form name="frmEntry" method="post" action="Invoice.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Inv. No</td>
		<td align="left"><input type="text" maxlength="15" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">*</font></td></tr><tr>
		<td align="left">Date</td>
		<td align="left"><input type="text" maxlength="10" name="DATECREATED" size="10" value="<%=DATECREATED%>">
		<a href="javascript:show_calendar('frmEntry.DATECREATED');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>
		<tr>
		<td align="left">Receive Note</td>
		<td align="left">      
		<%if SudahClose="Y" then%>
		    <select name="xCode"><option value="<%=xCode%>" Selected><%=xCode%></option></select>
		<%else%>
		    <select name="xCode">
		    <%
			dim rsCallCombo
			Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
			rsCallCombo.CursorType = 3
			rsCallCombo.CursorLocation = 3
			rsCallCombo.Open "Select distinct(RECEIVENOTECODE) from RECEIVENOTEDETAIL Where AmountInvoice=0 or  AMOUNTINVOICE IS NULL and fcBA='" & session("fcBA") & "'", connstr
			if not rsCallCombo.eof then
			 Do While Not rsCallCombo.eof
				if rsCallCombo("RECEIVENOTECODE") = xCode then%>
				   <option value="<%=rsCallCombo("RECEIVENOTECODE")%>" Selected><%=rsCallCombo("RECEIVENOTECODE")%></option> 
				<%Else%>
				   <option value="<%=rsCallCombo("RECEIVENOTECODE")%>"><%=rsCallCombo("RECEIVENOTECODE")%></option>  
				<%End if%>
				<%rsCallCombo.movenext
			 loop
			End if
			rsCallCombo.Close
			set rsCallCombo = Nothing%>
            </select>
		<%End if%>
        </td>
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
	    <a href="javascript:launchwin('PrintInvoice.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<a href="Invoice.Asp?fcCode=<%=fcCode%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 Then
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
	   <tr bgcolor="#6666FF">
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
	   <tr bgcolor="#6666FF">
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
				<a href="Invoice.Asp?fcCode=<%=fcCode%>&ItemCode=<%=rsGRN("ITEMCODE")%>&txtJenis=Details">
				<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGRN(1)%>" border="0"></a></td>
		                
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
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Submit">
		  <%End if%>
		  <input Type="Reset" style="background-color: #298CF7; color: #F7C610; font-weight:
		  bold; font-size: 8pt; font-family:Verdana" name="Reset" value="Reset">
		  <input Type="Button" style="background-color: #298CF7; color: #F7C610; font-weight: 
		  bold; font-size: 8pt; font-family:Verdana" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </TABLE>
   </form>
<%End if%>  
 
  <%'-------------akhir----------------------%>
 </td>
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
</html>
