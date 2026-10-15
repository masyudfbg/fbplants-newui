<%
	response.buffer = True
	validatesession
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
	<title>ePlaS - Supplying Declaration Final</title>
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
Function cmdCheckHistNursery(initfcba, initNURSERYCODE, initLineNoHist)
	dim qCheckHistNursery,rsCheckHistNursery,qTotalDataNursery
	
	qCheckHistNursery = " select count(1) totalData" & _
						" from stocktranshist" & _
						" where fcba = '" & initfcba & "'" & _
						" and targetcode = '" & initNURSERYCODE & "'" & _
						" and lineno > " & initLineNoHist & ""
						
	'Response.Write qCheckHistNursery

	set rsCheckHistNursery = Server.CreateObject("ADODB.Recordset")
	rsCheckHistNursery.Open qCheckHistNursery, connstr
	if not rsCheckHistNursery.eof then
		qTotalDataNursery = cdbl(rsCheckHistNursery("totalData"))
	else
		qTotalDataNursery = 0
	end if 	

	cmdCheckHistNursery = qTotalDataNursery
End Function

Function CheckPrimary(initialFieldCode, initialFcba, initYear, initMonth)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	" select count(1) total" & _
	                " from supplyingdeclaration " & _
					" where fieldcode ='" & initialFieldCode & "'" & _
					" and year = '" & initYear & "'" & _
					" and month = '" & initMonth & "'" & _
					" and fcba = '" & initialFcba & "'"

	 ' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckPrimary = initialAvailable
End Function
%>
		<%
			'--Awal--
			Dim xCode,xstoreCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,v_CrossCekPO
			Dim Remarks
			Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
			Dim Jobcode, LocationType, LocationCode, Reference, Remark
			Dim qLineMax, rsqLineMax, tempMaxLine
			Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
			Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
			Dim sqlSQLCheck, rsSQL, qtyRemained
			Dim qCheckStore, rsqCheckStore
			Dim txtLocationType	

			Dim qSupplyingDeclaration, rsSupplyingDeclarationInfo, tempQuantity, tempDecType
			Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, v_FCCODE, v_FCNAME
			Dim qSupplyingInfo, rsSupplyingInfo, qDeclarationType, rsDeclarationType
			Dim DECLARATIONTYPE, parcode, totalstand
			Dim FIELDCODE 
			Dim TRANSDATE
			Dim txtDocumentNo 

			Dim txtGRNCode : txtGRNCode = 0
			Dim oRs : oRs = 0
			Dim oSql : oSql = 0
			Dim txtLocationCode : txtLocationCode = ""
			Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

			
			txtDocumentNo = Request("txtDocumentNo")	
			TRANSDATE	=	Request("TRANSDATE")
			txtTable =Request("txtTable")
			txtModule=Request("txtModule")
			txtSearch=Request("txtSearch")
			txtField =Request("txtField")
			
			if request("fccOde")="" then
				fcCode="" '----- Nantinya Ambil dari Format Master
			else
				fcCode = request("fccOde")
			end if
			
			If Len(Trim(fcCode)) = 0 Then fcCode = ""
			
			txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
			
			'---- Data Header
			if "" & TRANSDATE = "" then TRANSDATE	=	textDate(date) 
			
			
			'---- Data Details
			quantity=Request("quantity")
			' LocationType=Request("LocationType")
			' LocationCode=Request("LocationCode")
			' Reference=Request("Reference")
			' Remarks=Request("Remarks")
			' Amount=Request("Amount")
			' PurchaseLineNoKey = Request("PurchaseLineNoKey")
			
			'--- Action ke Database di Oracle [ insert,edit,delete ]
			if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
				REMARKS	=	request("REMARKS")
				if Request("txtJenis")= "Header" then
					'Simpan NO PO nya terlebih dahulu sebelum di delete
					connstr.BeginTrans
					
					Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
					If Cek_IsDailyClosing = "FALSE" then

						Set rsPO = Server.CreateObject("ADODB.Recordset")
						rsPO.Open " select documentno, transdate " & _
								  " from supplyingdeclarationhdr" & _
								  " where documentno = '" & txtDocumentNo & "'" & _
								  "       and fcBA = '" & Session("fcBA") & "'", connstr
						IF Not rsPO.eof then
							txtNoPO =  rsPO("documentno")
						End if
						rsPO.Close
						set rsPO = Nothing

						q = " delete from supplyingdeclarationhdr" & _
							" where documentno = '" & txtDocumentNo & "'" & _
							"       and fcba = '" & Session("fcBA") & "'"
						
						connStr.execute(q)
						fcCode = "(NEW)"
						
						if err.number = 0 then
							connStr.CommitTrans
						end if
						
					else
						err.number = 1001
						err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					End If
					
					cmdView	= "True"
				elseif Request("txtJenis")= "Details" then
					'Simpan NO PO nya terlebih dahulu sebelum di delete
					' connstr.BeginTrans
					
					' Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
					' If Cek_IsDailyClosing = "FALSE" then
						
					' 	Set rsPO = Server.CreateObject("ADODB.Recordset")
					' 	rsPO.Open "Select POCODE, STORECODE from ReceiveNote where RECEIVENOTECODE='" & fcCode & "'" & _
					' 	" and fcBA='" & Session("fcBA") & "'", connstr
					' 	IF Not rsPO.eof then
					' 		txtNoPO =  rsPO("POCODE")
					' 		txtStoreCode = rsPO("STORECODE")
					' 	End if
					' 	rsPO.Close
					' 	set rsPO = Nothing
						
					' 	q = " Select LineNoHist" & _
					' 		"       , quantity" & _
					' 		"       , PURCHASEITEMCODE" & _
					' 		"       , PURCHASELINENOKEY" & _
					' 		"       , LOCATIONTYPE" & _
					' 		"       , LOCATIONCODE" & _
					' 		" from ReceiveNoteDetail" & _
					' 		" where RECEIVENOTECODE = '" & fcCode & "'" & _
					' 		"       and PURCHASELINENOKEY = '" & Request("PURCHASELINENOKEY") & "'" & _
					' 		"       and fcBA = '" & Session("fcBA") & "'"
						
					' 	Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
					' 	rsReceiveDetail.Open q, connstr
								
					' 	IF Not rsReceiveDetail.eof then
					' 		txtLineNoHist = rsReceiveDetail("LINENOHIST")
					' 		originalQuantity = CDbl(rsReceiveDetail("QUANTITY"))
					' 		txtItemCodeDelete = rsReceiveDetail("PURCHASEITEMCODE")
					' 		txtItemLineKeyDelete = rsReceiveDetail("PURCHASELINENOKEY")
					' 		txtLocationType = rsReceiveDetail("LOCATIONTYPE")
					' 		txtLocationCode = rsReceiveDetail("LOCATIONCODE")
					' 	End if
					' 	rsReceiveDetail.Close
					' 	set rsReceiveDetail = Nothing					
						
						
						
						
					' 	' Add by NAF, 20 Jan 2014
					' 	' To handle Check Store for CAPEX
					' 	qCheckStore = " select LOCATIONTYPE_API.get_item('" & txtLocationType & "') item from dual "

					' 	set rsqCheckStore = Server.CreateObject("ADODB.Recordset")
					' 	rsqCheckStore.Open qCheckStore, connstr
					' 	if not rsqCheckStore.eof then
					' 		tempItem = rsqCheckStore("item")
					' 	else
					' 		tempItem = ""
					' 	end if 
					' 	rsqCheckStore.Close
					' 	Set rsqCheckStore = Nothing
					' 	' End of modification
									
						
					' 	if tempItem = "MATERIAL" then
						
					' 		If txtLocationType = "MN" Then
					' 			v_CrossCekDetail = cmdCheckHistNursery(Session("fcBA"), txtLocationCode, txtLineNoHist)
					
					' 				qLessZero = " select islesszero" & _
					' 							" from stock" & _
					' 							" where fccode = '" & Request("ItemCode") & "'" & _
					' 							"       and fcba = '" & session("fcBA") & "'" 
									
					' 				set rsLessZero = Server.CreateObject("ADODB.Recordset")
					' 				rsLessZero.Open qLessZero, connstr
					' 				if not rsLessZero.eof then
					' 					tempislesszero = rsLessZero("islesszero")
					' 				else
					' 					tempislesszero = "FALSE"
					' 				end if
									
					' 				qQuantityOnHand = " select quantity" & _
					' 								   " from stockonhandnursery" & _
					' 								   " where storecode = '" & txtStoreCode & "'" & _
					' 								   "       and stockcode  = '" & Request("ItemCode") & "'" & _
					' 								   "       and nurserycode  = '" & txtLocationCode & "'" & _
					' 								   "       and fcba = '" & session("fcBA") &"' "
													  
					' 				set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
					' 				rsQuantityOnHand.Open qQuantityOnHand, connstr
					' 				if not rsQuantityOnHand.eof then
					' 					QuantityOnHandNursery = cdbl(rsQuantityOnHand("quantity"))
					' 				else
					' 					QuantityOnHandNursery = 0
					' 				end if
									
					' 				If tempislesszero = "TRUE" Or (tempislesszero = "FALSE" And ((QuantityOnHandNursery - originalQuantity) >= 0)) Then
									
					' 					q =	" Delete from stocktranshist" & _
					' 						" where lineno = " & txtLineNoHist & "" & _
					' 						"       and fcba = '" & Session("fcBA") & "'"

					' 					connStr.execute(q)

					' 					q = " update stockonhand" & _
					' 						" set quantity = quantity - " & originalQuantity & "" & _
					' 						"     , lastupdate = sysdate " & _
					' 						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
					' 						" where storecode = '" & txtStoreCode & "'" & _
					' 						"       and stockcode = '" & Request("ItemCode") & "'" & _
					' 						"       and fcba = '" & session("fcBA") &"'"

					' 					connStr.execute(q)

					' 					q = " update stockonhandnursery " & _
					' 						" set quantity = quantity - " & originalQuantity & "" & _
					' 						"     , lastupdate = sysdate " & _
					' 						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
					' 						" where storecode = '" & txtStoreCode & "'" & _
					' 						"       and stockcode  = '" & Request("ItemCode") & "'" & _ 
					' 						"       and nurserycode  = '" & txtLocationCode & "'" & _ 
					' 						"       and fcba = '" & session("fcBA") &"'"

					' 					connStr.execute(q)
										
					' 					q = " delete from receivenotedetail" & _
					' 						" where receivenotecode = '" & fcCode & "'" & _
					' 						"       and purchaselinenokey = '" & PurchaseLineNoKey & "'" & _
					' 						"       and fcba = '" & Session("fcBA") & "'"
										
					' 					connStr.execute(q)
										
					' 					q = " update purchaseorderdetails" & _
					' 						" set rowstateid = 'Opened'" & _
					' 						" where pocode = '" & txtNoPo & "'" & _
					' 						"       and linenokey = '" & txtItemLineKeyDelete & "'" & _
					' 						"       and fcBA = '" & Session("fcBA") & "'"

					' 					connStr.execute(q)
										
					' 					v_CrossCekPO =  cmdCrossCekPOUpdate (txtNoPo, txtItemLineKeyDelete,  originalQuantity )
										
					' 					if err.number=0 then
					' 						connStr.CommitTrans
					' 					End if
										
					' 				Else
					' 					err.number = 1001
					' 					Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
					' 				End If
								
					' 		Else
							
					' 			v_CrossCekDetail = cmdCheckHist(txtLineNoHist, Request("ItemCode"))
								
					' 			' Add new function to update qty received of PO Detail, by NAF
					' 			v_CrossCekPO =  cmdCrossCekPOUpdate (txtNoPo, txtItemLineKeyDelete,  originalQuantity ) 
					' 			' End of Modification
								
					' 			qLessZero = " select islesszero" & _
					' 						" from stock" & _
					' 						" where fccode = '" & Request("ItemCode") & "'" & _
					' 						"       and fcba = '" & session("fcBA") & "'" 
								
					' 			set rsLessZero = Server.CreateObject("ADODB.Recordset")
					' 			rsLessZero.Open qLessZero, connstr
					' 			if not rsLessZero.eof then
					' 				tempislesszero = rsLessZero("islesszero")
					' 			else
					' 				tempislesszero = "FALSE"
					' 			end if
								
					' 			qQuantityOnHand = " select quantity" & _
					' 							   " from stockonhand" & _
					' 							   " where storecode = '" & txtStoreCode & "'" & _
					' 							   "       and stockcode  = '" & Request("ItemCode") & "'" & _
					' 							   "       and fcba = '" & session("fcBA") &"' "
												  
					' 			set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
					' 			rsQuantityOnHand.Open qQuantityOnHand, connstr
					' 			if not rsQuantityOnHand.eof then
					' 				QuantityOnHandStock = cdbl(rsQuantityOnHand("quantity"))
					' 			else
					' 				QuantityOnHandStock = 0
					' 			end if
								
					' 			If tempislesszero = "TRUE" Or (tempislesszero = "FALSE" And ((QuantityOnHandStock - originalQuantity) >= 0)) Then
									
					' 				q =  "Delete from stocktranshist " & _
					' 					 "where " & _
					' 					 "lineno = " & txtLineNoHist & " and " & _
					' 					 "fcba = '" & Session("fcBA") & "'"
								
					' 				connStr.execute(q)
									
					' 				q = " update stockonhand " & _
					' 					" set " & _
					' 					" quantity = quantity - " & originalQuantity & " " & _
					' 					", LASTUPDATE = sysdate " & _
					' 					", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					' 					"where " & _
					' 					"storecode = '" & txtStoreCode & "' and " & _
					' 					"stockcode  = '" & txtItemCodeDelete & "' and " & _ 
					' 					"fcba = '" & session("fcBA") &"' "
										
					' 				connStr.execute(q)							
								
					' 				q="Delete from RECEIVENOTEDETAIL where  RECEIVENOTECODE='" & fcCode & "' and PURCHASELINENOKEY='" & PurchaseLineNoKey & _
					' 				"' and fcBA='" & Session("fcBA") & "'"
									
					' 				connStr.execute(q)
									
					' 				q = " update purchaseorderdetails" & _
					' 					" set rowstateid = 'Opened'" & _
					' 					" where pocode = '" & txtNoPo & "'" & _
					' 					"       and linenokey = '" & txtItemLineKeyDelete & "'" & _
					' 					"       and fcBA = '" & Session("fcBA") & "'"

					' 				connStr.execute(q)
								
					' 				if err.number=0 then
					' 					connStr.CommitTrans
					' 				End if
									
					' 			Else
					' 				err.number = 1001
					' 				Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
					' 			End If
								
					' 		End If
							
					' 	else
					' 		' Add new function to update qty received of PO Detail, by NAF
					' 		v_CrossCekPO =  cmdCrossCekPOUpdate (txtNoPo, txtItemLineKeyDelete,  originalQuantity ) 
					' 		' End of Modification

					' 		q="Delete from RECEIVENOTEDETAIL where  RECEIVENOTECODE='" & fcCode & "' and PURCHASELINENOKEY='" & PurchaseLineNoKey & _
					' 				"' and fcBA='" & Session("fcBA") & "'"
									
					' 		connStr.execute(q)
						
					' 		if err.number=0 then
					' 			connStr.CommitTrans
					' 		End if
								
					' 	end if
						
						
					' else
					' 	err.number = 1001
					' 	err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
					' End If
					
					cmdView	= "True"
					
				elseif Request("txtJenis")="New" then
					connstr.BeginTrans
					
					Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
					If Cek_IsDailyClosing = "FALSE" then
					
					'--Add by Ukie to handle Auto Number of PR
					' If Ucase(fcCode) = Ucase("(NEW)") Then
					' 	oSql = " select nvl(controlnumber_api.get_nextnumber(upper('receivenote')" & _
					' 		   "                                         , " & year(textDate(DATECREATED)) & "" & _
					' 		   "                                         , " & month(textDate(DATECREATED)) & "" & _
					' 		   "                                         , '" & Session("fcBA") & "'" & _
					' 		   " ), 'FALSE') grncode from dual"
						
					' 	'Response.Write oSql
							   
					' 	Set oRs = Server.CreateObject("ADODB.Recordset")
					' 	oRs.Open oSql, connstr, 3, 3
					' 	If Not oRs.EOF Then
					' 		txtGRNCode = oRs("grncode")
					' 	End If
					' 	oRs.Close
					' 	Set oRs = Nothing
					' 	fcCode = txtGRNCode
					' End If
					
					' If txtGRNCode = "FALSE" Then
					' 	err.Number = 1001
					' 	err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
					' End If
					
					'Response.Write fcCode & txtGRNCode
					'--End Added
					
					
					' Dim rsPO
					'   Set rsPO = Server.CreateObject("ADODB.Recordset")
					'   rsPO.Open "Select RECEIVENOTECODE from ReceiveNote where RECEIVENOTECODE='" & fcCode & "'" & _
					' 			" and fcBA='" & Session("fcBA") & "'", connstr
					'   IF rsPO.eof then	    
					' 	 Dim txtCodeSupplier
					' 	 q="Select SUPPLIERCODE From PurchaseOrder where Pocode ='" & xCode & "'"  & _
					' 	   " and fcBA='" & Session("fcBA") & "'"
					' 	 Dim rsLPO  
					' 	 set rsLPO = Server.CreateObject("ADODB.Recordset")
					' 	 rsLPO.Open q,connstr,1,1
					' 	 if not rsLPO.eof then
					' 		txtCodeSupplier=rsLPO("SupplierCode")
					' 	 End if 
					' 	 rsLPO.Close
					' 	 set rsLPO = Nothing
						 q="Insert into ReceiveNote (RECEIVENOTECODE, RNDATE, POCODE,SupplierCode, Storecode, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
						 "values ('" & fcCode & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & xCode & "','" & txtCodeSupplier & "','" & _
						 xstoreCode & "','" & _
						 Session("CodeName") & "','" & Session("CodeName") & "','" & _
						 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
						 connStr.execute(q)
						 'response.Write(q)
					  ' End if
					  rsPO.Close	
					  Dim n
	  
						For n=1 to (cdbl(request("no"))-1) 
							if request("LT_"&TRIM(n))<>"" and request("LC_"&TRIM(n))<>"" and request("QtyReceived_"&TRIM(n))<>"" then
								 'Cari Control Job Stock Code 
								 if  request("JobCode_"&TRIM(n))="" Then
									 Dim txtControlJOb
									 q="select sg.CONTROLJOB as CONTROLJOB  from stockgroup sg,stock s" & _
									   " where (sg.FCCODE = s.STOCKGROUPCODE and sg.fcba = s.fcba) " & _
									   " and s.FCCODE='" & request("Stock_"&TRIM(n)) & "'" & _
									   " and s.fcBA='" & Session("fcBA") & "'"
									 set rsStock = Server.CreateObject("ADODB.Recordset")
									 rsStock.Open q, connstr
									 if not rsStock.eof then
										txtControlJob=rsStock("CONTROLJOB")
									 End if 
									 'rsStock.Close
									 'set rsStock = Nothing
								 else
									txtControlJob=request("JobCode_"&TRIM(n))
								 end if
								 
								if  request("AMOUNTReceived_"&TRIM(n))="" Then
									Dim txtamountreceived
									q = " select p.linenokey" & _
										"        , p.itemcode" & _
										"        , s.fcname" & _
										"        , sum(p.quantity) as quantity" & _
										"        , sum(amount) as amount" & _
										" from purchaseorderdetails p" & _
										"      , stock s" & _
										" where (p.itemcode = s.fccode and p.fcba = s.fcba)" & _
										"       and pocode = '" & xCode & "'" & _
										"       and s.fccode = '" & request("Stock_"&TRIM(n)) & "'" & _
										"       and p.linenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
										"       and p.fcba = '" & session("fcBA") & "'" & _
										" group by p.linenokey" & _
										"          , p.itemcode" & _
										"          , s.fcname"

									set rsamr = Server.CreateObject("ADODB.Recordset")
									rsamr.Open q, connstr
									
									if not rsamr.eof then
										txtamountreceived = Round((cdbl(rsamr("Amount"))/cdbl(rsamr("Quantity")))*request("QtyReceived_"&TRIM(n)), 0)
										PurchaseLineNoKey = rsamr("lineNoKey")
									End if 
								end if
								 
								 
								' Perlu diperbaiki, 2012-12-28
								sqlSQLCheck = " select (quantity - received) as qtyRemained from purchaseorderdetails " & _ 
												 " where POCode='" & xCode & "' and LineNoKey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
												 " and fcBA='" & session("fcBA") & "'"
												
								Set rsSQL = Server.CreateObject("ADODB.Recordset")

								rsSQL.Open sqlSQLCheck, connstr, adOpenStatic 

								if not rsSQL.eof then
									qtyRemained = cdbl(rsSQL("qtyRemained"))
								End if
								rsSQL.Close
								set rsSQL = nothing		

								if ( qtyRemained - cdbl(request("QtyReceived_"&TRIM(n))) ) < 0 then
										err.Number = 1001
										err.description = "Total QTY yang diterima tidak boleh lebih besar dari total QTY dalam item PO"
								Else
									q="Update PurchaseOrderDetails  set RECEIVED=received+" & request("QtyReceived_"&TRIM(n))  & _
									  ", LASTUPDATE = sysdate " & _
									  ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
									" where POCode='" & xCode & "' and LineNoKey= '" & request("LineNoKey_"&TRIM(n)) & "'" & _
									" and fcBA='" & session("fcBA") & "'"
									connStr.execute(q)							
												
									' Add by NAF, 20 Jan 2014
									' To handle Check Store for CAPEX
									qCheckStore = " select LOCATIONTYPE_API.get_item('" & request("LT_"&TRIM(n)) & "') item from dual "
									
									set rsqCheckStore = Server.CreateObject("ADODB.Recordset")
									rsqCheckStore.Open qCheckStore, connstr
									if not rsqCheckStore.eof then
										tempItem = rsqCheckStore("item")
									else
										tempItem = ""
									end if 
									rsqCheckStore.Close
									Set rsqCheckStore = Nothing
									' End of modification
												
									' Add by NAF to handle if condition, except MATERIAL
									if tempItem = "MATERIAL" then 
									
										' Add by NAF to handle Inventory Stock On Hand
										' Insert the History of Received Inventory Transaction
										
										qLineMax = "select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
												   "where " & _
												   "fcba = '" & session("fcBA") & "'" 
											
										set rsqLineMax = Server.CreateObject("ADODB.Recordset")
										rsqLineMax.Open qLineMax, connstr
										if not rsqLineMax.eof then
											tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
										else
											tempMaxLine = 0
										end if 

										tempMaxLine = tempMaxLine + 1
											
										
										q = "Insert into StockTransHist ( " & _
											" LINENO, FCBA, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY, LASTUPDATE, LASTTIME) values ( "& _
											  tempMaxLine & ",'" & _ 
											  session("fcBA") & "','" & _
											  request("Stock_"&TRIM(n)) &"'," & _
											  "to_date('" & DATECREATED & "','mm/dd/yyyy') ,'" & _
											  "RECEIVED" & "','" & _
											  xstoreCode & "'," & _
											  "'PR','" & _
											  xCode & "','" & _
											  request("LT_"&TRIM(n)) & "','" & _
											  request("LC_"&TRIM(n)) & "'," & _
											  request("QtyReceived_"&TRIM(n)) & ", sysdate, To_Char(sysdate,'HH24:MI') ) "
											  
										'response.Write(q)  
											  
										connStr.execute(q)
										
										'****** Add By Ukie at 20.03.2013 to handle rounding
										If ( qtyRemained - cdbl(request("QtyReceived_"&TRIM(n))) ) = 0 Then
											oQuery = " select nvl(sum(nvl(amount, 0)), 0) amount" & _
													 " from purchaseorderdetails" & _
													 " where pocode = '" & xCode & "'" & _
													 "       and linenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
													 "       and fcba = '" & session("fcBA") & "'"

											set oRecordSet = Server.CreateObject("ADODB.Recordset")
											oRecordSet.Open oQuery, connstr
											If Not oRecordSet.eof Then
												nFixedPOAmount = cdbl(oRecordSet("amount"))
											Else
												nFixedPOAmount = 0
											End If
											oRecordSet.Close()
											set oRecordSet = Nothing
											
											oQuery = " select nvl(sum(nvl(amountreceived, 0)), 0) amountreceived" & _
													 " from receivenotedetail rnd" & _
													 "      , receivenote rn" & _
													 " where rnd.receivenotecode = rn.receivenotecode" & _
													 "       and rnd.fcba = rn.fcba" & _
													 "       and rn.pocode = '" & xCode & "'" & _
													 "       and rnd.purchaselinenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
													 "       and rnd.fcba = '" & session("fcBA") & "'"

											set oRecordSet = Server.CreateObject("ADODB.Recordset")
											oRecordSet.Open oQuery, connstr
											If Not oRecordSet.eof Then
												nCurrentRNAmount = cdbl(oRecordSet("amountreceived"))
											Else
												nCurrentRNAmount = 0
											End If
											oRecordSet.Close()
											set oRecordSet = Nothing
											
											
											Selisih = (nFixedPOAmount - nCurrentRNAmount)
											txtamountreceived = Selisih
											
											
											q = " update purchaseorderdetails" & _
												" set rowstateid = 'Completed'" & _
												" where pocode = '" & xCode & "'" & _
												"       and linenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
												"       and fcba = '" & session("fcBA") & "'"
												
											'response.Write q
											connStr.execute(q)
										End If
										
										q="Insert into  RECEIVENOTEDETAIL (RECEIVENOTECODE, LOCATIONTYPE, LOCATIONCODE, PURCHASEITEMCODE, " & _
										"JOBCODE, QUANTITY, QTYRETURN, AMOUNTRECEIVED, AMOUNTRETURN, FCENTRY, FCEDIT, FCIP, FCBA, linenohist, purchaselinenokey, LASTUPDATE, LASTTIME) " & _
										"values ('" & fcCode & "','" & request("LT_"&TRIM(n)) & "','" & request("LC_"&TRIM(n)) &  _
										"','" & request("Stock_"&TRIM(n)) & "','" & txtControlJOb & "','" & request("QtyReceived_"&TRIM(n)) & _
										"',0,'"& txtamountreceived & "',0," & _
										"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
										Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',"& tempMaxLine & ",'" & request("LineNoKey_"&TRIM(n)) & "', sysdate, To_Char(sysdate,'HH24:MI'))"
										'response.Write(q)
									
										connStr.execute(q)
									else
									
										q="Insert into  RECEIVENOTEDETAIL (RECEIVENOTECODE, LOCATIONTYPE, LOCATIONCODE, PURCHASEITEMCODE, " & _
											"JOBCODE, QUANTITY, QTYRETURN, AMOUNTRECEIVED, AMOUNTRETURN, FCENTRY, FCEDIT, FCIP, FCBA, purchaselinenokey, LASTUPDATE, LASTTIME) " & _
											"values ('" & fcCode & "','" & request("LT_"&TRIM(n)) & "','" & request("LC_"&TRIM(n)) &  _
											"','" & request("Stock_"&TRIM(n)) & "','" & txtControlJOb & "','" & request("QtyReceived_"&TRIM(n)) & _
											"',0,'"& txtamountreceived & "',0," & _
											"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
											Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','" & request("LineNoKey_"&TRIM(n)) & "', sysdate, To_Char(sysdate,'HH24:MI'))"
											'response.Write(q)
										
										connStr.execute(q)
										
										
									end if

									' Add by NAF to handle if condition, except MATERIAL
									if tempItem = "MATERIAL" then 										
										If request("LT_"&TRIM(n)) = "MN" Then
										
											oQuery = " select count(*) total" & _
													 " from stockonhandnursery" & _
													 " where storecode = '" & xstoreCode & "'" & _
													 "       and stockcode = '" & request("Stock_"&TRIM(n)) & "'" & _
													 "       and nurserycode = '" & request("LC_"&TRIM(n)) & "'" & _
													 "       and fcba = '" &  session("fcBA") & "'"

											set oRecordSet = Server.CreateObject("ADODB.Recordset")
											oRecordSet.Open oQuery, connstr
											if not oRecordSet.eof then
												nTotalRecordOnStockOnHandNursery = cdbl(oRecordSet("total"))
											else
												nTotalRecordOnStockOnHandNursery = 0
											end if
											oRecordSet.Close()
											set oRecordSet = Nothing
											
											If nTotalRecordOnStockOnHandNursery = 0 then
											
												' Just Insert the New Record with the following/received Quantity
												q = " Insert into stockonhandnursery" & _
													" ( storecode, stockcode, nurserycode, fcba, quantity, lastupdate, lasttime ) values (" & _
													" '" & xstoreCode & "'" & _
													" , '" & request("Stock_"&TRIM(n)) & "'" & _
													" , '" & request("LC_"&TRIM(n)) & "'" & _
													" , '" & session("fcBA") & "'" & _
													" , '" & cdbl(request("QtyReceived_"&TRIM(n))) & "'" & _
													" , sysdate" & _
													" , to_char(sysdate,'hh24:mi')" & _
													" )"

												connStr.execute(q)

											Else
											
												q = " update stockonhandnursery " & _
													" set quantity = quantity + " & cdbl(request("QtyReceived_"&TRIM(n))) & " " & _
													"     , LASTUPDATE = sysdate " & _
													"     , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
													" where storecode = '" & xstoreCode & "'" & _
													"       and stockcode  = '" & request("Stock_"&TRIM(n)) & "'" & _
													"       and nurserycode  = '" & request("LC_"&TRIM(n)) & "'" & _
													"       and fcba = '" & session("fcBA") &"'"

												connStr.execute(q)
											End If
										End If

										qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
															 "where " & _
															 "storecode = '" & xstoreCode & "' and " & _												 
															 "stockcode = '" & request("Stock_"&TRIM(n)) & "' and " & _
															 "fcba = '" &  session("fcBA") & "'"

										set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
										rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
										if not rsTotalRecordOnHand.eof then
											tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
										else
											tempTotalRecordOnHand = 0
										end if 
										
										if tempTotalRecordOnHand = 0 then
										
											' Just Insert the New Record with the following/received Quantity
											q = "Insert into stockonhand ( " & _
												" STORECODE, STOCKCODE, FCBA, QUANTITY, LASTUPDATE, LASTTIME ) values ( '" & _
												xstoreCode & "','" & _
												request("Stock_"&TRIM(n)) & "','" & _
												session("fcBA") & "', '" & cdbl(request("QtyReceived_"&TRIM(n))) & "', sysdate, To_Char(sysdate,'HH24:MI') ) "
												
												
											connStr.execute(q)
											
										Else	  

											q = "update stockonhand " & _
												"set quantity = quantity + " & cdbl(request("QtyReceived_"&TRIM(n))) & " " & _ 
												", LASTUPDATE = sysdate " & _
												", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
												"where " & _
												"storecode = '" & xstoreCode & "' and " & _
												"stockcode  = '" & request("Stock_"&TRIM(n)) & "' and " & _ 
												"fcba = '" & session("fcBA") &"' "
												 
											connStr.execute(q)
										End If
									End if
								End If
							End If
						Next
						
						if err.number=0 then
							connStr.CommitTrans
						End if
						
					else
						err.number = 1001
						err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
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
 q="select * from RECEIVENOTE  " & _
  " where RECEIVENOTECODE='" & fcCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
 SudahClose="N"
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    DATECREATED=rsVoucher("RNDATE")
	xCode=rsVoucher("POCODE")
	xstoreCode = rsVoucher("STORECODE")
	SudahClose="Y"
 End if
 rsVoucher.Close
 Set rsVoucher = Nothing
 
If Err.Number <> 0 then
	connStr.RollbackTrans
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if


Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & TRANSDATE & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & TRANSDATE & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
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
 <form name="frmEntry" method="post" action="supplyingdeclarationfinal.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Document No</td>
			<td align="left"><input type="text" maxlength="50" name="txtDocumentNo" size="30" value="<%=txtDocumentNo%>">
			<font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
			<td align="left">Document Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="transdate" size="10" value="<%=transdate%>">
			<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left"> 
				<input Type="Hidden" Value="True" Name="cmdView">  
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">  
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">   	 
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;  			
				<%
					dim dataModule,dataSubmodule,dataDetail
					dataModule 		= "Purchase"
					dataSubmodule 	= "Receive Note"
					dataDetail 		= "NULL"
				%>
		
		<!-- END:HADIR -->
		</td> 
		</tr> 
    </table>    
 </form>
<%
If cmdView = "True" and LEN(TRIM(txtDocumentNo)) > 0 and len(trim(transdate)) > 0 Then
	Dim rsGRN,intRecordsPerPage,cColor
	set rsGRN = server.CreateObject("Adodb.recordset")
	If SudahClose = "Y" Then
		' q	= " select purchaseitemcode" & _
		'       "        , purchaselinenokey purchaselinenokey" & _
		' 	  "        , linenohist" & _
		' 	  "        , jobcode" & _
		' 	  "        , locationtype" & _
		' 	  "        , locationcode" & _
		' 	  "        , nvl(sum(nvl(quantity, 0)), 0) as quantity" & _
		' 	  "        , nvl(sum(nvl(amountreceived, 0)), 0) as amountreceived" & _			  
		' 	  " from receivenotedetail p" & _
		' 	  "      , stock s" & _
		' 	  " where (p.purchaseitemcode = s.fccode and p.fcba = s.fcba)" & _
		' 	  "       and receivenotecode = '" & fcCode & "'" & _
		' 	  "       and p.fcBA = '" & session("fcBA") & "'" & _
		' 	  " group by purchaseitemcode" & _
		' 	  "          , purchaselinenokey" & _
		' 	  "          , linenohist" & _
		' 	  "          , jobcode" & _
		' 	  "          , locationtype" & _
		' 	  "          , locationcode" & _
		' 	  " order by purchaseitemcode" & _
		' 	  "          , purchaselinenokey"
	Else	 
		q	= " SELECT sdd.documentno, sdh.transdate, sdd.declarationtype, sdd.fieldcode, sdd.quantity, sdd.remarks ,pr.FCNAME as declarationtypename" & _
		  " FROM supplyingdeclarationdtl sdd" & _
		  " INNER JOIN supplyingdeclarationhdr sdh"& _
		  " ON (sdd.documentno = sdh.documentno AND sdd.fcba = sdh.fcba)" & _
		  " INNER JOIN parameter pr ON (sdd.fcba = pr.fcba AND sdd.declarationtype = pr.fccode AND pr.parcode = 'AGROTECH_SUPPLYINGTYPE')" & _
		  " WHERE sdd.fcba = '" & session("fcBA") & "'" & _
		  " AND sdd.documentno = '" & txtDocumentNo & "'"
   End if
   'response.Write q	
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.CursorType = 3
   rsGRN.CursorLocation = 3
   rsGRN.Open q, connstr%>
   <form action="supplyingdeclarationfinal.Asp" method="post" name="eForm">
   <table class="data" border="1" width="100%" >
   <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">Declaration Type</td>
	   <td align="center">Field Code</td>
	   <td align="center">Remarks</td>
	   <td align="center">Total Stand (Qty)</td>
	   <td align="center">Approval Quantity</td>
   	   </tr>
   <%Else%>
 	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">Declaration Type</td>
	   <td align="center">Field Code</td>
	   <td align="center">Remarks</td>
	   <td align="center">Total Stand (Qty)</td>
	   <td align="center">Approval Quantity</td>
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
				<%IF SudahClose = "Y" then%>
					<%If val_IsTrue = "Y" Then%>
					<td valign="middle">
						<a href="supplyingdeclarationfinal.Asp?fieldcode=<%=fieldcode%>&txtdocumentno=<%=rsGRN("documentno")%>&purchaseLineNoKey=<%=rsGRN("PurchaseLineNoKey")%>&linenomorhist=<%=rsGRN("linenohist")%>&DATECREATED=<%=DATECREATED%>&txtJenis=Details">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGRN(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
					</td>
					<%Else%>
					<td>
						<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
					</td>
					<%End If%>
					<%for i = 0 to nCountofField-1%>
						<td align="left" title="<%=rsGRN(i).Name%>" valign="top"><%=rsGRN(i)%></td>
					<%next%>
				<%End if%>
				
				<%if SudahClose<>"Y" Then%>
					<td align="Right" title="Declaration Type" valign="top"><%=FbpNum2(rsGRN("declarationtype"))%></td>
					<td align="left" title="Fieldcode" valign="top"><%=rsGRN("fieldcode")%></td>
					<td align="Right" title="Remarks" valign="top"><%=FbpNum2(rsGRN("remarks"))%></td>
					<td align="left" title="Total Stand" valign="top"><%=rsGRN("quantity")%></td>
					<td align="left" width="2">
		        		<input type="text" name="quantity_<%=no%>" width="2" value="<%=rsGRN("quantity")%>">
		        	</td>
				<%End if%>
			</tr>
		  <%
		  rsGRN.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsGRN.Close
   set rsGRN = Nothing%>
	  <tr><td colspan="11" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="xstoreCode" value="<%=xstoreCode%>">
		  <input type="hidden" name="No" value="<%=No%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%If val_IsTrue = "Y" Then%>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Submit" onClick="return confirm('Apakah Anda sudah Cek jumlah dan qualitas dari tiap Barang yg diterima ?');">
		  <%End if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
     
   </TABLE>
    </div>
   </form>
<%
End if  
 
'-------------akhir----------------------
  %>
 </div> </div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
