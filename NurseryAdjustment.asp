<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Nursery Adjustment</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
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
 <!-------------awal!------------->
 <%


Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, LineNoHist, ITEMCODE
DIM v_NURSERYCODE,v_CULLINGDATE
DIM v_QUANTITY
Dim cmdView
Dim qNurseryInfo, rsNurseryInfo, tempStoreCode, tempStockCode, doneError
Dim qLineMax, rsqLineMax, tempMaxLine
Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
Dim qLessZero, rsLessZero, tempislesszero
Dim done, qLineUpper, rsqLineUpper, totalUpper, qQuantity, rsqQuantity, totalQuantity, QuantityOriginal
Dim qDeleteHist
Dim rsReceiveDetail, originalQuantity, v_CrossCekDetail

Dim QuantityOnHandNursery : QuantityOnHandNursery = 0

Dim txtDocumentNo : txtDocumentNo = Request("txtDocumentNo")
Dim txtAdjustmentDate : txtAdjustmentDate = Request("txtAdjustmentDate")
Dim txtNurseryCode : txtNurseryCode = Request("txtNurseryCode")
Dim txtLineNoHist : txtLineNoHist = Request("txtLineNoHist")
Dim txtQuantity : txtQuantity = Request("txtQuantity")
Dim txtReference : txtReference = Request("txtReference")
Dim txtRemarks : txtRemarks = Request("txtRemarks")
Dim rsMaxLineNoKey, strMaxLineNoKey, tempMaxLineNoKey
Dim txtLineNoKey : txtLineNoKey = Request("txtLineNoKey")
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

Dim CekLineHistNursery : CekLineHistNursery = 1

Function CheckMasterNursery(initialNurseryCode, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from nursery" & _
					" where nurserycode = '" & initialNurseryCode & "'" & _
					"       and fcba = '" & initialFcba & "'"
							
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

	CheckMasterNursery = initialAvailable
End Function


Sub InsertHeader
	q	=	" insert into nurseryadjhdr" & _
			" (documentno, adjustmentdate, fcentry, fcedit, fcip, fcba, lastupdate, lasttime)" & _
			" values (" & _
			" '" & txtDocumentNo & "'" & _
			" , to_date('" & txtAdjustmentDate & "', 'mm/dd/yyyy')" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			" , '" & Session("fcBA") & "'" & _
			" , sysdate" & _
			" , to_char(sysdate, 'hh24:mi')" & _
			" )"
			
	'Response.Write q
	connStr.execute(q)
End Sub

Sub InsertDetail

	strMaxLineNoKey = " select nvl(max(linenokey),0) tempMaxLineNoKey" & _
					  " from nurseryadjdtl" & _
					  " where fcba = '" & session("fcBA") & "'" 
		
	set rsMaxLineNoKey = Server.CreateObject("ADODB.Recordset")
	rsMaxLineNoKey.Open strMaxLineNoKey, connstr
	if not rsMaxLineNoKey.eof then
		tempMaxLineNoKey = cdbl(rsMaxLineNoKey("tempMaxLineNoKey"))
	else
		tempMaxLineNoKey = 0
	end if 

	tempMaxLineNoKey = tempMaxLineNoKey + 1
	
	q	=	" insert into nurseryadjdtl" & _
			" (documentno, linenokey, nurserycode, quantity, reference, remarks, fcentry, fcedit, fcip, fcba, linenohist, lastupdate, lasttime)" & _
			" values (" & _
			" '" & txtDocumentNo & "'" & _
			" , '" & tempMaxLineNoKey & "'" & _
			" , '" & txtNurseryCode & "'" & _
			" , '" & txtQuantity & "'" & _
			" , '" & txtReference & "'" & _
			" , '" & txtRemarks & "'" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Session("CodeName") & "'" & _
			" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			" , '" & Session("fcBA") & "'" & _
			" , '" & txtLineNoHist & "'" & _
			" , sysdate" & _
			" , to_char(sysdate, 'hh24:mi')" & _
			" )"

	'Response.Write q
	connStr.execute(q)
End Sub

Sub UpdateDetail
	q	=	" update nurseryadjdtl" & _
			" set quantity = '" & txtQuantity & "'" & _
			"     , reference = '" & txtReference & "'" & _
			"     , remarks = '" & txtRemarks & "'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , lastupdate = sysdate" & _
			"     , lasttime = to_char(sysdate, 'hh24:mi')" & _
	        " where documentno = '" & txtDocumentNo & "'" & _
			"       and nurserycode = '" & txtNurseryCode & "'" & _
			"       and linenokey = '" & txtLineNoKey & "'" & _
			"       and fcba = '" & Session("fcBA") & "'"

	'Response.Write q
	connStr.execute(q)
End Sub

Sub DeleteDetail
	q	=	" delete from nurseryadjdtl" & _
	        " where documentno = '" & txtDocumentNo & "'" & _
			"       and linenokey = '" & txtLineNoKey & "'" & _
			"       and fcba = '" & Session("fcBA") & "'"

	'Response.Write q
	connStr.execute(q)
End Sub

Sub DeleteHeader
	q	=	" delete from nurseryadjhdr" & _
	        " where documentno = '" & txtDocumentNo & "'" & _
			"       and fcba = '" & Session("fcBA") & "'"

	'Response.Write q
	connStr.execute(q)
End Sub

Function cmdCheckHistNursery(initfcba, initNURSERYCODE, initLineNoHist)
	dim qCheckHistNursery,rsCheckHistNursery,qTotalDataNursery
	
	qCheckHistNursery = " select count(1) totalData" & _
						" from stocktranshist" & _
						" where fcba = '" & initfcba & "'" & _
						" and sourcecode = '" & initNURSERYCODE & "'" & _
						" and lineno > " & initLineNoHist & ""

	set rsCheckHistNursery = Server.CreateObject("ADODB.Recordset")
	rsCheckHistNursery.Open qCheckHistNursery, connstr
	if not rsCheckHistNursery.eof then
		qTotalDataNursery = cdbl(rsCheckHistNursery("totalData"))
	else
		qTotalDataNursery = 0
	end if 	

	cmdCheckHistNursery = qTotalDataNursery
End Function


If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), txtAdjustmentDate)
		If Cek_IsDailyClosing = "FALSE" then

			DeleteHeader
			
			if err.number = 0 then
			   connStr.CommitTrans
			end if		
		else
			err.number = 1001
			err.description = "Tanggal """ & txtAdjustmentDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
	
	elseif Request("txtJenis")= "Delete" then
	    connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), txtAdjustmentDate)
		If Cek_IsDailyClosing = "FALSE" then
			qNurseryInfo =  " select storecode" & _
							"        , stockcode" & _
							" from nursery" & _
							" where nurserycode = '" & txtNurseryCode & "'" & _
							"       and fcba = '"& Session("fcBA") & "'"
							
			'Response.Write qNurseryInfo
					
			set rsNurseryInfo = Server.CreateObject("ADODB.Recordset")
			rsNurseryInfo.Open qNurseryInfo, connstr
			if not rsNurseryInfo.eof then
				tempStoreCode = rsNurseryInfo("storecode")
				tempStockCode = rsNurseryInfo("stockcode")				
			else
				tempStoreCode = ""
				tempStockCode = ""
			end if 
			
			If tempStoreCode = "" Then
				err.number = 1001
				err.description = "Storecode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
			Else
				If tempStockCode = "" Then
					err.number = 1001
					err.description = "Stockcode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
				Else
					q =	" select linenohist" & _
						"        , quantity*-1 quantity" & _
						" from nurseryadjdtl" & _
						" where documentno = '" & txtDocumentNo & "'" & _
						"       and nurserycode = '" & txtNurseryCode & "'" & _
						"       and linenokey = '" & txtLineNoKey & "'" & _
						"       and fcba = '" & Session("fcBA") & "'"
					'Response.Write q
					Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
					rsReceiveDetail.Open q, connstr
					
					IF Not rsReceiveDetail.eof then
						txtLineNoHist = rsReceiveDetail("LINENOHIST")
						originalQuantity = Cdbl(rsReceiveDetail("QUANTITY"))
					End if
					rsReceiveDetail.Close
					set rsReceiveDetail = Nothing
					
					' xxxxxxxxxxxxxxxxxxxxxxx
					CekLineHistNursery = cmdCheckHistNursery(Session("fcBA"), txtNurseryCode, txtLineNoHist)
					
						qLessZero = " select islesszero" & _
									" from stock" & _
									" where fccode = '" & tempStockCode & "'" & _
									"       and fcba = '" & session("fcBA") & "'" 
						
						set rsLessZero = Server.CreateObject("ADODB.Recordset")
						rsLessZero.Open qLessZero, connstr
						if not rsLessZero.eof then
							tempislesszero = rsLessZero("islesszero")
						else
							tempislesszero = "FALSE"
						end if
						
						qQuantityOnHand = " select quantity" & _
										   " from stockonhandnursery" & _
										   " where storecode = '" & tempStoreCode & "'" & _
										   "       and stockcode  = '" & tempStockCode & "'" & _
										   "       and nurserycode  = '" & txtNurseryCode & "'" & _
										   "       and fcba = '" & session("fcBA") &"' "
										  
						set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
						rsQuantityOnHand.Open qQuantityOnHand, connstr
						if not rsQuantityOnHand.eof then
							QuantityOnHandNursery = cdbl(rsQuantityOnHand("quantity"))
						else
							QuantityOnHandNursery = 0
						end if
						
						If tempislesszero = "TRUE" Or (tempislesszero = "FALSE" And ((QuantityOnHandNursery + originalQuantity) >= 0)) Then
						
							q =	" Delete from stocktranshist" & _
								" where lineno = " & txtLineNoHist & "" & _
								"       and fcba = '" & Session("fcBA") & "'"

							connStr.execute(q)

							q = " update stockonhand" & _
								" set quantity = quantity + " & originalQuantity & "" & _
								"     , lastupdate = sysdate " & _
								"     , lasttime = to_char(sysdate,'hh24:mi') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode = '" & tempStockCode & "'" & _
								"       and fcba = '" & session("fcBA") &"'"

							connStr.execute(q)

							q = " update stockonhandnursery " & _
								" set quantity = quantity + " & originalQuantity & "" & _
								"     , lastupdate = sysdate " & _
								"     , lasttime = to_char(sysdate,'hh24:mi') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode  = '" & tempStockCode & "'" & _ 
								"       and nurserycode  = '" & txtNurseryCode & "'" & _ 
								"       and fcba = '" & session("fcBA") &"'"

							connStr.execute(q)
							
							DeleteDetail
						Else
							err.number = 1001
							Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
						End If
				End If
			End If

			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & txtAdjustmentDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
        cmdView	= "True"

	elseif Request("txtJenis")="New" then
	
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), txtAdjustmentDate)
		If Cek_IsDailyClosing = "FALSE" then
		
			Dim oRS, oSQL
			Set oRS = Server.CreateObject("ADODB.Recordset")
			oSQL	=	" select documentno" & _
						" from nurseryadjhdr" & _
						" where documentno = '" & txtDocumentNo & "'" & _
						"       and fcba = '" & Session("FCBA") & "'"
			oRS.Open oSQL, connstr
			
			IF oRS.eof then	    
				InsertHeader
			End If
			oRS.Close
			set oRS = Nothing
			
			If Len(Trim(txtNurseryCode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Kode Batch Harus diisi, tidak boleh kosong"
			Else
				Dim val_NurseryCode : val_NurseryCode = 0
				val_NurseryCode = CheckMasterNursery(txtNurseryCode, session("fcBA"))
				If val_NurseryCode = 0 Then
					Err.Number = 1001
					Err.Description = "Kode Batch """ & txtNurseryCode & """ tidak ada di master Nursery iPlaS, Koreksi kembali Kode Batch yang anda masukkan."
				End If
			End If
			
			If val_NurseryCode > 0 Then
				qNurseryInfo = 	" select storecode" & _
								"        , stockcode" & _
								" from nursery" & _
								" where nurserycode = '" &  txtNurseryCode & "'" & _
								"       and fcba = '"& Session("fcBA") & "'"
					
				set rsNurseryInfo = Server.CreateObject("ADODB.Recordset")
				rsNurseryInfo.Open qNurseryInfo, connstr
				if not rsNurseryInfo.eof then
					tempStoreCode = rsNurseryInfo("storecode")
					tempStockCode = rsNurseryInfo("stockcode")
				else
					tempStoreCode = ""
					tempStockCode = ""
				end if 
				
				If tempStoreCode = "" Then
					err.number = 1001
					err.description = "Storecode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
				Else
					If tempStockCode = "" Then
						err.number = 1001
						err.description = "Stockcode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
					Else
						
						qLessZero =	" select islesszero" & _
									" from stock" & _
									" where fccode = '" & tempStockCode & "'" & _
									"       and fcba = '" & session("fcBA") & "'" 
						
						set rsLessZero = Server.CreateObject("ADODB.Recordset")
						rsLessZero.Open qLessZero, connstr
						if not rsLessZero.eof then
							tempislesszero = rsLessZero("islesszero")
						else
							tempislesszero = "FALSE"
						end if
						
						qQuantityOnHand =	" select quantity" & _
											" from stockonhandnursery" & _
											" where storecode = '" & tempStoreCode & "'" & _
											"       and stockcode = '" & tempStockCode & "'" & _
											"       and nurserycode = '" & txtNurseryCode & "'" & _
											"       and fcba = '" &  session("fcBA") & "'"
										  
						set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
						rsQuantityOnHand.Open qQuantityOnHand, connstr
						if not rsQuantityOnHand.eof then
							tempQuantityOnHand = cdbl(rsQuantityOnHand("quantity"))
						else
							tempQuantityOnHand = 0
						end if
						
						' We have to Update the Quantity by Adding them with the New one.
						If (((tempQuantityOnHand + txtQuantity) >= 0) And tempislesszero = "FALSE") or tempislesszero = "TRUE"  Then
							qLineMax = " select nvl(max(lineno),0) tempMaxLine" & _
									   " from stocktranshist" & _
									   " where fcba = '" & session("fcBA") & "'" 
								
							set rsqLineMax = Server.CreateObject("ADODB.Recordset")
							rsqLineMax.Open qLineMax, connstr
							if not rsqLineMax.eof then
								tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
							else
								tempMaxLine = 0
							end if 

							tempMaxLine = tempMaxLine + 1
							txtLineNoHist = tempMaxLine
							
							q = " insert into stocktranshist" & _
								" (lineno, fcba, stockcode, transdate, transtype, storecode, sourcetype" & _
								" , sourcecode, targettype, targetcode, quantity, lastupdate, lasttime)" & _
								" values ("& _
								"" & txtLineNoHist & "" & _
								" , '" & session("fcBA") & "'" & _
								" , '" & tempStockCode &"'" & _
								" , to_date('" & txtAdjustmentDate & "','mm/dd/yyyy')" & _
								" , 'MN_ADJUSTMENT'" & _
								" , '" & tempStoreCode & "'" & _
								" , 'MN'" & _
								" , '" & txtNurseryCode & "'" & _
								" , 'MN'" & _
								" , '" & txtNurseryCode & "'" & _
								" , " & txtQuantity & "" & _
								" , sysdate" & _
								" , to_char(sysdate,'hh24:mi')" & _
								" )"
								  
							connStr.execute(q)
							
							InsertDetail
							
							'Lihat Record di tabel stockonhand
							qTotalRecordOnHand =	" select count(*) totalRecordOnHand" & _
													" from stockonhand" & _
													" where storecode = '" & tempStoreCode & "'" & _
													"       and stockcode = '" & tempStockCode & "'" & _
													"       and fcba = '" &  session("fcBA") & "'"

							set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
							rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
							if not rsTotalRecordOnHand.eof then
								tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
							else
								tempTotalRecordOnHand = 0
							end if
							
							'Jika tidak ada
							If tempTotalRecordOnHand = 0 Then
								' Just Insert the New Record with the following/received Quantity
								q	=	" insert into stockonhand" & _
										" (storecode, stockcode, fcba, quantity, lastupdate, lasttime)" & _
										" values (" & _
										" '" & tempStoreCode & "'" & _
										" , '" & tempStockCode & "'" & _
										" , '" & session("fcBA") & "'" & _
										" , '" & txtQuantity & "'" & _
										" , sysdate" & _
										" , To_Char(sysdate,'HH24:MI')" & _
										" )"

								connStr.execute(q)
							Else
							
								q = " update stockonhand" & _
									" set quantity = quantity + " & txtQuantity & "" & _
									"     , lastupdate = sysdate " & _
									"     , lasttime = to_char(sysdate,'hh24:mi') " & _
									" where storecode = '" & tempStoreCode & "'" & _
									"       and stockcode = '" & tempStockCode & "'" & _
									"       and fcba = '" &  session("fcBA") & "'"
									 
								connStr.execute(q)
							End If
							
							'Lihat Record di tabel stockonhandnursery
							qTotalRecordOnHand =	" select count(*) totalRecordOnHand" & _
													" from stockonhandnursery" & _
													" where storecode = '" & tempStoreCode & "'" & _
													"       and stockcode = '" & tempStockCode & "'" & _
													"       and nurserycode = '" & txtNurseryCode & "'" & _
													"       and fcba = '" &  session("fcBA") & "'"

							set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
							rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
							if not rsTotalRecordOnHand.eof then
								tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
							else
								tempTotalRecordOnHand = 0
							end if
							
							'Jika tidak ada
							If tempTotalRecordOnHand = 0 Then
								' Just Insert the New Record with the following/received Quantity
								q	=	" insert into stockonhandnursery" & _
										" (storecode, stockcode, nurserycode, fcba, quantity, lastupdate, lasttime)" & _
										" values (" & _
										" '" & tempStoreCode & "'" & _
										" , '" & tempStockCode & "'" & _
										" , '" & txtNurseryCode & "'" & _
										" , '" & session("fcBA") & "'" & _
										" , '" & txtQuantity & "'" & _
										" , sysdate" & _
										" , To_Char(sysdate,'HH24:MI')" & _
										" )"

								connStr.execute(q)
							Else
							
								q = " update stockonhandnursery" & _
									" set quantity = quantity + " & txtQuantity & "" & _
									"     , lastupdate = sysdate " & _
									"     , lasttime = to_char(sysdate,'hh24:mi') " & _
									" where storecode = '" & tempStoreCode & "'" & _
									"       and stockcode = '" & tempStockCode & "'" & _
									"       and nurserycode = '" & txtNurseryCode & "'" & _
									"       and fcba = '" &  session("fcBA") & "'"
									 
								connStr.execute(q)
							End If
						Else
							err.number = 1001
							Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
						End If
					End If
				End If
			End If
			
			If err.number = 0 Then
				connStr.CommitTrans
			End If
		else
			err.number = 1001
			err.description = "Tanggal """ & txtAdjustmentDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	    connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), txtAdjustmentDate)
		If Cek_IsDailyClosing = "FALSE" then

			qNurseryInfo =  " select storecode" & _
							"        , stockcode" & _
							" from nursery" & _
							" where nurserycode = '" & txtNurseryCode & "'" & _
							"       and fcba = '"& Session("fcBA") & "'"
							
			'Response.Write qNurseryInfo
					
			set rsNurseryInfo = Server.CreateObject("ADODB.Recordset")
			rsNurseryInfo.Open qNurseryInfo, connstr
			if not rsNurseryInfo.eof then
				tempStoreCode = rsNurseryInfo("storecode")
				tempStockCode = rsNurseryInfo("stockcode")				
			else
				tempStoreCode = ""
				tempStockCode = ""
			end if 
			
			If tempStoreCode = "" Then
				err.number = 1001
				err.description = "Storecode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
			Else
				If tempStockCode = "" Then
					err.number = 1001
					err.description = "Stockcode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
				Else
					q =	" select linenohist" & _
						"        , quantity*-1 quantity" & _
						" from nurseryadjdtl" & _
						" where documentno = '" & txtDocumentNo & "'" & _
						"       and nurserycode = '" & txtNurseryCode & "'" & _
						"       and linenokey = '" & txtLineNoKey & "'" & _
						"       and fcba = '" & Session("fcBA") & "'"
					'Response.Write q
					Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
					rsReceiveDetail.Open q, connstr
					
					IF Not rsReceiveDetail.eof then
						txtLineNoHist = rsReceiveDetail("LINENOHIST")
						originalQuantity = Cdbl(rsReceiveDetail("QUANTITY"))
					End if
					rsReceiveDetail.Close
					set rsReceiveDetail = Nothing
					
					' xxxxxxxxxxxxxxxxxxxxxxx
					CekLineHistNursery = cmdCheckHistNursery(Session("fcBA"), txtNurseryCode, txtLineNoHist)

					
						qLessZero = " select islesszero" & _
									" from stock" & _
									" where fccode = '" & tempStockCode & "'" & _
									"       and fcba = '" & session("fcBA") & "'" 
						
						set rsLessZero = Server.CreateObject("ADODB.Recordset")
						rsLessZero.Open qLessZero, connstr
						if not rsLessZero.eof then
							tempislesszero = rsLessZero("islesszero")
						else
							tempislesszero = "FALSE"
						end if
						
						qQuantityOnHand = " select quantity" & _
										   " from stockonhandnursery" & _
										   " where storecode = '" & tempStoreCode & "'" & _
										   "       and stockcode  = '" & tempStockCode & "'" & _
										   "       and nurserycode  = '" & txtNurseryCode & "'" & _
										   "       and fcba = '" & session("fcBA") &"' "
										  
						set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
						rsQuantityOnHand.Open qQuantityOnHand, connstr
						if not rsQuantityOnHand.eof then
							QuantityOnHandNursery = cdbl(rsQuantityOnHand("quantity"))
						else
							QuantityOnHandNursery = 0
						end if
						
						If tempislesszero = "TRUE" Or (tempislesszero = "FALSE" And ((QuantityOnHandNursery + originalQuantity + txtQuantity) >= 0)) Then
						
							q =	" Delete from stocktranshist" & _
								" where lineno = " & txtLineNoHist & "" & _
								"       and fcba = '" & Session("fcBA") & "'"

							connStr.execute(q)
							
							q = " insert into stocktranshist" & _
								" (lineno, fcba, stockcode, transdate, transtype, storecode, sourcetype" & _
								" , sourcecode, targettype, targetcode, quantity, lastupdate, lasttime)" & _
								" values ("& _
								"" & txtLineNoHist & "" & _
								" , '" & session("fcBA") & "'" & _
								" , '" & tempStockCode &"'" & _
								" , to_date('" & txtAdjustmentDate & "','mm/dd/yyyy')" & _
								" , 'MN_ADJUSTMENT'" & _
								" , '" & tempStoreCode & "'" & _
								" , 'MN'" & _
								" , '" & txtNurseryCode & "'" & _
								" , 'MN'" & _
								" , '" & txtNurseryCode & "'" & _
								" , " & txtQuantity & "" & _
								" , sysdate" & _
								" , to_char(sysdate,'hh24:mi')" & _
								" )"
								  
							connStr.execute(q)

							q = " update stockonhand" & _
								" set quantity = quantity + " & originalQuantity + txtQuantity & "" & _
								"     , lastupdate = sysdate " & _
								"     , lasttime = to_char(sysdate,'hh24:mi') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode = '" & tempStockCode & "'" & _
								"       and fcba = '" & session("fcBA") &"'"

							connStr.execute(q)

							q = " update stockonhandnursery " & _
								" set quantity = quantity + " & originalQuantity + txtQuantity & "" & _
								"     , lastupdate = sysdate " & _
								"     , lasttime = to_char(sysdate,'hh24:mi') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode  = '" & tempStockCode & "'" & _ 
								"       and nurserycode  = '" & txtNurseryCode & "'" & _ 
								"       and fcba = '" & session("fcBA") &"'"

							connStr.execute(q)
							
							UpdateDetail
						Else
							err.number = 1001
							Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
						End If
				End If
			End If

			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & txtAdjustmentDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
Else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

If Len(Trim(txtAdjustmentDate)) = 0 Then
	txtAdjustmentDate = textDate(date)
Else
	txtAdjustmentDate = Request("txtAdjustmentDate")
End If

Dim rsNurseryAdjustment, strNurseryAdjustment
strNurseryAdjustment	=	" select documentno, to_char(adjustmentdate,'mm/dd/yyyy') adjustmentdate" & _
							" from nurseryadjhdr" & _
							" where documentno = '" & txtDocumentNo & "'" & _
							"       and fcba = '" & session("fcBA") & "'"
						
set rsNurseryAdjustment = server.CreateObject("Adodb.recordset")
rsNurseryAdjustment.Open strNurseryAdjustment, connstr
if Not rsNurseryAdjustment.Eof then
	txtDocumentNo	=	rsNurseryAdjustment("documentno")
	txtAdjustmentDate	=	rsNurseryAdjustment("adjustmentdate")	
End if
'response.Write strNurseryAdjustment
rsNurseryAdjustment.Close
Set rsNurseryAdjustment = Nothing

If Err.Number <> 0 then
	connStr.RollbackTrans
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & txtAdjustmentDate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & txtAdjustmentDate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Nursery')"
	   
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

 <!-------------ISI Payment Form!------------->
 
	<form name="frmEntry" method="post" action="NurseryAdjustment.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Adjustment No</td>
			<td align="left">
				<input type="text" maxlength="15" name="txtDocumentNo" size="30" value="<%=txtDocumentNo%>">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Adjustment Date</td>
			<td align="left"><input class="datepicker" type="text" size="15" name="txtAdjustmentDate" value="<%=txtAdjustmentDate%>"></td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="#" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" Then%>
					<a href="NurseryAdjustment.Asp?txtDocumentNo=<%=txtDocumentNo%>&txtAdjustmentDate=<%=txtAdjustmentDate%>&txtJenis=HeaderDelete">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0" onClick="return confirm('Are you sure to delete');">
					</a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End If%>
				<!-- HADIR -->
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Nursery"
				dataSubmodule 	= "Adjustment"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td>
		</tr>
		</table>
	</form>
<%
if cmdView = "True" and len(trim(txtDocumentNo)) > 0 Then
   Dim rsVH,intRecordsPerPage,cColor
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	=	" select linenokey" & _
			"        , nurserycode" & _
			"        , quantity" & _
			"        , reference" & _
			"        , remarks" & _
			"        , linenohist" & _
			" from nurseryadjdtl" & _
			" where fcba = '" & session("fcBA") & "'" & _
			"       and documentno = '" & txtDocumentNo & "'"
			
'response.Write q
Set rsVH = Server.CreateObject("ADODB.Recordset")
rsVH.CursorType = 3
rsVH.CursorLocation = 3
rsVH.Open q, connstr
%>
   <form action="NurseryAdjustment.asp" method="post" name="eForm">
	   <table class="data" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td align="left" width="20%">Nursery Code</td>
				<td align="left">
					<%If Request("txtJenis") = "Edit" Then%>
					<input type="text" size="15" name="txtNurseryCode" value="<%=txtNurseryCode%>"  maxlength="15" readonly="True" style="background:#999999">
					<%Else%>
					<input type="text" size="15" name="txtNurseryCode" value="<%=txtNurseryCode%>" maxlength="15">
					<%End If%>
				</td>
			</tr>
			<%If Request("txtJenis") = "Edit" Then%>
				<input type="hidden" size="15" name="txtLineNoKey" value="<%=txtLineNoKey%>" readonly="True" style="background:#FFFFCC">
			<%End if%>
			<tr>
				<td align="left" width="20%">Quantity</td>
				<td align="left">
					<input type="text" size="15" name="txtQuantity" value="<%=txtQuantity%>" maxlength="32">
					<font color="Red"> (+) Adjust Positive | (-) Adjust Negative</font>
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Reference</td>
				<td align="left">
					<input type="text" size="50" name="txtReference" value="<%=txtReference%>" maxlength="50">
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Remarks</td>
				<td align="left">
					<textarea name="txtRemarks" cols="50" rows="3"maxlength="400"><%=txtRemarks%></textarea>
				</td>
			</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="txtDocumentNo" value="<%=txtDocumentNo%>">
					<input type="hidden" name="txtAdjustmentDate" value="<%=txtAdjustmentDate%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					<%If val_IsTrue = "Y" Then%>
					<%IF Request("txtJenis")="Edit" then%>
					   <input type="hidden" name="txtJenis" value="Edit">
					   <input Type="Submit"  name="Add" value="Update">
					<%Else%>
					   <input type="hidden" name="txtJenis" value="New">
					   <input Type="Submit"  name="Add" value="Add">
					<%End if%>
					<%End If%>
					<input Type="Reset"  name="Reset" value="Reset">
					<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
			  </td>
		  </tr>
	   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="2">&nbsp;</td>
		<td align="left">LINE NO (KEY)</td>
		<td align="left">NURSERY CODE</td>
		<td align="left">QUANTITY</td>
		<td align="left">REFERENCE</td>
		<td align="left">REMARKS</td>
		<td align="left">LINE NO (HISTORY)</td>
	</tr>
   
   <%
   if Not rsVH.eof then
      nCountofField = rsVH.fields.count - 1
	  no = 1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
			<tr bgcolor="<%=cColor%>">
				<td align="left" valign="top"><%=NO%></td>
				
				<%If val_IsTrue = "Y" Then%>
				<td valign="middle" width="30">
					<a href="NurseryAdjustment.asp?txtDocumentNo=<%=txtDocumentNo%>&txtNurseryCode=<%=rsVH("nurserycode")%>&txtLineNoKey=<%=rsVH("linenokey")%>&txtQuantity=<%=rsVH("quantity")%>&txtReference=<%=rsVH("reference")%>&txtRemarks=<%=rsVH("remarks")%>&txtJenis=Edit">
						<img src="IMAGES/edit.gif" alt="Edit Data : <%=txtNurseryCode%>" border="0">
					</a>
				</td>
				<td>
					<a href="NurseryAdjustment.asp?txtDocumentNo=<%=txtDocumentNo%>&txtNurseryCode=<%=rsVH("nurserycode")%>&txtLineNoKey=<%=rsVH("linenokey")%>&txtAdjustmentDate=<%=txtAdjustmentDate%>&txtJenis=Delete">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtNurseryCode%>" border="0" onClick="return confirm('Are you sure to delete');">
					</a>
				</td>
				<%Else%>
					<td colspan="2">
						<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
					</td>
				<%End If%>
			
				<%for i = 0 to nCountofField%>
			    <td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
				<%next%>	
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
