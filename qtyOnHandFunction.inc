

<%
DIM txtnamakolom,qryQtyOnHand1,rsQtyOnHand1,cmdQtyOnHand,txtnamakolom1,ENDQTY


function cmdInsertStockOnHand(txtStoreCode,txtItemCode,txtQty,txtKolomTransaksi)

	DIM txtnamakolomqty, txtnamakolomamount, txtModuleName
	IF txtKolomTransaksi ="RN" THEN
			txtModuleName 		= "ReceiveNote.asp"
	End If
	IF txtKolomTransaksi ="SIV" THEN
			txtModuleName 		= "IssueDetails.asp"
	End If
	IF txtKolomTransaksi ="TRT" THEN
			txtModuleName 		= "TransportTask.asp"
	End If
	IF txtKolomTransaksi ="RCV" THEN
			txtModuleName 		= "ReceivedTask.asp"
	End If

	IF txtKolomTransaksi ="TRTDIV" THEN
			txtModuleName 		= "TransportTaskDiv.asp"
	End If

	IF txtKolomTransaksi ="RT" THEN
			txtModuleName 		= "StockReturn.asp"
	End If
	IF txtKolomTransaksi ="WRT" THEN
			txtModuleName 		= "WriteOff.asp"
	End If

	qryQtyOnHand1 = "Insert into stockonhand ( " & _
		" STORECODE, STOCKCODE, FCBA, QUANTITY, LASTUPDATE, LASTTIME ) values ( '" & _
		txtStoreCode & "','" & _
		txtItemCode & "','" & _
		Session("fcBA") & "', '" & _ 
		txtQty & "', sysdate, To_Char(sysdate,'HH24:MI') ) "
	connStr.execute(qryQtyOnHand1) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

End Function 

'FUNCTION ini di proses ketika transaksi transporttask divisi yang kolom gudang nya dari store To (gudang tujuan)
function cmdInsertStockOnHandStoreTo(txtStoreCodeTo,txtItemCode,txtQty,txtKolomTransaksi)

	DIM txtnamakolomqty, txtnamakolomamount, txtModuleName
	IF txtKolomTransaksi ="RN" THEN
			txtModuleName 		= "ReceiveNote.asp"
	End If
	IF txtKolomTransaksi ="SIV" THEN
			txtModuleName 		= "IssueDetails.asp"
	End If
	IF txtKolomTransaksi ="TRT" THEN
			txtModuleName 		= "TransportTask.asp"
	End If
	IF txtKolomTransaksi ="RCV" THEN
			txtModuleName 		= "ReceivedTask.asp"
	End If

	IF txtKolomTransaksi ="TRTDIV" THEN
			txtModuleName 		= "TransportTaskDiv.asp"
	End If

	IF txtKolomTransaksi ="RT" THEN
			txtModuleName 		= "StockReturn.asp"
	End If
	IF txtKolomTransaksi ="WRT" THEN
			txtModuleName 		= "WriteOff.asp"
	End If

	qryQtyOnHand1 = "Insert into stockonhand ( " & _
		" STORECODE, STOCKCODE, FCBA, QUANTITY, LASTUPDATE, LASTTIME ) values ( '" & _
		txtStoreCodeTo & "','" & _
		txtItemCode & "','" & _
		Session("fcBA") & "', '" & _ 
		txtQty & "', sysdate, To_Char(sysdate,'HH24:MI') ) "
	connStr.execute(qryQtyOnHand1) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

End Function 

' function ini digunakan untuk mengurangi qty di tabel stok on hand
Function cmdUpdateStockOnHand(txtStoreCode, txtItemCode, originalQuantity, txtKolomTransaksi,txtAction)
	DIM txtnamakolomqty, txtnamakolomamount, txtModuleName
	IF txtKolomTransaksi ="RN" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "ReceiveNote.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "ReceiveNote.asp"
		End If
	End If
	IF txtKolomTransaksi ="SIV" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "IssueDetails.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "IssueDetails.asp"
		End If
		
		IF txtAction = "Update" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "IssueDetails.asp"
		End If
	End If
	IF txtKolomTransaksi ="TRT" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "TransportTask.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "TransportTask.asp"
		End If
		
		IF txtAction = "Update" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "TransportTask.asp"
		End If
	End If
	IF txtKolomTransaksi ="RCV" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "ReceivedTask.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "ReceivedTask.asp"
		End If
	End If

	IF txtKolomTransaksi ="TRTDIV" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "TransportTaskDiv.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "TransportTaskDiv.asp"
		End If
		IF txtAction = "Update" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "TransportTaskDiv.asp"
		End If
	End If

	IF txtKolomTransaksi ="RT" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "StockReturn.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "StockReturn.asp"
		End If
		IF txtAction = "Update" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "StockReturn.asp"
		End If
	End If
	IF txtKolomTransaksi ="WRT" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "WriteOff.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "WriteOff.asp"
		End If
		IF txtAction = "Update" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "WriteOff.asp"
		End If
	End If
	
	'Response.End
	qryQtyOnHand1 = " Update stockonhand set QUANTITY= " & txtnamakolom & " '" & cdbl(originalQuantity) & "' " & _
	" where STORECODE ='" & txtStoreCode & "' and STOCKCODE='" & txtItemCode & "' and " & _
	" fcBA = '" & Session("fcBA") & "'" 
	connStr.execute(qryQtyOnHand1) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

End Function


' function ini digunakan untuk mengurangi qty di tabel stok on hand untuk gudang yang tujuan
Function cmdUpdateStockOnHandStoreTo(txtStoreCodeTo, txtItemCode, originalQuantity, txtKolomTransaksi,txtAction)
	DIM txtnamakolomqty, txtnamakolomamount,txtModuleName
	
	IF txtKolomTransaksi ="TRTDIV" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QUANTITY +"
			txtModuleName 		= "TransportTaskDiv.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "TransportTaskDiv.asp"
		End If
		IF txtAction = "Update" then
			txtnamakolom 		= "QUANTITY -"
			txtModuleName 		= "TransportTaskDiv.asp"
		End If
	End If
	
	'Response.End
	qryQtyOnHand1 = " Update stockonhand set QUANTITY= " & txtnamakolom & " '" & cdbl(originalQuantity) & "' " & _
	" where STORECODE ='" & txtStoreCodeTo & "' and STOCKCODE='" & txtItemCode & "' and " & _
	" fcBA = '" & Session("fcBA") & "'" 
	connStr.execute(qryQtyOnHand1) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry
	

End Function

Function UpdateNewStockOnHand(txtStoreCode, txtItemCode, txtQty, txtKolomTransaksi,txtAction)
	DIM txtnamakolomqty, txtnamakolomamount, txtModuleName
	IF txtKolomTransaksi ="SIV" THEN
		IF txtAction = "Update" then
				txtnamakolom 		= "QUANTITY -"
				txtModuleName 		= "IssueDetails.asp"
		End If
	End If
	IF txtKolomTransaksi ="TRT" THEN
		IF txtAction = "Update" then
				txtnamakolom 		= "QUANTITY -"
				txtModuleName 		= "TransportTask.asp"
		End If
	End If
	IF txtKolomTransaksi ="TRTDIV" THEN
		IF txtAction = "Update" then
				txtnamakolom 		= "QUANTITY -"
				txtModuleName 		= "TransportTaskDiv.asp"
		End If
	End If
	IF txtKolomTransaksi ="RT" THEN
		IF txtAction = "Update" then
				txtnamakolom 		= "QUANTITY +"
				txtModuleName 		= "StockReturn.asp"
		End If
	End If
	IF txtKolomTransaksi ="WRT" THEN
		IF txtAction = "Update" then
				txtnamakolom 		= "QUANTITY +"
				txtModuleName 		= "WriteOff.asp"
		End If
	End If
	qryQtyOnHand1 = " Update stockonhand set QUANTITY= " & txtnamakolom & " '" & cdbl(txtQty) & "' " & _
	" where STORECODE ='" & txtStoreCode & "' and STOCKCODE='" & txtItemCode & "' and " & _
	" fcBA = '" & Session("fcBA") & "'" 
	connStr.execute(qryQtyOnHand1) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry


End Function

Function UpdateNewStockOnHandStoreTo(txtStoreCodeTo, txtItemCode, txtQty, txtKolomTransaksi,txtAction)
	DIM txtnamakolomqty, txtnamakolomamount, txtModuleName
	IF txtKolomTransaksi ="TRTDIV" THEN
		IF txtAction = "Update" then
				txtnamakolom 		= "QUANTITY +"
				txtModuleName 		= "TransportTaskDiv.asp"
		End If
	End If
	qryQtyOnHand1 = " Update stockonhand set QUANTITY= " & txtnamakolom & " '" & cdbl(txtQty) & "' " & _
	" where STORECODE ='" & txtStoreCodeTo & "' and STOCKCODE='" & txtItemCode & "' and " & _
	" fcBA = '" & Session("fcBA") & "'" 
	connStr.execute(qryQtyOnHand1) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

End Function


'function ini digunakan untuk insert ke tabel stockledger
Function cmdInsertStockLedger(txtStoreCode,txtItemCode,txtDateCreated,txtQty,txtKolomTransaksi)
	DIM txtnamakolomqty, txtnamakolomamount,txtModuleName
	IF txtKolomTransaksi ="RN" THEN
		txtnamakolom 		= "QTYRECEIVED"
		txtnamakolomamount 	= "AMTRECEIVED"
		txtModuleName 		= "ReceiveNote.asp"
	End If
	IF txtKolomTransaksi ="SIV" THEN
		txtnamakolom 		= "QTYISSUED"
		txtnamakolomamount 	= "AMTISSUED"
		txtModuleName 		= "IssueDetails.asp"
	End If
	IF txtKolomTransaksi ="TRT" THEN
		txtnamakolom 		= "QTYTRANSPORTTASK"
		txtnamakolomamount 	= "AMTTRANSPORTTASK"
		txtModuleName 		= "TransportTask.asp"
	End If
	IF txtKolomTransaksi ="RCV" THEN
		txtnamakolom 		= "QTYRECEIVEDTASK"
		txtnamakolomamount 	= "AMTRECEIVEDTASK"
		txtModuleName 		= "ReceivedTask.asp"
	End If
	IF txtKolomTransaksi ="TRTDIV" THEN
		txtnamakolom 		= "QTYTRANSPORTTASKDIV"
		txtnamakolomamount 	= "AMTTRANSPORTTASKDIV"
		txtModuleName 		= "TransportTaskDiv.asp"
	End If

	IF txtKolomTransaksi ="RT" THEN
		txtnamakolom 		= "QTYRETURNED"
		txtnamakolomamount 	= "AMTRETURNED"
		txtModuleName 		= "StockReturn.asp"
	End If
	IF txtKolomTransaksi ="WRT" THEN
		txtnamakolom 		= "QTYWRITEOFF"
		txtnamakolomamount 	= "AMTWRITEOFF"
		txtModuleName 		= "WriteOff.asp"
	End If
	

	q = "INSERT INTO STOCKLEDGER (STORECODE, ITEMCODE,YEAR,MONTH," & txtnamakolom & "," & txtnamakolomamount & ",FCBA,LASTUPDATE,LASTTIME) " & _
	"  VALUES('" & txtStoreCode & "','" & txtItemCode & "',to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy'), " & _ 
	"  to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm'),'" & txtQty & "', " & _ 
	" 0,'" & Session("fcBA") & "',SYSDATE,To_Char(sysdate,'HH24:MI')) "
	connStr.execute(q) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

	q = "SELECT (NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)- " & _ 
	"NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)) ENDQTY " & _ 
	"FROM STOCKLEDGER " & _ 
	"WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCode & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	"AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	set rsStockLedger = Server.CreateObject("ADODB.Recordset")
	rsStockLedger.Open q, connstr
	if not rsStockLedger.eof then
		ENDQTY = rsStockLedger("ENDQTY")
	else
		ENDQTY = "0"
	end if 

	q = "UPDATE STOCKLEDGER SET ENDINGBALANCEQTY='" & ENDQTY & "' " & _ 
	" WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCode & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	" AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	connStr.execute(q)
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry
		

End Function 

'function ini digunakan untuk insert ke tabel stockledger yang isi kolom gudang nya diambil dari store To
Function cmdInsertStockLedgerStoreTo(txtStoreCodeTo,txtItemCode,txtDateCreated,txtQty,txtKolomTransaksi)
	DIM txtnamakolomqty, txtnamakolomamount, txtModuleName
	
	IF txtKolomTransaksi ="TRTDIV" THEN
		txtnamakolom 		= "QTYRECEIVETASKDIV"
		txtnamakolomamount 	= "AMTRECEIVETASKDIV"
		txtModuleName		= "ReceivedTaskDivisi"
	End If
	

	q = "INSERT INTO STOCKLEDGER (STORECODE, ITEMCODE,YEAR,MONTH," & txtnamakolom & "," & txtnamakolomamount & ",FCBA,LASTUPDATE,LASTTIME) " & _
	"  VALUES('" & txtStoreCodeTo & "','" & txtItemCode & "',to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy'), " & _ 
	"  to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm'),'" & txtQty & "', " & _ 
	" 0,'" & Session("fcBA") & "',SYSDATE,To_Char(sysdate,'HH24:MI')) "
	connStr.execute(q) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

	'update qty end yg gudangnya dari store to (gudang tujuan)
	q = "SELECT (NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)- " & _ 
	"NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)) ENDQTY " & _ 
	"FROM STOCKLEDGER " & _ 
	"WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCodeTo & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	"AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	set rsStockLedger = Server.CreateObject("ADODB.Recordset")
	rsStockLedger.Open q, connstr
	if not rsStockLedger.eof then
		ENDQTY = rsStockLedger("ENDQTY")
	else
		ENDQTY = "0"
	end if 

	q = "UPDATE STOCKLEDGER SET ENDINGBALANCEQTY='" & ENDQTY & "' " & _ 
	" WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCodeTo & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	" AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	connStr.execute(q)
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

End Function 

'Function ini digunakan untuk UPDATE kolom qty di tabel stockledger sesuai dengan transaksi masing-masing modul
Function cmdUpdateStockLedger(txtStoreCode,txtItemCode,txtDateCreated,txtQty,txtKolomTransaksi,txtAction) 

	DIM txtnamakolom,txtnamakolom1, txtModuleName
	IF txtKolomTransaksi = "RN" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QTYRECEIVED"
			txtnamakolom1 		= "QTYRECEIVED +"
			txtModuleName 		= "ReceiveNote.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QTYRECEIVED"
			txtnamakolom1 		= "QTYRECEIVED -"
			txtModuleName 		= "ReceiveNote.asp"
		End If
	End If
	IF txtKolomTransaksi = "SIV" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QTYISSUED"
			txtnamakolom1 		= "QTYISSUED +"
			txtModuleName 		= "IssueDetails.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QTYISSUED"
			txtnamakolom1 		= "QTYISSUED -"
			txtModuleName 		= "IssueDetails.asp"
		End If

		'Jika action Update kembali ksn dulu qty pada tabel stockledger
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYISSUED"
			txtnamakolom1 		= "QTYISSUED -"
			txtModuleName 		= "IssueDetails.asp"
		End If
	End If
	IF txtKolomTransaksi = "TRT" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QTYTRANSPORTTASK"
			txtnamakolom1 		= "QTYTRANSPORTTASK +"
			txtModuleName 		= "TransportTask.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QTYTRANSPORTTASK"
			txtnamakolom1 		= "QTYTRANSPORTTASK -"
			txtModuleName 		= "TransportTask.asp"
		End If

		'Jika action Update kembali ksn dulu qty pada tabel stockledger
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYTRANSPORTTASK"
			txtnamakolom1 		= "QTYTRANSPORTTASK -"
			txtModuleName 		= "TransportTask.asp"
		End If
	End If

	IF txtKolomTransaksi = "RCV" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QTYRECEIVEDTASK"
			txtnamakolom1 		= "QTYRECEIVEDTASK +"
			txtModuleName 		= "ReceivedTask.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QTYRECEIVEDTASK"
			txtnamakolom1 		= "QTYRECEIVEDTASK -"
			txtModuleName 		= "ReceivedTask.asp"
		End If
	End If
	IF txtKolomTransaksi = "TRTDIV" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QTYTRANSPORTTASKDIV"
			txtnamakolom1 		= "QTYTRANSPORTTASKDIV +"
			txtModuleName 		= "TransportTaskDiv.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QTYTRANSPORTTASKDIV"
			txtnamakolom1 		= "QTYTRANSPORTTASKDIV -"
			txtModuleName 		= "TransportTaskDiv.asp"
		End If
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYTRANSPORTTASKDIV"
			txtnamakolom1 		= "QTYTRANSPORTTASKDIV -"
			txtModuleName 		= "TransportTaskDiv.asp"
		End If
	End If
	IF txtKolomTransaksi = "RT" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QTYRETURNED"
			txtnamakolom1 		= "QTYRETURNED +"
			txtModuleName 		= "StockReturn.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QTYRETURNED"
			txtnamakolom1 		= "QTYRETURNED -"
			txtModuleName 		= "StockReturn.asp"
		End If
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYRETURNED"
			txtnamakolom1 		= "QTYRETURNED -"
			txtModuleName 		= "StockReturn.asp"
		End If
	End If
	IF txtKolomTransaksi = "WRT" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QTYWRITEOFF"
			txtnamakolom1 		= "QTYWRITEOFF +"
			txtModuleName 		= "WriteOff.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QTYWRITEOFF"
			txtnamakolom1 		= "QTYWRITEOFF -"
			txtModuleName 		= "WriteOff.asp"
		End If
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYWRITEOFF"
			txtnamakolom1 		= "QTYWRITEOFF -"
			txtModuleName 		= "WriteOff.asp"
		End If
	End If

	q = "update stockledger " & _
	"  SET " & txtnamakolom & " = " & txtnamakolom1 & " " & cdbl(txtQty) & " " & _ 
	", AMTISSUED = 0 " & _ 
	", LASTUPDATE = sysdate " & _
	", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
	" WHERE " & _
	" STORECODE = '" & txtStoreCode & "' and " & _
	" ITEMCODE = '" & txtItemCode & "' and " & _ 
	" YEAR  	= to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') and " & _ 
	" MONTH 	= to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm') and " & _ 
	" FCBA 		= '" & Session("fcBA") & "' "
	connStr.execute(q) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

	'update qty end 
	q = "SELECT (NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)- " & _ 
	"NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)) ENDQTY " & _ 
	"FROM STOCKLEDGER " & _ 
	"WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCode & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	"AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	set rsStockLedger = Server.CreateObject("ADODB.Recordset")
	rsStockLedger.Open q, connstr
	if not rsStockLedger.eof then
		ENDQTY = rsStockLedger("ENDQTY")
	else
		ENDQTY = "0"
	end if 

	q = "UPDATE STOCKLEDGER SET ENDINGBALANCEQTY='" & ENDQTY & "' " & _ 
	" WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCode & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	" AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	connStr.execute(q)
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

	

End Function 


'Function ini digunakan untuk UPDATE kolom qty di tabel stockledger yang isi gudangnya diambil dari store to
Function cmdUpdateStockLedgerStoreTo(txtStoreCodeTo,txtItemCode,txtDateCreated,txtQty,txtKolomTransaksi,txtAction) 
	DIM txtnamakolom,txtnamakolom1, txtModuleName
	IF txtKolomTransaksi = "TRTDIV" THEN
		IF txtAction = "Insert" then
			txtnamakolom 		= "QTYRECEIVETASKDIV"
			txtnamakolom1 		= "QTYRECEIVETASKDIV +"
			txtModuleName 		= "TransportTaskDiv.asp"
		END IF
		IF txtAction = "Delete" then
			txtnamakolom 		= "QTYRECEIVETASKDIV"
			txtnamakolom1 		= "QTYRECEIVETASKDIV -"
			txtModuleName 		= "TransportTaskDiv.asp"
		End If
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYRECEIVETASKDIV"
			txtnamakolom1 		= "QTYRECEIVETASKDIV -"
			txtModuleName 		= "TransportTaskDiv.asp"
		End If
	End If

	q = "update stockledger " & _
	"  SET " & txtnamakolom & " = " & txtnamakolom1 & " " & txtQty & " " & _ 
	", AMTISSUED = 0 " & _ 
	", LASTUPDATE = sysdate " & _
	", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
	" WHERE " & _
	" STORECODE = '" & txtStoreCodeTo & "' and " & _
	" ITEMCODE = '" & txtItemCode & "' and " & _ 
	" YEAR  	= to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') and " & _ 
	" MONTH 	= to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm') and " & _ 
	" FCBA 		= '" & Session("fcBA") & "' "
	connStr.execute(q) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

		'update qty end 
	q = "SELECT (NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)- " & _ 
	"NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)) ENDQTY " & _ 
	"FROM STOCKLEDGER " & _ 
	"WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCodeTo & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	"AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	set rsStockLedger = Server.CreateObject("ADODB.Recordset")
	rsStockLedger.Open q, connstr
	if not rsStockLedger.eof then
		ENDQTY = rsStockLedger("ENDQTY")
	else
		ENDQTY = "0"
	end if 

	q = "UPDATE STOCKLEDGER SET ENDINGBALANCEQTY='" & ENDQTY & "' " & _ 
	" WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCodeTo & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	" AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	connStr.execute(q)
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

End Function 

Function UpdateNewStockLedger(txtStoreCode,txtItemCode,txtDateCreated,txtQty,txtKolomTransaksi,txtAction)
	DIM txtnamakolom,txtnamakolom1, txtModuleName
	'Update kembali qty pada tabel stockledger
	IF txtKolomTransaksi = "SIV" THEN
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYISSUED"
			txtnamakolom1 		= "QTYISSUED +"
			txtModuleName 		= "IssueDetails.asp"
		End If
	End If
	IF txtKolomTransaksi = "TRT" THEN
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYTRANSPORTTASK"
			txtnamakolom1 		= "QTYTRANSPORTTASK +"
			txtModuleName 		= "TransportTask.asp"
		End If
	End If
	IF txtKolomTransaksi = "TRTDIV" THEN
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYTRANSPORTTASKDIV"
			txtnamakolom1 		= "QTYTRANSPORTTASKDIV +"
			txtModuleName 		= "TransportTaskDiv.asp"
		End If
	End If
	IF txtKolomTransaksi = "RT" THEN
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYRETURNED"
			txtnamakolom1 		= "QTYRETURNED +"
			txtModuleName 		= "StockReturn.asp"
		End If
	End If
	IF txtKolomTransaksi = "WRT" THEN
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYWRITEOFF"
			txtnamakolom1 		= "QTYWRITEOFF +"
			txtModuleName 		= "WriteOff.asp"
		End If
	End If

	q = "update stockledger " & _
	"  SET " & txtnamakolom & " = " & txtnamakolom1 & " " & txtQty & " " & _ 
	", AMTISSUED = 0 " & _ 
	", LASTUPDATE = sysdate " & _
	", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
	" WHERE " & _
	" STORECODE = '" & txtStoreCode & "' and " & _
	" ITEMCODE = '" & txtItemCode & "' and " & _ 
	" YEAR  	= to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') and " & _ 
	" MONTH 	= to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm') and " & _ 
	" FCBA 		= '" & Session("fcBA") & "' "
	connStr.execute(q) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

		'update qty end 
	q = "SELECT (NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)- " & _ 
	"NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)) ENDQTY " & _ 
	"FROM STOCKLEDGER " & _ 
	"WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCode & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	"AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	set rsStockLedger = Server.CreateObject("ADODB.Recordset")
	rsStockLedger.Open q, connstr
	if not rsStockLedger.eof then
		ENDQTY = rsStockLedger("ENDQTY")
	else
		ENDQTY = "0"
	end if 

	q = "UPDATE STOCKLEDGER SET ENDINGBALANCEQTY='" & ENDQTY & "' " & _ 
	" WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCode & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	" AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	connStr.execute(q)
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

End Function

Function UpdateNewStockLedgerStoreTo(txtStoreCodeTo,txtItemCode,txtDateCreated,txtQty,txtKolomTransaksi,txtAction)
	DIM txtnamakolom,txtnamakolom1, txtModuleName
	'Update kembali qty pada tabel stockledger
	
	IF txtKolomTransaksi = "TRTDIV" THEN
		IF txtAction = "Update" then
			txtnamakolom 		= "QTYRECEIVETASKDIV"
			txtnamakolom1 		= "QTYRECEIVETASKDIV +"
			txtModuleName		= "ReceivedtaskDiv"
		End If
	End If

	q = "update stockledger " & _
	"  SET " & txtnamakolom & " = " & txtnamakolom1 & " " & txtQty & " " & _ 
	", AMTISSUED = 0 " & _ 
	", LASTUPDATE = sysdate " & _
	", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
	" WHERE " & _
	" STORECODE = '" & txtStoreCodeTo & "' and " & _
	" ITEMCODE = '" & txtItemCode & "' and " & _ 
	" YEAR  	= to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') and " & _ 
	" MONTH 	= to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm') and " & _ 
	" FCBA 		= '" & Session("fcBA") & "' "
	connStr.execute(q) 
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

		'update qty end 
	q = "SELECT (NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)- " & _ 
	"NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)) ENDQTY " & _ 
	"FROM STOCKLEDGER " & _ 
	"WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCodeTo & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	"AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	set rsStockLedger = Server.CreateObject("ADODB.Recordset")
	rsStockLedger.Open q, connstr
	if not rsStockLedger.eof then
		ENDQTY = rsStockLedger("ENDQTY")
	else
		ENDQTY = "0"
	end if 

	q = "UPDATE STOCKLEDGER SET ENDINGBALANCEQTY='" & ENDQTY & "' " & _ 
	" WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCodeTo & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	" AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	connStr.execute(q)
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry

End Function

Function getSisaStock(txtStoreCodeTo,txtItemCode,txtDateCreated) 
	dim sisastock
	q = "SELECT (NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)- " & _ 
	"NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)) ENDQTY " & _ 
	"FROM STOCKLEDGER " & _ 
	"WHERE ITEMCODE='" & txtItemCode & "' AND STORECODE='" & txtStoreCodeTo & "' AND FCBA='" & Session("fcBA") & "' " & _ 
	"AND YEAR=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'yyyy') AND MONTH=to_char(to_date('" & txtDateCreated & "','mm/dd/yyyy'),'mm')"
	
	set rsStockLedger = Server.CreateObject("ADODB.Recordset")
	rsStockLedger.Open q, connstr
	if not rsStockLedger.eof then
		ENDQTY = rsStockLedger("ENDQTY")
	else
		ENDQTY = "0"
	end if 
	getSisaStock = ENDQTY
	'Response.Write sisastock
End Function

%>