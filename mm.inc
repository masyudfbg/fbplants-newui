<%
Function cmdCrossCekPO(txtGRNCode,txtNoPO,txtItemCode)
    dim qCrossCekPO,rsCrossCekPO,qTblPO
	rsCrossCekPO=Server.CreateObject("ADODB.Recordset")
	if txtItemCode="" Then
	   qCrossCekPO = " Update PURCHASEORDERDETAILS set RECEIVED=0 " & _
                     " where POCODE ='" & txtNoPO & "' and " &  _
				     " fcBA='" & Session("fcBA") & "'" 
	else
	   qCrossCekPO = " Update PURCHASEORDERDETAILS set RECEIVED=0 " & _
                     " where POCODE ='" & txtNoPO & "' and ITEMCODE='" & txtItemCode & "' and " & _
				     " fcBA='" & Session("fcBA") & "'" 
	end if
	Set rsCrossCekPO = connStr.Execute(qCrossCekPO)
	cmdCrossCekPO="DataOK"
End Function

Function cmdCrossCekPOUpdate (txtNoPo, txtItemCodeDelete,  originalQuantity )
   
    Dim strUpdate, rsCrossCekPOUpdate

    strUpdate ="Update PurchaseOrderDetails  set RECEIVED=received - " & originalQuantity & _
							" where POCode='" & txtNoPo & "' and LineNoKey= '" & txtItemCodeDelete & "'" & _
							" and fcBA='" & session("fcBA") & "'"
							
	Set rsCrossCekPOUpdate = connStr.execute(strUpdate)
	
	cmdCrossCekPOUpdate="DataOK"
	
End Function


Function cmdCheckHist(numLineNoHist, txtStockCode)
    dim qCheckHist,rsCheckHist,qTotalData
	
	qCheckHist = 	" select count(1) totalData from stocktranshist " & _
					" where " & _
					" lineno > " & numLineNoHist & " and " & _
					" fcba = '" & Session("fcBA") & "' and " & _
					" stockcode = '" & txtStockCode & "' " 
	
	
	set rsCheckHist = Server.CreateObject("ADODB.Recordset")
	rsCheckHist.Open qCheckHist, connstr
	if not rsCheckHist.eof then
		qTotalData = cdbl(rsCheckHist("totalData"))
	else
		qTotalData = 0
	end if 	
	
	
	cmdCheckHist=qTotalData
End Function


Function cmdCrossCekReceiveNote(txtGRNCode,txtNoPO,txtItemCode)
    dim qCrossCekPO,rsCrossCekPO,qTblPO
	rsCrossCekPO=Server.CreateObject("ADODB.Recordset")
	if txtItemCode="" Then
	   qCrossCekPO = " Update RECEIVENOTEDETAIL set AMOUNTINVOICE=0 " & _
                     " where RECEIVENOTECODE in (Select RECEIVENOTECODE from RECEIVENOTE where POCODE ='" & txtNoPO & "')" & " and " &  _
				     " fcBA='" & Session("fcBA") & "'" 
	else
	   qCrossCekPO = " Update RECEIVENOTEDETAIL set AMOUNTINVOICE=0 " & _
                     " where RECEIVENOTECODE in (Select RECEIVENOTECODE from RECEIVENOTE where POCODE ='" & txtNoPO & "') and PURCHASEITEMCODE='" & txtItemCode & "' and " & _
				     " fcBA='" & Session("fcBA") & "'" 
	end if
	'response.write "ini : " & qCrossCekPO
	response.flush()
	Set rsCrossCekPO = connStr.Execute(qCrossCekPO)
	cmdCrossCekReceiveNote="DataOK"
End Function



%>