	<!--#include file="connDb1.Inc"-->
    <%
    

		prcode   = Request("prcode")
		keterangan   = Request("keterangan")
		itemcode   = Request("itemcode")
		lineid   = Request("lineid")
        fcba   = Request("fcba")
        pocode   = Request("pocode")
        linenokeypod   = Request("linenokeypod")

	IF pocode<>"" THEN

		sqlview =  "UPDATE PURCHASEORDERDETAILS SET KETERANGAN = '" & keterangan & "',keterangan1=keterangan1 || '/' || '" & keterangan & "' " & _
		" WHERE FCBA='" & fcba & "' AND PRCODE ='" & prcode  & "' AND ITEMCODE  ='" & itemcode  & "' AND LINENOKEY  ='" & linenokeypod  & "'"
		response.write sqlview
		connStr.Execute sqlview

		sqlview =  "UPDATE PURCHASEREQUESTDETAILS SET KETERANGAN = '" & keterangan & "',keterangan1=keterangan1 || '/' || '" & keterangan & "' " & _
    	" WHERE FCBA='" & fcba & "' AND PRCODE ='" & prcode  & "' AND ITEMCODE  ='" & itemcode  & "' AND LINEID  ='" & lineid  & "'"
		response.write sqlview
    	connStr.Execute sqlview
	
    ELSE 
		sqlview =  "UPDATE PURCHASEREQUESTDETAILS SET KETERANGAN = '" & keterangan & "',keterangan1=keterangan1 || '/' || '" & keterangan & "' " & _
    	" WHERE FCBA='" & fcba & "' AND PRCODE ='" & prcode  & "' AND ITEMCODE  ='" & itemcode  & "' AND LINEID  ='" & lineid  & "'"
		response.write sqlview
    	connStr.Execute sqlview
		
	END IF
    %>