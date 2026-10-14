<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

		Dim sqlhdr,rshdr,sqldtl,rsdtl,sqlnomor,rsnomor,currentnomor,currentnomorfix,currentnomorview
		Dim invoiceno,refno,tgl,company,fcba,agrementcode,invoiceby,contractorcode,contractorname
		Dim dp,qty,amount,rate,remain,fcbases
		Dim tglsplit,bulansplit,thnsplit
		Dim sqlupdate,sqlupdateqty
		Dim sqlcek,rscek

		'invoiceno			= Request("fcCode")
		refno				= Request("refinvoice")
		tgl					= Request("DATECREATED")
		company				= Request("fccompanycode")
		fcba				= Request("businessunit")
		agrementcode		= Request("agrementcode")
		invoiceby			= Request("invoiceby")
		contractorcode		= Request("contractorcode")
		contractorname		= Request("contractorname")
		dp					= Request("dp")
		qty					= Request("qty")
		amount				= Request("amountback")
		rate				= Request("rateback")
		remain				= Request("remainback")
		fcbases				= session("fcba")

		
		
			bulansplit  = left(tgl,2)
			tglsplit    = mid(tgl,4,2)
			thnsplit    = right(tgl,4)
		IF len(fcbases)<=3 then
			' buat nomor transaksi'
			sqlnomor = "SELECT SUBSTR(INVOICECODE,16,5) AS LastNo " & _
			" FROM CONTRACTORINVOICEDP " & _
			" WHERE SUBSTR(INVOICECODE,11,4)='" & thnsplit & "' " & _
			" AND FCBA='" & fcbases & "' ORDER BY SUBSTR(INVOICECODE,16,5) DESC FETCH FIRST 1 ROWS ONLY"
		else 
			' buat nomor transaksi'
			sqlnomor = "SELECT SUBSTR(INVOICECODE,17,5) AS LastNo " & _
			" FROM CONTRACTORINVOICEDP " & _
			" WHERE SUBSTR(INVOICECODE,12,4)='" & thnsplit & "' " & _
			" AND FCBA='" & fcbases & "' ORDER BY SUBSTR(INVOICECODE,16,5) DESC FETCH FIRST 1 ROWS ONLY"
		end if
		Set rsnomor = Server.CreateObject("ADODB.Recordset")
		rsnomor.CursorType = 3
		rsnomor.CursorLocation = 3
		rsnomor.Open sqlnomor, connstr
		if rsnomor.eof then
				tglprev = "0000"
				currentnomor = tglprev + 1
				currentnomor = "0000" & currentnomor
				currentnomorfix = right(currentnomor,5)
				invoiceno=fcbases&"/"&"IVC"&"/"&bulansplit&thnsplit&"/"&currentnomorfix
				'Response.Write invoiceno
		else 
				tglprev = rsnomor("LastNo")
				currentnomor = tglprev + 1
				currentnomor = "0000" & currentnomor
				currentnomorfix = right(currentnomor,5)
				invoiceno	= fcbases&"/"&"IVC"&"/"&bulansplit&thnsplit&"/"&currentnomorfix
			'	Response.Write invoiceno

		end if
		rsnomor.close 
		set rsnomor = nothing
		
		if amount = "" then 
			amount = 0 
		end if
		if dp = "" then 
			dp = 0 
		end if
		if rate = "" then 
			rate = 0 
		end if
		if remain = "" then 
			remain = 0 
		end if


		' insert data hdr'
		sqlhdr = "insert into CONTRACTORINVOICEDP(INVOICE_BY, INVOICECODE, TRANSDATE, AGREEMENTCODE, CONTRACTORCODE, " & _
       	" QTY, RATE,AMOUNT, DP, REMAIN,FCBA, FCENTRY, FCEDIT, FCIP, LASTUPDATE,LASTTIME, PAYMENT) " & _ 
		" values('" & invoiceby & "','" & invoiceno & "',TO_DATE('" & tgl & "', 'mm/dd/yyyy'),'" & agrementcode & "','" & contractorcode & "','" & qty & "', " & _ 
		" '" & (rate) & "','" & (amount) & "','" & (dp) & "','" & (remain) & "','" & fcbases & "','" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
		" '" & Request.ServerVariables("REMOTE_ADDR") & "',SYSDATE,TO_CHAR(SYSDATE, 'HH24:MI'),'0')"
		connStr.execute(sqlhdr)
		'Response.Write sqlhdr
		response.redirect("contractorinvoicedp.asp")

%>

</body>
</html>