<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

		Dim sqlhdr,rshdr,sqldtl,rsdtl,sqlnomor,rsnomor,currentnomor,currentnomorfix,currentnomorview
		Dim invoiceno,refno,tgl,company,fcba,pocode,invoiceby,suppliercode,suppliername
		Dim dp,otherchargestock,amount,vat,amountplusvat,otherchargenonstock,ttlamount,nofaktur,fakturdate,ttlppn
		Dim tglsplit,bulansplit,thnsplit
		Dim sqlupdate,sqlupdateqty,sqldelete
		Dim sqlcek,rscek,gandtotalPO,DPPO,prosestipe,diskon

		prosestipe			= Request("prosestipe")

		if  prosestipe="1" then


			refno				= Request("refno")
			tgl					= Request("DATECREATED")
			company				= Request("fccompanycode")
			fcba				= Request("businessunit")
			pocode				= Request("pocode")
			invoiceby			= Request("invoiceby")
			suppliercode		= Request("suppliercode")
			suppliername		= Request("suppliername")
			dp					= Request("downpayment")
			otherchargestock	= Request("otherchargestock")
			diskon				= Request("diskon")
			amount				= Request("amounthidden")
			ttlppn				= Request("totalppn")
			vat					= Request("vat")
			amountplusvat		= Request("amountvat")
			otherchargenonstock	= Request("othercharge")
			ttlamount			= Request("totalamounthidden")
			nofaktur			= Request("nofaktur")
			fakturdate			= Request("fakturdate")
			fcbases				= session("fcba")
			jumldata			= Request("jumldata")

			persentaseall		= Request("persentaseall")
			amountbypercentins		= Request("amountbypercentins")

			'DETAIL'
			itemcode			= Request("itemcode")
			qty					= Request("qty")
			price				= Request("price")
			percent				= Request("percent")
			priceqty			= Request("priceqty")

			nilai_toggle		= Request("nilai_toggle")

			'tglsplit = "list($thn, $bln, $hr) = explode("-",$tglinv);
			bulansplit  = left(tgl,2)
			tglsplit    = mid(tgl,4,2)
			thnsplit    = right(tgl,4)

			gandtotalPO = 0
			DPPO = 0
			
			'validasi cek total dp tdk boleh lebih dari tot pocode
			sqlcek = "SELECT a.pocode,a.fcba,a.grandtotal,sum(Nvl(b.amountvat,0)) totalamount " & _
					"FROM purchaseorder a " & _ 
					"left JOIN   invoicedphdr b ON a.pocode=b.pocode AND a.fcba=b.fcba " & _
					"where a.fcba='" & fcba & "' and a.pocode='" & pocode & "'" & _
					"GROUP BY a.pocode,a.fcba,a.grandtotal"
			Set rscek = Server.CreateObject("ADODB.Recordset")
			rscek.CursorType = 3
			rscek.CursorLocation = 3
			rscek.Open sqlcek, connstr
			if not rscek.eof then
				gandtotalPO = cdbl(rscek("grandtotal"))
				DPPO        = cdbl(rscek("totalamount"))
			end if
			rscek.close 
			set rscek = nothing
			
			err.number = 0		
			DPPO = cdbl(DPPO) 
			
		'	if gandtotalPO >= DPPO then
			
			' buat nomor transaksi'
			sqlnomor = "SELECT SUBSTR(invoicecode,-5,5) AS LastNo " & _
			" FROM invoicedphdr " & _
			" WHERE SUBSTR(invoicecode,-10,4)='" & thnsplit & "' " & _
			" AND FCBA='" & fcba & "' ORDER BY SUBSTR(invoicecode,-5,5) DESC FETCH FIRST 1 ROWS ONLY"
			Set rsnomor = Server.CreateObject("ADODB.Recordset")
			rsnomor.CursorType = 3
			rsnomor.CursorLocation = 3
			rsnomor.Open sqlnomor, connstr
			if rsnomor.eof then
					tglprev = "0000"
					currentnomor = tglprev + 1
					currentnomor = "0000" & currentnomor
					currentnomorfix = right(currentnomor,5)
					invoiceno=fcba&"/"&"PY"&"/"&bulansplit&thnsplit&"/"&currentnomorfix
					'Response.Write "ok"
			else 
					tglprev = rsnomor("LastNo")
					currentnomor = tglprev + 1
					currentnomor = "0000" & currentnomor
					currentnomorfix = right(currentnomor,5)
					invoiceno	= fcba&"/"&"PY"&"/"&bulansplit&thnsplit&"/"&currentnomorfix
			end if
			rsnomor.close 
			set rsnomor = nothing
			

			if otherchargestock = "" then 
				otherchargestock = 0 
			end if
			if otherchargenonstock = "" then 
				otherchargenonstock = 0 
			end if
			if diskon = "" then 
				diskon = 0 
			end if
			if ttlppn = "" then 
				ttlppn = 0 
			end if
			if amountplusvat = "" then 
				amountplusvat = 0 
			end if
			if amount = "" then 
				amount = 0 
			end if


			' insert data hdr'
			sqlhdr = "insert into invoicedphdr(invoicecode,invoicedate,suppliercode,company,fcentry,fcedit, " & _ 
			" fcip,fcba,lastupdate,lasttime,pocode,amount,vat,amountvat,totalamount,no_faktur,fakturdate, " & _ 
			" refinvoice,othercharge,dp,invoice_by,totalppn,otherchargestock,payment,diskon,PERCENT,AMTBYPERCENT,TOGGLE) " & _ 
			" values('" & invoiceno & "',TO_DATE('" & tgl & "', 'mm/dd/yyyy'),'" & suppliercode & "','" & company & "','" & Session("CodeName") & "', " & _ 
			" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "',SYSDATE, " & _ 
			" TO_CHAR(SYSDATE, 'HH24:MI'),'" & pocode & "','" & cdbl(amount) & "','" & Int(vat) & "','" & cdbl(amountplusvat) & "', " & _ 
			" '" & cdbl(ttlamount) & "','" & nofaktur & "',TO_DATE('" & fakturdate & "', 'mm/dd/yyyy'),'" & refno & "', " & _ 
			" '" & cdbl(otherchargenonstock) & "','" & cdbl(dp) & "','" & invoiceby & "','" & cdbl(ttlppn) & "','" & cdbl(otherchargestock) & "','0', " & _ 
			" '" & cdbl(diskon) & "', "& cdbl(persentaseall) &", '"& cdbl(amountbypercentins) &"', '"& cdbl(nilai_toggle) &"')"
			connStr.execute(sqlhdr)
			' Response.Write sqlhdr

			nQtdFields = request.form("jumldata")
			nCount = 0

			' insert data dtl'
			do while cint(nCount) < cint(nQtdFields)  
			if request.form("qty" & nCount) > 0 then
					SQL = "insert into invoicedpdetail(invoicecode,itemcode,quantity,price,amount,fcentry,fcedit,fcip,fcba,lastupdate,lasttime,percent) " & _ 
					" values ('" & invoiceno & "','" & request.form("itemcode" & nCount) & "','" & request.form("qty" & nCount) & "', " & _
					" '" & request.form("price" & nCount) & "','" & request.form("amountdtl" & nCount) & "', " & _ 
					" '" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "', " & _ 
					" SYSDATE,TO_CHAR(SYSDATE, 'HH24:MI'), '"& request.form("percent" & nCount) &"')"
					connStr.execute(SQL)

					'update ke tabel purchaseorderdetails di kolom qtypayreq'
					if downpayment <= 0 then

					sqlupdate = "update purchaseorderdetails set qtypayreq=qtypayreq+'" & request.form("qty" & nCount) & "' where pocode='" & pocode & "' AND itemcode='" & request.form("itemcode" & nCount) & "'"
					connStr.execute(sqlupdate)
					end if
			end if
			nCount = nCount + 1
			loop

			'update ke tabel purchaseorder di kolom potype'
			'sqlupdate = "update purchaseorder set potype='1' where pocode='" & pocode & "'"
			'connStr.execute(sqlupdate)
			err.number = 0		
		'	Else
		'	err.number = 1001		
			'End IF
			response.redirect("paymentrequest.asp?fccode="&invoiceno&"&errno="&err.number&"&gandtotalPO="&gandtotalPO&"&DPPO="&DPPO)
			'exit sub
		end if 
		if  prosestipe="2" then 
			'proses edit
			invoiceno			= Request("fcCode")
			refno				= Request("refno")
			tgl					= Request("DATECREATED")
			company				= Request("fccompanycode")
			fcba				= Request("businessunit")
			pocode				= Request("pocode")
			invoiceby			= Request("invoiceby")
			suppliercode		= Request("suppliercode")
			suppliername		= Request("suppliername")
			dp					= Request("dp")
			otherchargestock	= Request("otherchargestock")
			amount				= Request("amounthidden")
			ttlppn				= Request("totalppn")
			vat					= Request("vat")
			amountplusvat		= Request("amountvat")
			otherchargenonstock	= Request("othercharge")
			diskon				= Request("diskon")
			ttlamount			= Request("totalamounthidden")
			nofaktur			= Request("nofaktur")
			fakturdate			= Request("fakturdate")
			fcbases				= session("fcba")
			jumldata			= Request("jumldata")

			'DETAIL'
			itemcode			= Request("itemcode")
			qty					= Request("qty")
			price				= Request("price")
			priceqty			= Request("priceqty")


			if otherchargestock = "" then 
				otherchargestock = 0 
			end if
			if otherchargenonstock = "" then 
				otherchargenonstock = 0 
			end if
			if diskon = "" then 
				diskon = 0 
			end if
			if ttlppn = "" then 
				ttlppn = 0 
			end if
			if amountplusvat = "" then 
				amountplusvat = 0 
			end if
			if amount = "" then 
				amount = 0 
			end if


			' insert data hdr'
			sqlhdr = "UPDATE invoicedphdr SET invoicedate=TO_DATE('" & tgl & "', 'mm/dd/yyyy'), " & _ 
			" amount='" & cdbl(amount) & "',vat='" & Int(vat) & "',amountvat='" & cdbl(amountplusvat) & "',totalamount='" & cdbl(ttlamount) & "', " & _
			" no_faktur='" & nofaktur & "',fakturdate=TO_DATE('" & fakturdate & "', 'mm/dd/yyyy'), " & _ 
			" refinvoice='" & refno & "',othercharge='" & cdbl(otherchargenonstock) & "',dp='" & cdbl(dp) & "',invoice_by='" & invoiceby & "', " & _ 
			" totalppn='" & cdbl(ttlppn) & "',otherchargestock='" & cdbl(otherchargestock) & "',diskon='" & cdbl(diskon) & "',percent='" & cdbl(persentaseall) & "',amtbypercent='" & cdbl(amountbypercent) & "'" & _ 
			" where invoicecode='" & invoiceno & "'"
			connStr.execute(sqlhdr)
			'Response.Write sqlhdr

			'delete tabel detail
			sqldelete = "delete from invoicedpdetail where invoicecode='" & invoiceno & "'"
			connStr.execute(sqldelete)

			nQtdFields = request.form("jumldata")
			nCount = 0

			' insert data dtl'
			do while cint(nCount) < cint(nQtdFields)  
					SQL = "insert into invoicedpdetail(invoicecode,itemcode,quantity,price,amount,fcentry,fcedit,fcip,fcba,lastupdate,lasttime) " & _ 
					" values ('" & invoiceno & "','" & request.form("itemcode" & nCount) & "','" & request.form("qty" & nCount) & "', " & _
					" '" & request.form("price" & nCount) & "','" & request.form("amountdtl" & nCount) & "', " & _ 
					" '" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "', " & _ 
					" SYSDATE,TO_CHAR(SYSDATE, 'HH24:MI'))"
					connStr.execute(SQL)
				'	Response.Write SQL

					'update ke tabel purchaseorderdetails di kolom qtypayreq'
					sqlupdate = "update purchaseorderdetails set qtypayreq=qtypayreq+'" & request.form("qty" & nCount) & "' where pocode='" & pocode & "' AND itemcode='" & request.form("itemcode" & nCount) & "'"
					connStr.execute(sqlupdate)
			nCount = nCount + 1
			loop
		end if
	response.redirect("paymentrequest.asp")


%>

</body>
</html>