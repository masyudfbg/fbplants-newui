<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqldtl,sqldeletehdr,sqldeletedtl,sqlhdr
			dim nojournal,tglbkm,remarks,POSTINGFLG,fcba,view

			nojournal  = Request("jurnalno")
			jurnaldate = Request("jurnaldate")
			remarks    = Request("remarks")
			POSTINGFLG = Request("POSTINGFLG")
			fcba       = Request("fcba")
			view       = Request("view")
			sourcecode = Request("sourcecode")
			
			if remarks = "" then 
				remarks = "-"
			end if
			
			
			if POSTINGFLG = "0" then

			if sourcecode = "COM-SLS" Then
			sqldeletedtl = "update BKMJOURNALHDR set postingflg='1' where fcba='" & fcba & "' AND docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			'Response.write  sqldeletedtl
			end if

			else 


			' update postingflg'
			if sourcecode = "COM-SLS" Then
			sqldeletedtl = "update BKMJOURNALHDR set postingflg=0 where fcba='" & fcba & "' AND docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)
			end if

			end if
			
			sqlhdr = "BEGIN fac_closing_accounting_api.ProsesClosing(TO_NUMBER(TO_CHAR(TO_DATE('" & jurnaldate & "','DD/MM/YYYY'),'YYYY')), " & _ 
			                                                               "TO_NUMBER(TO_CHAR(TO_DATE('" & jurnaldate & "','DD/MM/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
				     "END;"													 
			'response.write(	sqlhdr)																 
			connStr.execute(sqlhdr)

			response.redirect("salesinv_jurnal_view.asp?invoicecode="&nojournal&"&fcba=" & fcba & "&tglinvoice=" & jurnaldate )
			
%>

</body>
</html>