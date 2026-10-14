<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqldtl,sqldeletehdr,sqldeletedtl,sqlhdr
			dim nojournal,jurnaldate,remarks,POSTINGFLG,fcba,view

			nojournal  = Request("jurnalno")
			jurnaldate = Request("jurnaldate")
			remarks    = Request("remarks")
			POSTINGFLG = Request("POSTINGFLG")
			fcba       = Request("fcba")
			view       = Request("view")
			
			if remarks = "" then 
				remarks = "-"
			end if
			
			
			if POSTINGFLG = "0" then

			' update postingflg pada tabel bkmjournalhdr untuk paymentvoucher'
			sqldeletedtl = "update BKMJOURNALHDR set postingflg=1 where fcba='" & fcba & "' AND docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "update PAYMENTVOUCHER set postingflg='1',ROWSTATEID='Approved' where fcba='" & fcba & "' AND vouchercode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			else 

			' update postingflg pada tabel bkmjournalhdr untuk paymentvoucher'
			sqldeletedtl = "DELETE BKMJOURNALDTL where fcba='" & fcba & "' AND docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "DELETE BKMJOURNALHDR where fcba='" & fcba & "' AND docno='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "DELETE PAYMENTVOUCHER_HIST where fcba='" & fcba & "' AND vouchercode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			sqldeletedtl = "update PAYMENTVOUCHER set postingflg='0',ROWSTATEID='Confirmed'  where fcba='" & fcba & "' AND vouchercode='" & nojournal & "'"
			connStr.execute(sqldeletedtl)

			

			end if
			
			sqlhdr = "BEGIN fac_closing_accounting_api.ProsesClosing(TO_NUMBER(TO_CHAR(TO_DATE('" & jurnaldate & "','DD/MM/YYYY'),'YYYY')), " & _ 
			                                                               "TO_NUMBER(TO_CHAR(TO_DATE('" & jurnaldate & "','DD/MM/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
				     "END;"													 
		'	response.write(	sqlhdr)																 
			connStr.execute(sqlhdr)
	
           ' if view="0" then
			response.redirect("paymentvoucher_jurnal_input.asp?fccode="&nojournal&"&fcba=" & fcba & "&datecreated=" & jurnaldate)
			'else
			'response.redirect("paymentvoucher_jurnal_view.asp?tglbkm="&jurnaldate&"&fcba=" & fcba & "&bkmno=" & nojournal )
			'end if
%>

</body>
</html>