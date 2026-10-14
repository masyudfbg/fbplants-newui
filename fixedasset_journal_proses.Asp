<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqldtl,sqldeletehdr,sqldeletedtl,sqlhdr
			dim nojournal,tglbkm,remarks,POSTINGFLG

			nojournal = Request("jurnalno")
			tglbkm = Request("jurnaldate")
			remarks = Request("remarks")
			POSTINGFLG = Request("POSTINGFLG")
			bulansplit = Request("month")
			thnsplit   = Request("year")
			businessunit = Request("fcba")
			
			if remarks = "" then 
				remarks = "-"
			end if
			
			if POSTINGFLG = "0" then
			sqldeletedtl = "BEGIN fac_journal_fac_fix_api.FIXASSET_JOURNAL_POST(" & thnsplit & "," & bulansplit & ",'" & businessunit & "','system'); " & _ 
				           "END;"	
			
			connStr.execute(sqldeletedtl)

			else 
			sqldeletedtl = "BEGIN fac_journal_fac_fix_api.FIXASSET_JOURNAL_UPOST(" & thnsplit & "," & bulansplit & ",'" & businessunit & "','system'); " & _ 
				           "END;"	
			connStr.execute(sqldeletedtl)

			end if
			'response.write sqldeletedtl
			response.redirect("fixedasset_journal.asp?fcba="&businessunit&"&month="&bulansplit&"&tahun="&thnsplit)
%>

</body>
</html>