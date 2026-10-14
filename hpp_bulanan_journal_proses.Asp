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
			sqldeletedtl = "BEGIN fac_hpp_bulanan_api.ProsesJournal_Post(" & bulansplit & "," & thnsplit & ",'" & businessunit & "','system'); " & _ 
				           "END;"	
			
			connStr.execute(sqldeletedtl)

			else 
			sqldeletedtl = "BEGIN fac_hpp_bulanan_api.ProsesJournal_UPost(" & bulansplit & "," & thnsplit & ",'" & businessunit & "','system'); " & _ 
				           "END;"	
			connStr.execute(sqldeletedtl)

			end if
			response.write sqldeletedtl
			response.redirect("hpp_bulanan_jurnal.asp?fcba="&businessunit&"&period_month="&bulansplit&"&period_year="&thnsplit)
%>

</body>
</html>