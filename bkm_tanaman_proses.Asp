<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqlhdr,tglinput
			dim fcba
			dim nQtdFields,nCount,linenokey
			
			tglinput 		= Request("tglinput")
			fcba 			= session("fcba")
		

			nQtdFields 	= request.form("jumldata")
			nCount 		= 1
			linenokey 	= 1	
	
			do while cint(nCount) <= cint(nQtdFields)  

			 
						SQL = "insert into TANAMANTRN( TANGGAL, KODE, JUMLAH,FCBA) " & _ 
						" values (TO_DATE('" & tglinput & "', 'mm/dd/yyyy'),'" & request.form("kodeitem" & nCount) & "','" & request.form("jumlah" & nCount) & "','" & fcba & "')"
					''	Response.Write SQL
						connStr.execute(SQL)
						
			 nCount 	= nCount + 1
			 linenokey  = linenokey + 1
			 loop

			response.redirect("tanaman_input.asp?notrn="&tglinput)
			
	
%>

</body>
</html>