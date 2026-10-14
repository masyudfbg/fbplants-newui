<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%

dim txtcari,fcba,tahun,bulan,msg
Dim TimeAwal,TimeAkhir,Hitung,Lama
    fcba  = session("fcba")
	tahun = Request("period_year")
    bulan = Request("period_month")
	
	

	sqlexec = "BEGIN " & _ 
              "fac_hpp_bulanan_api.ProcessHPP_Pabrik(" & bulan & "," & tahun & ",'" & fcba & "');" & _
			  "END; " 
	'response.write	sqlexec	  
    connStr.Execute(sqlexec)
  

Response.Write ("<script>window.close();</script>")   
Response.Flush
Response.End
%>
