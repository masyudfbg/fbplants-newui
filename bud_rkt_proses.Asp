<%validatesession%>
<!--#include file="connDb1.Inc"-->
<html>
<body>
<%

			dim sqlhdr,sqlupah,sqlnomor,rsnomor,bkmno,sqldelkosong
			dim fcba
			dim nQtdFields,nCount,linenokey,gangcode,qty,price
			
			periode_year 	= Request("tahun")
			company     	= Request("company")
			fcba     	    = Request("businessunitbsb")			
			proses 			= Request("SIMPAN")
			
			'Response.write periode_year &"|" & company & "|" & proses
			
			if proses = "SIMPAN" then
			
			
			nQtdFields 	= request.form("jumldata")
			nCount 		= 1
			linenokey 	= 1	

			' insert data dtl upah   ovtimehours typeovrtime
			
			do while cint(nCount) <= cint(nQtdFields)  
			qty   = request.form("qty" & nCount)
			price = request.form("price" & nCount)
			if  qty = "" then 
				qty = 0
			end if
			if  price = "" then 
				price = 0
			end if
			
			if  trim(request.form("fccode" & nCount)) <> "" then

						SQL = "MERGE INTO bud_rkt a USING (" & _
						" SELECT '" & company & "' fccompanycode, '" & fcba & "' fcbusinessunit," & periode_year & " period_year, "  & _ 
						"'" & request.form("fccode" & nCount) & "' jobgroupcode, " & _ 
						"'" & request.form("fcname" & nCount) & "' jobgroupname, " & _ 
						"'" & request.form("uom" & nCount) & "' uom, " & _ 
						"'" & request.form("job_forfieldstatus" & nCount) & "' job_forfieldstatus, " & _ 
						"'" & request.form("ownership" & nCount) & "' ownership, " & _ 
						"'" & request.form("qtyunit" & nCount) & "' qtyunit, " & _ 
						qty & " qty, " & _ 
						price & " price, " & _ 
						"'" & request.form("remark" & nCount) & "' remark FROM DUAL " & _ 
						") b ON (a.fccompanycode=b.fccompanycode AND a.fcbusinessunit=b.fcbusinessunit AND a.period_year=b.period_year AND a.jobgroupcode=b.jobgroupcode AND a.ownership=b.ownership) " & _
						" WHEN MATCHED THEN " & _ 
						" 	UPDATE SET a.qty=b.qty, a.price = b.price , a.amount = (b.qty * b.price ) , a.remarks=b.remark, a.lastupdate=SYSDATE, fcedit='" & Session("CodeName")  & "'" & _
						" WHEN NOT MATCHED THEN " & _
						" INSERT (fccompanycode,fcbusinessunit,period_year ,jobgroupcode,jobgroupname,uom,job_forfieldstatus,qty ,price,amount,remarks ,createdate ,lastupdate,fcentry ,fcedit,fcip,ownership,qtyunit ) " & _
						" VALUES (b.fccompanycode,b.fcbusinessunit,b.period_year ,b.jobgroupcode,b.jobgroupname,b.uom,b.job_forfieldstatus,b.qty ,b.price,(b.qty*b.price),b.remark,SYSDATE,SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "',b.ownership,b.qtyunit )"
						
						'Response.Write SQL & "<br>"
						
						connStr.execute(SQL)
						
			end if	

			 nCount 	= nCount + 1
			 linenokey  = linenokey + 1
			 loop

			response.redirect("bud_rkt_input.asp?company="&company&"&fcba="&fcba&"&year="&periode_year)
			

			

			'-------------------Batas proses upah-------------------------------------'

			end if
			
			
	
%>

</body>
</html>