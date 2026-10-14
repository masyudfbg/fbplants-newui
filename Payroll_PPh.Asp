
<%
	'Amandemen
	'1. Harga Beras Jamsostek Allowded A17 Rate 1
	'2. Code Allowance = D03
	'3. D03     Rate JKK  =Rate_1 
	'           Rate Jabatan= Rate_2 
	'           Rate Pensiun= Rate_3 
	'           Rate Pajak Terutang= Rate_4 
	'           Max Potongan Biaya Jabatan= MAX_AMT 
	'           Total Biaya (RP) ditanggung Pemerintah =  MIN_AMT 
	
    ' Cari Harga Beras Jamsostek
	qBA="select nvl(sum(RATE_1),0) as Rate_1 from ALLOWDED where FcCode='A17' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.CursorType = 3
	rsBA.CursorLocation = 3
	rsBA.Open qBA, connstr
	if not rsBA.eof then
	   nHargaBerasJamsostek=cdbl(rsBA("Rate_1"))
	else
	   nHargaBerasJamsostek=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY9002 Harga Beras AllowDed Code A17 is Empty"
	   err.number=9002
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
    ' Cari Rate Jamsostek
	qBA=" select nvl(sum(RATE_1),0) as Rate_1,nvl(sum(RATE_2),0) as Rate_2,nvl(sum(RATE_3),0) as Rate_3," & _
	    " nvl(sum(RATE_4),0) as Rate_4,nvl(sum(MAX_AMT),0) as MAX_AMT,nvl(sum(MIN_AMT),0) as MIN_AMT " & _
	    " from ALLOWDED where FcCode='D03' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.CursorType = 3
	rsBA.CursorLocation = 3
	rsBA.Open qBA, connstr
	if not rsBA.eof then
	   nRateJKK=cdbl(rsBA("Rate_1"))
	   nRateJabatan=cdbl(rsBA("Rate_2"))
	   nRatePensiun=cdbl(rsBA("Rate_3"))
	   nRatePjkTerutang=cdbl(rsBA("Rate_4"))
	   nMaxJabatan=cdbl(rsBA("MAX_AMT"))
	   nPPh4Pemerintah = cdbl(rsBA("MIN_AMT"))
	else
	   nRateJKK=0
	   nRateJabatan=0
	   nRatePensiun=0
	   nRatePjkTerutang=0
	   nMaxJabatan=0
	   nPPh4Pemerintah = 0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY1101 Rate PPh Pasal 21 D03 is Empty"
	   err.number=9002
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
	
	qReceive=" select FCCODE,MARITALSTATUS,PRESENTGRADE,nvl(sum(PRESENTRATE),0) as PresentRate From Employee " & _
	         " WHERE DATETERMINATE is null and MARITALSTATUS<>'NT' " & _
	         " and fcBA='" & Session("fcBA") & "' GROUP BY FCCODE,MARITALSTATUS,PRESENTGRADE" 
	  
    Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDed.CursorType = 3
	rsEmpAllDed.CursorLocation = 3
	response.Write "Calculate PPH..."& "<br>"
	response.Flush()
	rsEmpAllDed.Open qReceive, connstr
	if not rsEmpAllDed.eof then
	   Do While Not rsEmpAllDed.eof
		  nGajiBruto    = 0
		  nBiayaJabatan =0
		  nBiayaPensiun =0
		  nTotalPotongan= 0
		  nGajiNet=0
		  nPPhPs21=0
	      nPemerintah=0

		  Set rsParameter = Server.CreateObject("ADODB.Recordset")
			  qParameter =" Select nvl(sum(RATE_2),0) as Rate_2 From PARAMETER where PARCODE='E0006' and " & _
			              " FCCODE='" & rsEmpAllDed("MARITALSTATUS") & "'" & _
				          " and fcBA='" & Session("fcBA") & "'"
		  rsParameter.CursorType = 3
		  rsParameter.CursorLocation = 3
		  rsParameter.Open qParameter, connstr
		  if not rsParameter.eof then
			 nBerasPekerja   = cdbl(rsParameter("Rate_2")) * nHargaBerasJamsostek
		  else
			 nBerasPekerja   = 0 
		  End if
		  if rsEmpAllDed("PRESENTGRADE") ="KBT" Then
		     JamsostekBase=cdbl(rsEmpAllDed("PRESENTRATE")) + nBerasPekerja
		  else  
		     JamsostekBase=(cdbl(rsEmpAllDed("PRESENTRATE"))*25) + nBerasPekerja
		  end if
		  rsParameter.Close
		  set rsParameter = Nothing
          'response.Write "JamsostekBase : " & JamsostekBase
		  'Gaji Total
		  Set rsParameter = Server.CreateObject("ADODB.Recordset")
			  qParameter =" Select nvl(sum(decode(substr(allowdedcode,1,1),'A',amount,amount*-1)),0) as TotalGaji " & _
		             " from EmpAllded where empcode='" & rsEmpAllDed("FCCODE") & "'" & _
                     " and to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			         " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and fcBA='" & Session("fcBA") & "'"
		  rsParameter.CursorType = 3
		  rsParameter.CursorLocation = 3
		  rsParameter.Open qParameter, connstr
		  if not rsParameter.eof then
			 nGajiTotal   = cdbl(rsParameter("TotalGaji")) 
		  else
			 nGajiTotal   = 0 
		  End if
		  rsParameter.Close
		  set rsParameter = Nothing
          'response.Write "nGajiTotal : " & nGajiTotal
  
 		  'Cari Nilai PTKP Employee Berdasakan Status
		  Set rsParameter = Server.CreateObject("ADODB.Recordset")
			  qParameter =" Select nvl(sum(RATE_1),0) as Rate_1 From PARAMETER where PARCODE='E0006' and " & _
			              " FCCODE='" & rsEmpAllDed("MARITALSTATUS") & "'" & _
				          " and fcBA='" & Session("fcBA") & "'"
		  rsParameter.CursorType = 3
		  rsParameter.CursorLocation = 3
		  rsParameter.Open qParameter, connstr
		  if not rsParameter.eof then
			 nPTKP= cdbl(rsParameter("Rate_1")) 
		  else
			 nPTKP= 0 
		  End if
		  rsParameter.Close
		  set rsParameter = Nothing
          'response.Write "nPTKP : " & nPTKP
	      
		  nGajiBruto    = nGajiTotal+(nRateJKK*JamsostekBase/100)
		  nBiayaJabatan =(nRateJabatan/100)*nGajiBruto
		  'response.Write "nBiayaJabatan : " & nBiayaJabatan
		  if nBiayaJabatan > nMaxJabatan Then
		     nBiayaJabatan=nMaxJabatan
		  end if
		  nBiayaPensiun =(nRatePensiun/100)*JamsostekBase
		  nTotalPotongan= nBiayaJabatan+nBiayaPensiun
		  'response.Write "nTotalPotongan : " & nTotalPotongan
		  nGajiNet=nGajiBruto-nTotalPotongan
		  'response.Write "nGajiNet : " & nGajiNet
		  nPPhPs21=(nRatePjkTerutang/100)*(nGajiNet-nPTKP)
		  

		  if nGajiBruto>=1000000 then
		     nPemerintah=(nRatePjkTerutang/100)*(nPPh4Pemerintah-nPTKP)
			 nPPhPs21=nPPhPs21-nPemerintah
          elseif nGajiBruto>=2000000 then
		     nPPhPs21=nPPhPs21
		  else
		     nPPhPs21=0
		  end if
		 ' response.Write "nGajiBruto : " & nGajiBruto
		  
		  if nPPhPs21 > 0 then
			 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
					    "('" & rsEmpAllDed("FCCODE") & "','" & "D03" & "'," &  " to_date('" & Session("v_Employee") & "','mm/dd/yyyy') " & "," & nPPhPs21 & _
					    ",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					    Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			 connstr.execute(qAllowance)
		  End if		
		  rsEmpAllDed.MoveNext
	  loop
	End if
	rsEmpAllDed.Close
	set rsEmpAllDed = Nothing
%>