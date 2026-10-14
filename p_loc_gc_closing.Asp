<%

Rownumber = 0
amountalokasi = 0
totalamountalokasi = 0
maxLineNoForTable = 0

firstrow_alokasi_jobcode = ""
firstrow_alokasi_tt = ""
firstrow_alokasi_tc = ""

jobcodeNoBasis = "1150901"

Set objRs = Server.CreateObject("ADODB.Recordset")								
objSql = " select nvl(max(lineno), 0) maxLineNo from gccb " & _
         " where fcba = '" &  Session("fcBA") & "' " 
objRs.Open objSql, connstr

if not objRs.EOF then
	maxLineNoForTable = cdbl(objRs("maxLineNo"))
else
	maxLineNoForTable = 0
end if

objRs.close
set objRs = Nothing

Set objRs = Server.CreateObject("ADODB.Recordset")
objSql = "select ControlJob from Module where NAME = 'General_Charges'"
objRs.Open objSql, connstr

if not objRs.EOF then
	txtControlJob = objRs("ControlJob")
else
	txtControlJob = ""
end if

objRs.close
set objRs = Nothing

Set objRs = Server.CreateObject("ADODB.Recordset")								
objSql = " select add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL" & _
         " where SYSTEM = 'General_Charges' and fcba = '" & Session("fcBA") & "'"
objRs.Open objSql, connstr

if not objRs.EOF then
	temp_datePeriodControlCA = objRs("LastPeriodPosted")
else
	temp_datePeriodControlCA = ""
end if

objRs.close
set objRs = Nothing

Set objRs = Server.CreateObject("ADODB.Recordset")								
objSql = " select nvl(max(lineno), 0) maxLineNo from cost_activity" & _
         " where fcba = '" &  Session("fcBA") & "' " 
objRs.Open objSql, connstr

if not objRs.EOF then
	maxLineNoForCA = cdbl(objRs("maxLineNo"))
else
	maxLineNoForCA = 0
end if

objRs.close
set objRs = Nothing

connstr.BeginTrans
sqldeletedata = " delete from gccb" & _
                " where fcba = '" & Session("fcBA") & "'" & _
                "       and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                "       and to_number(to_char(transdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
				"       and instr(referenceno, 'v2') <> 0"

'response.Write(sqldeletedata)			  
connStr.execute(sqldeletedata)

If Err.Number = 0 Then
	connstr.CommitTrans
	Response.Flush
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&sqldeletedata&"] ("&Err.Description&")")
	Response.Flush
End If

'--------------------- Find Basis by Unit (INTI, PLASMA)
		Set rsBasis = Server.CreateObject("ADODB.Recordset")
		rsBasis.CursorType = 3
		rsBasis.CursorLocation = 3
		nBasis = 0
		qBasis = "select fcba, MAX(basisinti)basisinti, MAX(basisplasma) basisplasma" & _
				 " from" & _
				 " (" & _
				 " select fcba, decode(ownership,'INTI',nvl(sum(unit), 0),0) basisinti, decode(ownership,'PLASMA',nvl(sum(unit), 0),0) basisplasma" & _
				 " from ( " & _
				 " select fcba " & _
				 " , decode(status, 'TM', decode(ownership,'INTI','6510101','6510103'), 'TBM', decode(ownership,'INTI','1322109','1322309'), '1150901') jobcode" & _
				 " , 'FF' targettype " & _
				 " , targetcode " & _
				 " , unit " & _
				 " , ownership" & _
				 " from ( " & _
				 " select fcba, ownership " & _
				 " , fccode targetcode " & _
				 " , field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')) status " & _
				 " , nvl(field_api.getvalid_ha (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')), 0) unit " & _
				 " from field " & _
				 " where fcba = '" & Session("fcBA") & "'" & _
				 " and nvl(field_api.getvalid_activation (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')),'N') = 'Y' " & _
				 " and field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')) is not null " & _
				 " ) " & _
				 " where unit <> 0 " & _
				 " )" & _
				 " group by fcba,ownership" & _
				 " )group by fcba"
				 
		rsBasis.Open qBasis, connstr
		if Not rsBasis.eof then
			nBasis = cdbl(rsBasis("basisinti"))
			nBasisPlasma = cdbl(rsBasis("basisplasma"))
		Else
			nBasis = 0
			nBasisPlasma = 0
		End if
		rsBasis.Close
		set rsBasis = Nothing
'--------------------- End Find Basis by Unit (INTI, PLASMA)

'--------------------- Total Cost of GC
qVariance = " select " & _
            " sourcegc.fcba, " & _
            " sourcegc.lastdate, " & _
            " 'GC' TARGETTYPE, " & _
            " sourcegc.sourcecode TARGETCODE, " & _
            " targetgc.amounttargetgc COSTTOTAL, " & _
            " sourcegc.amountsourcegc ALOKASI, " & _
            " (targetgc.amounttargetgc - sourcegc.amountsourcegc) as VAR from ( " & _
            " select fcba,LAST_DAY (TRUNC (tdate, 'mm')) lastdate, sourcecode, sum(amount) as amountsourcegc from iplasprod.cost_activity " & _
            " where fcba = '" & Session("fcBA") & "' " & _
            " and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
            " and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
            " and sourcetype = 'GC'" & _
			" and targettype ='FF'" & _
            " group by fcba,tdate,sourcecode" & _
            " order by fcba,tdate,sourcecode" & _
            " ) sourcegc" & _
            " full join (" & _
            " select fcba,targetcode, sum(amount) as amounttargetgc from iplasprod.cost_activity" & _
            " where fcba = '" & Session("fcBA") & "'" & _
            " and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
            " and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
            " and targettype = 'GC'" & _
            " group by fcba,targetcode" & _
            " order by fcba,targetcode" & _
            " ) targetgc" & _
            " on (" & _
            " sourcegc.fcba = targetgc.fcba and" & _
            " sourcegc.sourcecode = targetgc.targetcode" & _
            " )where targetgc.amounttargetgc-sourcegc.amountsourcegc  <> 0"
arrRS = GetRSArray(qVariance)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		totalamountalokasi = 0
		totalamountalokasiplasma = 0
		totalbiayain_v2 = 0
		
		totalbiayain_v2 = Cdbl(arrRS(6, i))
		
		if nBasis <> 0 then
			'buat for untuk 3 kondisi (6419998)
			If arrRS(3, i) = "99" Then
				maxLineNoForTable = maxLineNoForTable + 1
				connstr.BeginTrans
				qInsertCR = " insert into gccb ( fieldcode" & _
							"                          , targettype" & _
							"                          , targetcode" & _
							"                          , jobcode" & _
							"                          , quantity" & _
							"                          , amount" & _
							"                          , transdate" & _
							"                          , referenceno" & _
							"                          , fcentry" & _
							"                          , fcedit" & _
							"                          , fcip" & _
							"                          , fcba" & _
							"                          , lastupdate" & _
							"                          , lasttime" & _
							"                          , lineno)" & _
							" values ( '" & arrRS(3, i) & "'" & _
							"          , 'CB'" & _
							"          , 'BS" & "'" & _
							"          , '6429998'" & _
							"          , 0" & _
							"          , " & totalbiayain_v2 * -1 & "" & _
							"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
							"          , '" & arrRS(3, i) & "-Cr-v2'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & Session("fcBA") & "'" & _
							"          , sysdate" & _
							"          , to_char(sysdate, 'HH24:MI')" & _
							"          , " & maxLineNoForTable & ")"

				'response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qInsertCR&"] ("&Err.Description&")")
					Response.Flush
				End If
			End IF
			'6419999		
			If arrRS(3, i) = "98" Then
				maxLineNoForTable = maxLineNoForTable + 1
				connstr.BeginTrans
				qInsertCR = " insert into gccb ( fieldcode" & _
							"                          , targettype" & _
							"                          , targetcode" & _
							"                          , jobcode" & _
							"                          , quantity" & _
							"                          , amount" & _
							"                          , transdate" & _
							"                          , referenceno" & _
							"                          , fcentry" & _
							"                          , fcedit" & _
							"                          , fcip" & _
							"                          , fcba" & _
							"                          , lastupdate" & _
							"                          , lasttime" & _
							"                          , lineno)" & _
							" values ( '" & arrRS(3, i) & "'" & _
							"          , 'CB'" & _
							"          , 'BS" & "'" & _
							"          , '6419999'" & _
							"          , 0" & _
							"          , " & totalbiayain_v2 * -1 & "" & _
							"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
							"          , '" & arrRS(3, i) & "-Cr-v2'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & Session("fcBA") & "'" & _
							"          , sysdate" & _
							"          , to_char(sysdate, 'HH24:MI')" & _
							"          , " & maxLineNoForTable & ")"

				'response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qInsertCR&"] ("&Err.Description&")")
					Response.Flush
				End If
			End IF
						
			'6429998
			If arrRS(3, i) <> "98" And arrRS(3, i) <> "99" Then
			
				If arrRS(3, i) = "01" Then
					Response.Write arrRS(3, i) & "===" & totalbiayain_v2 & "</br>"
				End If
				
				maxLineNoForTable = maxLineNoForTable + 1
				connstr.BeginTrans
				qInsertCR = " insert into gccb ( fieldcode" & _
							"                          , targettype" & _
							"                          , targetcode" & _
							"                          , jobcode" & _
							"                          , quantity" & _
							"                          , amount" & _
							"                          , transdate" & _
							"                          , referenceno" & _
							"                          , fcentry" & _
							"                          , fcedit" & _
							"                          , fcip" & _
							"                          , fcba" & _
							"                          , lastupdate" & _
							"                          , lasttime" & _
							"                          , lineno)" & _
							" values ( '" & arrRS(3, i) & "'" & _
							"          , 'CB'" & _
							"          , 'BS" & "'" & _
							"          , '6419998'" & _
							"          , 0" & _
							"          , " & totalbiayain_v2 * -1 & "" & _
							"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
							"          , '" & arrRS(3, i) & "-Cr-v2'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & Session("fcBA") & "'" & _
							"          , sysdate" & _
							"          , to_char(sysdate, 'HH24:MI')" & _
							"          , " & maxLineNoForTable & ")"

				'response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qInsertCR&"] ("&Err.Description&")")
					Response.Flush
				End If
			End If			
		else 'tidak ada basis/ no activity
			' kondisi untuk non basis
		end if
	Next
	Erase arrRS
End If

'Generate Debet Inti V2
qlistactivity = "SELECT aloc.fcba," & _
				"     aloc.transdate, aloc.jobcode, 'FF' targettype, targetcode, unit, cost_to.sourcecode, cost_to.VAR" & _
				"     FROM (SELECT sourcegc.fcba, sourcegc.lastdate, 'GC' TARGETTYPE, sourcegc.sourcecode sourcecode," & _
				"     targetgc.amounttargetgc COSTTOTAL,sourcegc.amountsourcegc ALOKASI," & _
				"     (targetgc.amounttargetgc - sourcegc.amountsourcegc) AS VAR" & _
				"     FROM (  SELECT fcba, LAST_DAY (TRUNC (tdate, 'mm')) lastdate, sourcecode, SUM (amount) AS amountsourcegc" & _
				"     FROM iplasprod.cost_activity" & _
				"     WHERE fcba = '" & Session("fcBA") & "'" & _
				"     AND TO_NUMBER (TO_CHAR (tdate, 'yyyy')) =  " & Year(Session("v_General_Charges")) & "" & _
				"     AND TO_NUMBER (TO_CHAR (tdate, 'mm')) =  " & Month(Session("v_General_Charges")) & "" & _
				"     AND sourcetype = 'GC'" & _
				"     AND targettype = 'FF'" & _
				"     and sourcecode != '98'" & _
				"     GROUP BY fcba, tdate, sourcecode" & _
				"     ORDER BY fcba, tdate, sourcecode) sourcegc" & _
				"     FULL JOIN" & _
				"     (  SELECT fcba, targetcode, SUM (amount) AS amounttargetgc" & _
				"     FROM iplasprod.cost_activity" & _
				"     WHERE     fcba = '" & Session("fcBA") & "'" & _
				"     AND TO_NUMBER (TO_CHAR (tdate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				"     AND TO_NUMBER (TO_CHAR (tdate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
				"     AND targettype = 'GC'" & _
				"     GROUP BY fcba, targetcode" & _
				"     ORDER BY fcba, targetcode) targetgc" & _
				"     ON (    sourcegc.fcba = targetgc.fcba" & _
				"     AND sourcegc.sourcecode = targetgc.targetcode)" & _
				"     WHERE targetgc.amounttargetgc - sourcegc.amountsourcegc <> 0)" & _
				"     cost_to," & _
				"     (SELECT fcba," & _
				"     to_date('" & Month(Session("v_General_Charges")) & "/01/" & Year(Session("v_General_Charges")) & "', 'mm/dd/yyyy') transdate," & _
				"     DECODE (status,'TM', DECODE (ownership, 'INTI', '6510101', '6510103'),'TBM', DECODE (ownership, 'INTI', '1322109', '1322309'),'1150901')jobcode," & _
				"     'FF' targettype,targetcode,unit" & _
				"     FROM (SELECT fcba,ownership,fccode targetcode" & _
				"     , field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')) status" & _
				"     , nvl(field_api.getvalid_ha (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')), 0) unit" & _
				"     FROM field" & _
				"     WHERE" & _
				"     fcba = '" & Session("fcBA") & "'" & _
				"     AND ownership = 'INTI'" & _
				"     and nvl(field_api.getvalid_activation (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')),'N') = 'Y'" & _
				"     and field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')) is not null)" & _
				"     WHERE unit <> 0) aloc where cost_to.fcba = aloc.fcba order by cost_to.sourcecode asc"

'Response.Write qlistactivity
arrRS_dtl = GetRSArray_dtl(qlistactivity)

			If nBasis <> 0 then
				If IsArray(arrRS_dtl) Then
					totalamountalokasi = 0
					totalbiayain_v2 = 0
					For ii = LBound(arrRS_dtl, 2) To UBound(arrRS_dtl, 2)
					
					Rownumber = Rownumber + 1
					if Rownumber = 1 then
							temp_sourcecode = arrRS_dtl(6, ii)
							firstrow_alokasi_jobcode = arrRS_dtl(2, ii)
							firstrow_alokasi_tt = arrRS_dtl(3, ii)
							firstrow_alokasi_tc = arrRS_dtl(4, ii)
					end if
					
					If temp_sourcecode <> arrRS_dtl(6, ii) THEN						
						'Rounding
						if totalamountalokasi <> totalbiayain_v2 then
							
							maxLineNoForTable = maxLineNoForTable + 1
							connstr.BeginTrans
							qInsertDB = " insert into gccb ( fieldcode" & _
										"                          , targettype" & _
										"                          , targetcode" & _
										"                          , jobcode" & _
										"                          , quantity" & _
										"                          , amount" & _
										"                          , transdate" & _
										"                          , referenceno" & _
										"                          , fcentry" & _
										"                          , fcedit" & _
										"                          , fcip" & _
										"                          , fcba" & _
										"                          , lastupdate" & _
										"                          , lasttime" & _
										"                          , lineno)" & _
										" values ( '" & temp_sourcecode & "'" & _
										"          , '" & firstrow_alokasi_tt & "'" & _
										"          , '" & firstrow_alokasi_tc & "'" & _
										"          , '" & firstrow_alokasi_jobcode & "'" & _
										"          , 0" & _
										"          , " & totalbiayain_v2 - totalamountalokasi & "" & _
										"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
										"          , '" & temp_sourcecode & "-Db-v2'" & _
										"          , '" & Session("CodeName") & "'" & _
										"          , '" & Session("CodeName") & "'" & _
										"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
										"          , '" & Session("fcBA") & "'" & _
										"          , sysdate" & _
										"          , to_char(sysdate, 'HH24:MI')" & _
										"          , " & maxLineNoForTable & ")"

							'response.Write(qInsertDB)			  
							connStr.execute(qInsertDB)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&qInsertDB&"] ("&Err.Description&")")
								Response.Flush
							End If							
						end if		
						temp_sourcecode = arrRS_dtl(6, ii)
						totalamountalokasi = 0
						totalbiayain_v2 = 0
					End If
					
						totalbiayain_v2 = Cdbl(arrRS_dtl(7, ii))									
						amountalokasi = Round(cdbl(arrRS_dtl(5, ii))/ nBasis * totalbiayain_v2, 0)	
						temp_jobcode = arrRS_dtl(2, ii)
						
						If  arrRS_dtl(6, ii) = "99" Then
								
								'ganti jobcode ke depresiasi
								If temp_jobcode = "6510101" THEN
									temp_jobcode = "6510102"
								Else
									temp_jobcode = "1322209"
								End If
								
								'Response.Write qlistactivity
								
								maxLineNoForTable = maxLineNoForTable + 1
								connstr.BeginTrans
								qInsertDB = " insert into gccb ( fieldcode" & _
											"                          , targettype" & _
											"                          , targetcode" & _
											"                          , jobcode" & _
											"                          , quantity" & _
											"                          , amount" & _
											"                          , transdate" & _
											"                          , referenceno" & _
											"                          , fcentry" & _
											"                          , fcedit" & _
											"                          , fcip" & _
											"                          , fcba" & _
											"                          , lastupdate" & _
											"                          , lasttime" & _
											"                          , lineno)" & _
											" values ( '" & arrRS_dtl(6, ii) & "'" & _
											"          , '" & arrRS_dtl(3, ii) & "'" & _
											"          , '" & arrRS_dtl(4, ii) & "'" & _
											"          , '" & temp_jobcode & "'" & _
											"          , " & cdbl(arrRS_dtl(5, ii)) & "" & _
											"          , " & amountalokasi & "" & _
											"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
											"          , '" & arrRS_dtl(6, ii) & "-Db-v2'" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
											"          , '" & Session("fcBA") & "'" & _
											"          , sysdate" & _
											"          , to_char(sysdate, 'HH24:MI')" & _
											"          , " & maxLineNoForTable & ")"

								'response.Write(qInsertDB)			  
								connStr.execute(qInsertDB)
								
								If Err.Number = 0 Then
									connstr.CommitTrans
									Response.Flush
								Else
									connstr.RollbackTrans
									Response.write("Error Save ["&qInsertDB&"] ("&Err.Description&")")
									Response.Flush
								End If					
						End If
						
						If arrRS_dtl(6, ii) <> "99" Then 
									
										maxLineNoForTable = maxLineNoForTable + 1
										connstr.BeginTrans
										qInsertDB = " insert into gccb ( fieldcode" & _
													"                          , targettype" & _
													"                          , targetcode" & _
													"                          , jobcode" & _
													"                          , quantity" & _
													"                          , amount" & _
													"                          , transdate" & _
													"                          , referenceno" & _
													"                          , fcentry" & _
													"                          , fcedit" & _
													"                          , fcip" & _
													"                          , fcba" & _
													"                          , lastupdate" & _
													"                          , lasttime" & _
													"                          , lineno)" & _
													" values ( '" & arrRS_dtl(6, ii) & "'" & _
													"          , '" & arrRS_dtl(3, ii) & "'" & _
													"          , '" & arrRS_dtl(4, ii) & "'" & _
													"          , '" & temp_jobcode & "'" & _
													"          , " & cdbl(arrRS_dtl(5, ii)) & "" & _
													"          , " & amountalokasi & "" & _
													"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
													"          , '" & arrRS_dtl(6, ii) & "-Db-v2'" & _
													"          , '" & Session("CodeName") & "'" & _
													"          , '" & Session("CodeName") & "'" & _
													"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
													"          , '" & Session("fcBA") & "'" & _
													"          , sysdate" & _
													"          , to_char(sysdate, 'HH24:MI')" & _
													"          , " & maxLineNoForTable & ")"

										'response.Write(qInsertDB)			  
										connStr.execute(qInsertDB)
										
										If Err.Number = 0 Then
											connstr.CommitTrans
											Response.Flush
										Else
											connstr.RollbackTrans
											Response.write("Error Save ["&qInsertDB&"] ("&Err.Description&")")
											Response.Flush
										End If
						End If
						totalamountalokasi = totalamountalokasi + amountalokasi
					Next						
					Erase arrRS_dtl
				End If		
			End If
			
' End Generate Debet Inti V2

'Generate Debet Plasma V2 			
qlistactivityplasma = "SELECT aloc.fcba," & _
				"     aloc.transdate, aloc.jobcode, 'FF' targettype, targetcode, unit, cost_to.sourcecode, cost_to.VAR" & _
				"     FROM (SELECT sourcegc.fcba, sourcegc.lastdate, 'GC' TARGETTYPE, sourcegc.sourcecode sourcecode," & _
				"     targetgc.amounttargetgc COSTTOTAL,sourcegc.amountsourcegc ALOKASI," & _
				"     (targetgc.amounttargetgc - sourcegc.amountsourcegc) AS VAR" & _
				"     FROM (  SELECT fcba, LAST_DAY (TRUNC (tdate, 'mm')) lastdate, sourcecode, SUM (amount) AS amountsourcegc" & _
				"     FROM iplasprod.cost_activity" & _
				"     WHERE fcba = '" & Session("fcBA") & "'" & _
				"     AND TO_NUMBER (TO_CHAR (tdate, 'yyyy')) =  " & Year(Session("v_General_Charges")) & "" & _
				"     AND TO_NUMBER (TO_CHAR (tdate, 'mm')) =  " & Month(Session("v_General_Charges")) & "" & _
				"     AND sourcetype = 'GC'" & _
				"     AND targettype = 'FF'" & _
				"     and sourcecode = '98'" & _
				"     GROUP BY fcba, tdate, sourcecode" & _
				"     ORDER BY fcba, tdate, sourcecode) sourcegc" & _
				"     FULL JOIN" & _
				"     (  SELECT fcba, targetcode, SUM (amount) AS amounttargetgc" & _
				"     FROM iplasprod.cost_activity" & _
				"     WHERE     fcba = '" & Session("fcBA") & "'" & _
				"     AND TO_NUMBER (TO_CHAR (tdate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				"     AND TO_NUMBER (TO_CHAR (tdate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
				"     AND targettype = 'GC'" & _
				"     GROUP BY fcba, targetcode" & _
				"     ORDER BY fcba, targetcode) targetgc" & _
				"     ON (    sourcegc.fcba = targetgc.fcba" & _
				"     AND sourcegc.sourcecode = targetgc.targetcode)" & _
				"     WHERE targetgc.amounttargetgc - sourcegc.amountsourcegc <> 0)" & _
				"     cost_to," & _
				"     (SELECT fcba," & _
				"     to_date('" & Month(Session("v_General_Charges")) & "/01/" & Year(Session("v_General_Charges")) & "', 'mm/dd/yyyy') transdate," & _
				"     DECODE (status,'TM', DECODE (ownership, 'INTI', '6510101', '6510103'),'TBM', DECODE (ownership, 'INTI', '1322109', '1322309'),'1150901')jobcode," & _
				"     'FF' targettype,targetcode,unit" & _
				"     FROM (SELECT fcba,ownership,fccode targetcode" & _
				"     , field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')) status" & _
				"     , nvl(field_api.getvalid_ha (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')), 0) unit" & _
				"     FROM field" & _
				"     WHERE" & _
				"     fcba = '" & Session("fcBA") & "'" & _
				"     AND ownership = 'PLASMA'" & _
				"     and nvl(field_api.getvalid_activation (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')),'N') = 'Y'" & _
				"     and field_api.getvalid_status (fccode, fcba, to_date(" & Year(Session("v_General_Charges")) & "||lpad(" & Month(Session("v_General_Charges")) & ",2,'0')||'01','yyyymmdd')) is not null)" & _
				"     WHERE unit <> 0) aloc where cost_to.fcba = aloc.fcba"
			

			arrRS_dtl_plasma = GetRSArray_dtl_plasma(qlistactivityplasma)
			
			If nBasisPlasma <> 0 Then
				If IsArray(arrRS_dtl_plasma) Then
					totalamountalokasiplasma = 0
					totalbiayain_v3 = 0
					For ii = LBound(arrRS_dtl_plasma, 2) To UBound(arrRS_dtl_plasma, 2)
										
					totalbiayain_v3 = Cdbl(arrRS_dtl_plasma(7, ii))	
						Rownumber = Rownumber + 1
						
						if Rownumber = 1 then
							temp_sourcecodeplasma = arrRS_dtl_plasma(6, ii)
							firstrow_alokasi_jobcode = arrRS_dtl_plasma(2, ii)
							firstrow_alokasi_tt = arrRS_dtl_plasma(3, ii)
							firstrow_alokasi_tc = arrRS_dtl_plasma(4, ii)
						end if
						
						If temp_sourcecodeplasma <> arrRS_dtl_plasma(6, ii) Then
							'Rounding
							if totalamountalokasiplasma <> totalbiayain_v3 then
								
								maxLineNoForTable = maxLineNoForTable + 1
								connstr.BeginTrans
								qInsertDB = " insert into gccb ( fieldcode" & _
											"                          , targettype" & _
											"                          , targetcode" & _
											"                          , jobcode" & _
											"                          , quantity" & _
											"                          , amount" & _
											"                          , transdate" & _
											"                          , referenceno" & _
											"                          , fcentry" & _
											"                          , fcedit" & _
											"                          , fcip" & _
											"                          , fcba" & _
											"                          , lastupdate" & _
											"                          , lasttime" & _
											"                          , lineno)" & _
											" values ( '" & arrRS_dtl_plasma(6, ii) & "'" & _
											"          , '" & firstrow_alokasi_tt & "'" & _
											"          , '" & firstrow_alokasi_tc & "'" & _
											"          , '" & firstrow_alokasi_jobcode & "'" & _
											"          , 0" & _
											"          , " & totalbiayain_v3 - totalamountalokasiplasma & "" & _
											"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
											"          , '" & arrRS_dtl_plasma(6, ii) & "-Db-v2'" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
											"          , '" & Session("fcBA") & "'" & _
											"          , sysdate" & _
											"          , to_char(sysdate, 'HH24:MI')" & _
											"          , " & maxLineNoForTable & ")"

								'response.Write(qInsertDB)			  
								connStr.execute(qInsertDB)
								
								If Err.Number = 0 Then
									connstr.CommitTrans
									Response.Flush
								Else
									connstr.RollbackTrans
									Response.write("Error Save ["&qInsertDB&"] ("&Err.Description&")")
									Response.Flush
								End If
							end if
							temp_sourcecodeplasma = arrRS_dtl_plasma(6, ii)
							totalamountalokasiplasma = 0
							totalbiayain_v3 = 0
						End If
						amountalokasiplasma = Round(cdbl(arrRS_dtl_plasma(5, ii))/ nBasisPlasma * totalbiayain_v3, 2)
						
						temp_jobcode_plasma = arrRS_dtl_plasma(2, ii)
						
						If  arrRS_dtl_plasma(6, ii) = "04" Then
								
								Response.Write arrRS_dtl_plasma(6, ii) & " Plasma </br>"
								maxLineNoForTable = maxLineNoForTable + 1
								connstr.BeginTrans
								qInsertDB = " insert into gccb ( fieldcode" & _
											"                          , targettype" & _
											"                          , targetcode" & _
											"                          , jobcode" & _
											"                          , quantity" & _
											"                          , amount" & _
											"                          , transdate" & _
											"                          , referenceno" & _
											"                          , fcentry" & _
											"                          , fcedit" & _
											"                          , fcip" & _
											"                          , fcba" & _
											"                          , lastupdate" & _
											"                          , lasttime" & _
											"                          , lineno)" & _
											" values ( '" & arrRS_dtl_plasma(6, ii) & "'" & _
											"          , '" & arrRS_dtl_plasma(3, ii) & "'" & _
											"          , '" & arrRS_dtl_plasma(4, ii) & "'" & _
											"          , '" & temp_jobcode_plasma & "'" & _
											"          , " & cdbl(arrRS_dtl_plasma(5, ii)) & "" & _
											"          , " & amountalokasiplasma & "" & _
											"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
											"          , '" & arrRS_dtl_plasma(6, ii) & "-Db-v2'" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
											"          , '" & Session("fcBA") & "'" & _
											"          , sysdate" & _
											"          , to_char(sysdate, 'HH24:MI')" & _
											"          , " & maxLineNoForTable & ")"

								'response.Write(qInsertDB)			  
								connStr.execute(qInsertDB)
								
								If Err.Number = 0 Then
									connstr.CommitTrans
									Response.Flush
								Else
									connstr.RollbackTrans
									Response.write("Error Save ["&qInsertDB&"] ("&Err.Description&")")
									Response.Flush
								End If
						End If
						totalamountalokasiplasma = amountalokasiplasma + amountalokasiplasma						
					Next
					Erase arrRS_dtl_plasma
				End If						
			End If
' End Generate Debet Plasma V2

		Response.Write "Normalisasi Rounding" & "<br/>"
		Response.Flush
		
		qSelisih = " select fcba" & _
				   "        , fieldcode" & _
				   "        , transdate" & _
				   "        , amount" & _
				   " from (" & _
				   "     select fcba" & _
				   "            , fieldcode" & _
				   "            , transdate" & _
				   "            , sum(amount) amount" & _
				   "     from gccb" & _
				   "     where fcba = '" & Session("fcBA") & "'" & _
				   "           and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				   "           and to_number(to_char(transdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
				   "     group by fcba" & _
				   "              , fieldcode" & _
				   "              , transdate" & _
				   " )" & _
				   " where amount <> 0"
				   
		arrRS = GetRSArray(qSelisih)
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				Set objectRS = Server.CreateObject("ADODB.Recordset")
				objectSQL = " select min (lineno) minlineno" & _
							" from gccb" & _
							" where fcba = '" & Session("fcBA") & "'" & _
							"       and transdate = to_date('" & arrRS(2, i) & "', 'mm/dd/yyyy')" & _
							"       and fieldcode = '" & arrRS(1, i) & "'" & _
							"		and instr(referenceno, 'v2') <> 0 " & _
							"       and jobcode <> '6419998' and jobcode <> '6419999' and jobcode <> '6429998'"
				objectRS.Open objectSQL, connstr
				If Not objectRS.EOF Then
					FirstRecord = Cdbl(objectRS("minlineno"))
				Else
					FirstRecord = 0.0000001
				End If
				objectRS.close
				set objectRS = Nothing
				Response.Write arrRS(1, i) & "-" & cdbl(arrRS(3, i)) & "-" & FirstRecord & "</br>"
				UpdateAmoutSelisih = " update gccb" & _
									 " set amount = amount - " & cdbl(arrRS(3, i)) & "" & _
									 " where lineno = '" & FirstRecord & "'" & _
									 "       and fcba = '" & Session("fcBA") & "'"
				connStr.execute(UpdateAmoutSelisih)
			Next
			Erase arrRS
		End If
		
connStr.BeginTrans
InsertIntoCA = " insert into cost_activity (tdate" & _
               "                            , jobcode" & _
               "                            , source" & _
               "                            , sourcetype" & _
               "                            , sourcecode" & _
               "                            , sourcecode_status" & _
               "                            , targettype" & _
               "                            , targetcode" & _
               "                            , targetcode_status" & _
               "                            , referenceno" & _
               "                            , quantity" & _
               "                            , output" & _
               "                            , amount" & _
               "                            , fcba" & _
               "                            , closing_date" & _
               "                            , process_date" & _
               "                            , title_reference_01" & _
               "                            , auto_reference_01" & _
               "                            , title_reference_02" & _
               "                            , auto_reference_02" & _
               "                            , title_reference_03" & _
               "                            , auto_reference_03" & _
               "                            , title_reference_04" & _
               "                            , auto_reference_04" & _
               "                            , title_reference_05" & _
               "                            , auto_reference_05" & _
               "                            , debetcredit" & _
               " 							, accountingperiod, accountingyear " & _
               "                            , lastupdate" & _
               "                            , lasttime" & _
               "                            , lineno)" & _
               " select tdate" & _
               "        , jobcode" & _
               "        , source" & _
               "        , sourcetype" & _
               "        , sourcecode" & _
               "        , sourcecode_status" & _
               "        , targettype" & _
               "        , targetcode" & _
               "        , targetcode_status" & _
               "        , referenceno" & _
               "        , quantity" & _
               "        , output" & _
               "        , amount" & _
               "        , fcba" & _
               "        , closing_date" & _
               "        , sysdate process_date" & _
               "        , cost_activity_info_api.get_title_reference_01 (sourcetype,targettype,'Dr')" & _
               "        , targetcode" & _
               "        , cost_activity_info_api.get_title_reference_02 (sourcetype,targettype,'Dr')" & _
               "        , referenceno" & _
               "        , ''" & _
               "        , ''" & _
               "        , ''" & _
               "        , ''" & _
               "        , ''" & _
               "        , ''" & _
               "        , decode ( greaterorequalthan ( amount, 0), 'TRUE','Dr','Cr')" & _
               "        , " & Month (Session("v_General_Charges")) & "" & _
               "        , " & Year (Session("v_General_Charges")) & "" & _
               "        , sysdate lastupdate" & _
               "        , to_char(sysdate,'hh24:mi') lasttime" & _
               "        , rownum + " & maxLineNoForCA & " lineno " & _
               " from ( select transdate tdate" & _
               "               , jobcode jobcode" & _
               "               , 'GC' source" & _
               "               , 'GC' sourcetype" & _
               "               , fieldcode as sourcecode" & _
               "               , '' sourcecode_status" & _
               "               , targettype targettype" & _
               "               , targetcode targetcode" & _
               "               , '' targetcode_status" & _
               "               , referenceno referenceno" & _
               "               , sum(quantity) quantity" & _
               "               , 0 output" & _
               "               , sum(amount) amount" & _
               "               , fcba" & _
               "               , to_date('" & temp_datePeriodControlCA & "','mm/dd/yyyy') closing_date  " & _
               "        from gccb" & _
               "        where fcba = '" & Session("fcBA") & "'" & _
               "              and to_number(to_char(transdate,'yyyy')) = " & Year (Session("v_General_Charges")) & "" & _
               "              and to_number(to_char(transdate,'mm')) = " & Month (Session("v_General_Charges")) & "" & _
               "              and instr(referenceno, 'v2') <> 0" & _
               "        group by transdate" & _
               "                 , jobcode" & _
               "                 , 'GC' " & _
               "                 , 'GC' " & _
               "                 , fieldcode" & _
               "                 , ''" & _
               "                 , targettype" & _
               "                 , targetcode" & _
               "                 , ''" & _
               "                 , referenceno" & _
               "                 , fcba" & _
               " )"
			   
connstr.execute(InsertIntoCA)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&InsertIntoCA&"] ("&Err.Description&")")
	Response.Flush
End If

%>