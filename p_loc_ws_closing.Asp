<%

Dim qVariance, qlistactivity, Rownumber, amountalokasi, totalamountalokasi, sqldeletedata, totalbiayain_v2
Dim firstrow_alokasi_jobcode, firstrow_alokasi_tt, firstrow_alokasi_tc, maxLineNoForTable, objRs, objSql, ii, arrRS_dtl
Dim maxLineNoForCA, temp_datePeriodControlCA, InsertIntoCA

Rownumber = 0
amountalokasi = 0
totalamountalokasi = 0
maxLineNoForTable = 0

firstrow_alokasi_jobcode = ""
firstrow_alokasi_tt = ""
firstrow_alokasi_tc = ""

jobcodeNoBasis = "108019997"

Set objRs = Server.CreateObject("ADODB.Recordset")								
objSql = " select nvl(max(lineno), 0) maxLineNo from Workshopcb " & _
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
objSql = "select ControlJob from Module where NAME = 'Workshop'"
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
sqldeletedata = " delete from Workshopcb" & _
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


'--------------------- Total Cost of Workshop
qVariance = " select fcba" & _
            "        , lastdate" & _
            "        , targettype" & _
            "        , targetcode" & _
            "        , costtotal" & _
            "        , alokasi" & _
            "        , costtotal - alokasi var" & _
            " from (" & _
            "     select fcba" & _
            "            , lastdate" & _
            "            , targettype" & _
            "            , targetcode" & _
            "            , costtotal " & _
            "            , new_loc_closing_api.cek_alokasi_per_source(fcba, lastdate, targettype, targetcode) alokasi" & _
            "     from (" & _
            "         select fcba" & _
            "                , last_day(trunc(tdate, 'mm')) lastdate" & _
            "                , targettype" & _
            "                , targetcode" & _
            "                , nvl(sum(amount),0) as costtotal" & _
            "         from cost_activity" & _
            "         where fcba = '" & Session("fcBA") & "'" & _
            "               and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
            "               and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
            "               and targettype = 'WS'" & _
            "         group by fcba, last_day(trunc(tdate, 'mm')), targettype, targetcode" & _
            "     )" & _
            " )" & _
            " where costtotal - alokasi <> 0"
				
arrRS = GetRSArray(qVariance)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		totalamountalokasi = 0
		totalbiayain_v2 = 0
		
		totalbiayain_v2 = Cdbl(arrRS(6, i))
		
	
		'--------------------- Find Basis by Unit
		Set rsBasis = Server.CreateObject("ADODB.Recordset")
		rsBasis.CursorType = 3
		rsBasis.CursorLocation = 3
		nBasis = 0
		qBasis = " select nvl(sum(unit), 0) basis" & _
		         " from workshopactivity" & _
			     " where fcba = '" & Session("fcBA") & "'" & _
				 "       and to_number(to_char(workshopdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				 "       and to_number(to_char(workshopdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _				 
				 "       and workshopcode = '" & arrRS(3, i) & "'" & _
				 "       and locationtype not in ('SOC', 'LOC')" & _
				 "       and unit is not null" & _
				 "       and unit <> 0"
				 
		rsBasis.Open qBasis, connstr
		if Not rsBasis.eof then
			nBasis = cdbl(rsBasis("Basis"))
		Else
			nBasis = 0
		End if
		rsBasis.Close
		set rsBasis = Nothing
		
		if nBasis <> 0 then
			maxLineNoForTable = maxLineNoForTable + 1
			connstr.BeginTrans
			qInsertCR = " insert into Workshopcb ( workshopcode" & _
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
						"          , '" & txtControlJob & "'" & _
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
			
			qlistactivity = " select workshopcode, workshopdate, job jobcode, locationtype, locationcode, nvl(sum(unit),0) unit" & _
		                    " from workshopactivity" & _
			                " where fcba = '" & Session("fcBA") & "'" & _
			                "       and to_number(to_char(workshopdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
			                "       and to_number(to_char(workshopdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _				 
			                "       and workshopcode = '" & arrRS(3, i) & "'" & _
			                "       and locationtype not in ('SOC', 'LOC')" & _
			                "       and unit is not null" & _
			                "       and unit <> 0" & _
							" group by workshopcode, workshopdate, job, locationtype, locationcode"
		
			arrRS_dtl = GetRSArray_dtl(qlistactivity)

			If IsArray(arrRS_dtl) Then
				For ii = LBound(arrRS_dtl, 2) To UBound(arrRS_dtl, 2)
					Rownumber = Rownumber + 1

					if Rownumber = 1 then
						firstrow_alokasi_jobcode = arrRS_dtl(2, ii)
						firstrow_alokasi_tt = arrRS_dtl(3, ii)
						firstrow_alokasi_tc = arrRS_dtl(4, ii)
					end if
					
					amountalokasi = Round(cdbl(arrRS_dtl(5, ii))/ nBasis * totalbiayain_v2, 0)
					
					maxLineNoForTable = maxLineNoForTable + 1
					connstr.BeginTrans
					qInsertDB = " insert into Workshopcb ( workshopcode" & _
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
								"          , '" & arrRS_dtl(3, ii) & "'" & _
								"          , '" & arrRS_dtl(4, ii) & "'" & _
								"          , '" & arrRS_dtl(2, ii) & "'" & _
								"          , " & cdbl(arrRS_dtl(5, ii)) & "" & _
								"          , " & amountalokasi & "" & _
								"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"          , '" & arrRS(3, i) & "-Db-v2'" & _
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
					
					totalamountalokasi = totalamountalokasi + amountalokasi
					
					Next
				Erase arrRS_dtl
			End If
			
			if totalamountalokasi <> totalbiayain_v2 then
			
				maxLineNoForTable = maxLineNoForTable + 1
				connstr.BeginTrans
				qInsertDB = " insert into Workshopcb ( workshopcode" & _
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
							"          , '" & firstrow_alokasi_tt & "'" & _
							"          , '" & firstrow_alokasi_tc & "'" & _
							"          , '" & firstrow_alokasi_jobcode & "'" & _
							"          , 0" & _
							"          , " & totalbiayain_v2 - totalamountalokasi & "" & _
							"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
							"          , '" & arrRS(3, i) & "-Db-v2'" & _
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
			
		else 'tidak ada basis/ no activity
		
			maxLineNoForTable = maxLineNoForTable + 1
			connstr.BeginTrans
			qInsertCR = " insert into Workshopcb ( workshopcode" & _
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
						"          , '" & txtControlJob & "'" & _
						"          , 0" & _
						"          , " & totalbiayain_v2 * -1 & "" & _
						"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
						"          , substr('" & arrRS(3, i) & "-Cr-v2-NoAct',1,14)" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"          , '" & Session("fcBA") & "'" & _
						"          , sysdate" & _
						"          , to_char(sysdate, 'HH24:MI')" & _
						"          , " & maxLineNoForTable & ")"

			'response.Write(qInsertCR)			  
			connStr.execute(qInsertCR)
			
			maxLineNoForTable = maxLineNoForTable + 1
			qInsertDB = " insert into Workshopcb ( workshopcode" & _
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
						"          , '" & jobcodeNoBasis & "'" & _
						"          , 0" & _
						"          , " & totalbiayain_v2 & "" & _
						"          , to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
						"          , substr('" & arrRS(3, i) & "-Db-v2-NoAct',1,14)" & _
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
				Response.write("Error Save ["&qInsertCR&"] - ["&qInsertDB&"] ("&Err.Description&")")
				Response.Flush
			End If
		end if
	
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
               "               , 'WS' source" & _
               "               , 'WS' sourcetype" & _
               "               , workshopcode as sourcecode" & _
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
               "        from workshopcb" & _
               "        where fcba = '" & Session("fcBA") & "'" & _
               "              and to_number(to_char(transdate,'yyyy')) = " & Year (Session("v_General_Charges")) & "" & _
               "              and to_number(to_char(transdate,'mm')) = " & Month (Session("v_General_Charges")) & "" & _
               "              and instr(referenceno, 'v2') <> 0" & _
               "        group by transdate" & _
               "                 , jobcode" & _
               "                 , 'WS' " & _
               "                 , 'WS' " & _
               "                 , workshopcode" & _
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