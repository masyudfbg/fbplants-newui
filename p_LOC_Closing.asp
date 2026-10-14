<%
	Response.write "********** LOC Closing ***********<br/><br/>"
	Response.Flush
%>
<%
Dim first_fieldcode, first_transdate, first_jobcode, first_targettype, first_targetcode
Dim temp_referenceno

'---- Control Job
qControlJob="select ControlJob from Module where NAME='LOC'"
Set rsControlJob = Server.CreateObject("ADODB.Recordset")
rsControlJob.CursorType = 3
rsControlJob.CursorLocation = 3
rsControlJob.Open qControlJob, connstr
if not rsControlJob.eof then
   txtControlJob=rsControlJob("ControlJob")
else
   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC200368 Control Job For LOC is Empty"
   txtControlJob=""
end if
rsControlJob.close
set rsControlJob = Nothing

response.Flush()

if txtControlJob<>"" then
	'---- Normalisasi
	Response.Write " Normalisasi LOC Base Costing" & "<br>"
	Response.Flush()

	connstr.BeginTrans
	qNormalisasi = " delete" & _
				   " from loccb" & _
				   " where fcba = '" & Session("fcBA") & "'" & _
				   "       and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				   "       and to_number(to_char(transdate,'mm')) = " & Month(Session("v_General_Charges")) & ""
	connstr.execute(qNormalisasi)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&qNormalisasi&"] ("&Err.Description&")")
		Response.Flush
	End If
	
		connstr.BeginTrans
		qNormalisasiDetail = " update cost_activity_hist_tab set deleteflag = 'REMOVE'" & _
						 " where " & _
						 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
								" where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_General_Charges")) &  _
								" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_General_Charges")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE IN ('LOC','SOC'))"
						 
		connstr.execute(qNormalisasiDetail)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		End If
		
		connstr.BeginTrans
		qNormalisasiDetail = " delete from cost_activity_hist_tab where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

		connstr.execute(qNormalisasiDetail)
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		End If

	connstr.BeginTrans
    qNormalisasi=" UPDATE COST_ACTIVITY set deleteflag = 'REMOVE' where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_General_Charges")) &  _
				 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_General_Charges")) & " and fcBA='" & Session("fcBA") & "'" & _
				 " and SOURCETYPE IN ('LOC','SOC')"
				 
	connstr.execute(qNormalisasi)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&qNormalisasi&"] ("&Err.Description&")")
		Response.Flush
	End If

	connstr.BeginTrans
	qNormalisasiDetail = " delete cost_activity where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

	connstr.execute(qNormalisasiDetail)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	Response.write("<font face=Verdana size=1 color=Green>Normalisasi LOC Base Costing, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Generating Data..." & "<br/>"
	Response.Flush
	
	connstr.BeginTrans
	ExecuteSQL = " delete from costtotal_pertarget" & _
				 " where fcba = '" & Session("fcBA") & "'" & _
				 "       and year = " & Year(Session("v_General_Charges")) & "" & _
				 "       and month = " & Month(Session("v_General_Charges")) & "" & _
				 "       and targettype in ('LOC', 'SOC')"
				 
	'Response.Write ExecuteSQL
	'Response.Flush
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	ExecuteSQL = " delete from loc_allocation" & _
				 " where fcba = '" & Session("fcBA") & "'"
				 
	'Response.Write ExecuteSQL
	'Response.Flush
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	ExecuteSQL = " delete from attendance_gad_closing" & _
				 " where fcba = '" & Session("fcBA") & "'"
				 
	'Response.Write ExecuteSQL
	'Response.Flush
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	ExecuteSQL = " delete from allowances_loc_tab" & _
				 " where fcba = '" & Session("fcBA") & "'"
				 
	'Response.Write ExecuteSQL
	'Response.Flush
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	If Session("Estate_Mill") = "E" then
		ExecuteSQL = " insert into allowances_loc_tab" & _
                     " select fcba" & _
                     "        , to_number(to_char(tdate, 'yyyy')) year" & _
                     "        , to_number(to_char(tdate, 'mm')) month" & _
                     "        , allowdedcode" & _
                     "        , targettype" & _
                     "        , targetcode" & _
                     "        , sum(amount) amount" & _
                     " from ( select fcba" & _
                     "               , trunc(tdate, 'mm') tdate" & _
                     "               , substr(referenceno, 1, 3) allowdedcode" & _
                     "               , targettype" & _
                     "               , targetcode                  " & _
                     "               , amount" & _
                     "        from cost_activity" & _
                     "        where fcba = '" & Session("fcBA") & "'" & _
                     "              and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                     "              and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
                     "              and sourcetype in ('ES', 'EM')" & _
                     "              and targettype in ('SOC', 'LOC')" & _
                     "        union all" & _
                     "        select fcba" & _
                     "               , trunc(tdate, 'mm') tdate" & _
                     "               , null allowdedcode" & _
                     "               , targettype" & _
                     "               , targetcode           " & _
                     "               , amount" & _
                     "        from cost_activity" & _
                     "        where fcba = '" & Session("fcBA") & "'" & _
                     "              and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                     "              and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
                     "              and sourcetype not in ('ES', 'EM')" & _
                     "              and targettype in ('SOC', 'LOC')" & _
                     " )" & _
                     " group by fcba" & _
                     "          , to_number(to_char(tdate, 'yyyy'))" & _
                     "          , to_number(to_char(tdate, 'mm'))" & _
                     "          , allowdedcode" & _
                     "          , targettype" & _
                     "          , targetcode"

	ElseIf (Session("Estate_Mill") = "M" Or Session("Estate_Mill") = "P") Then
		ExecuteSQL = 	" insert into attendance_gad_closing" & _
						" SELECT GANGCODE," & _
						"        FDDATE," & _
						"        SUPERVISION_1," & _
						"        SUPERVISION_2," & _
						"        SUPERVISION_3," & _
						"        SUPERVISION_4," & _
						"        EMPLOYEECODE," & _
						"        ATTENDANCE," & _
						"        JOBCODE," & _
						"        LOCATIONTYPE," & _
						"        LOCATIONCODE," & _
						"        MANDAYS," & _
						"        OTHRS," & _
						"        RATE," & _
						"        UNIT," & _
						"        OUTPUT," & _
						"        REFERENCE," & _
						"        REMARKS," & _
						"        FCENTRY," & _
						"        FCEDIT," & _
						"        FCIP," & _
						"        FCBA," & _
						"        LASTUPDATE," & _
						"        LASTTIME," & _
						"        LINENOKEY," & _
						"        OVERTIME_HOURS," & _
						"        TYPE_OVERTIME," & _
						"        CHARGEJOB," & _
						"        CHARGETYPE," & _
						"        CHARGECODE," & _
						"        BUCKET," & _
						"        SPBNO," & _
						"        KG_JANJANG," & _
						"        KG_BRONDOLAN" & _
						"   FROM attendance_gad" & _
						" where fcba = '" & Session("fcBA") & "'" & _
						"       and to_number(to_char(fddate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
						"       and to_number(to_char(fddate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
						"       and attendance not in ('UN', 'KB', 'KL')" & _
						"       and locationtype = 'MA'" & _
						"       and mandays <> 0"
	End If

	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	If Session("Estate_Mill")="E" then
		ExecuteSQL = " select * from dual"
	ElseIf (Session("Estate_Mill")= "M" Or Session("Estate_Mill")= "P") Then
		ExecuteSQL = " insert into loc_allocation" & _
                     " (fcba, fddate, jobcode, locationtype, locationcode, mandays, loc_group)" & _
                     " select fcba" & _
                     "        , fddate" & _
                     "        , jobcode" & _
                     "        , locationtype" & _
                     "        , locationcode" & _
                     "        , mandays" & _
                     "        , 'MILL' loc_group" & _
                     " from (" & _
                     "     select fcba" & _
                     "            , trunc(fddate, 'mm') fddate" & _
                     "            , jobcode" & _
                     "            , locationtype" & _
                     "            , locationcode" & _
                     "            , sum(mandays) mandays" & _
                     "     from attendance_gad_closing" & _
                     "     where fcba = '" & Session("fcBA") & "'" & _
                     "     group by fcba" & _
                     "              , trunc(fddate, 'mm')" & _
                     "              , jobcode" & _
                     "              , locationtype" & _
                     "              , locationcode" & _
                     " )"
	End If


	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	connstr.BeginTrans
	If Session("Estate_Mill")="E" then
		ExecuteSQL = " insert into costtotal_pertarget" & _
					 " (FCBA, YEAR, MONTH, TARGETTYPE, TARGETCODE, COSTTOTAL, BASIS)" & _
                     " select fcba" & _
                     "        , to_number(to_char(tdate, 'yyyy')) year" & _
                     "        , to_number(to_char(tdate, 'mm')) month" & _
                     "        , targettype" & _
                     "        , targetcode" & _
                     "        , sum(amount) costtotal" & _
                     "        , 0 basis" & _
                     " from (" & _
                     "     select ca.fcba" & _
                     "            , ca.tdate" & _
                     "            , ca.targettype" & _
                     "            , ca.targetcode" & _
                     "            , ca.amount" & _
                     "     from ( select fcba" & _
                     "                   , trunc(tdate, 'mm') tdate" & _
                     "                   , targettype" & _
                     "                   , targetcode" & _
                     "                   , substr(referenceno, 1, 3) allowdedcode" & _
                     "                   , amount" & _
                     "            from cost_activity" & _
                     "            where fcba = '" & Session("fcBA") & "'" & _
                     "                  and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                     "                  and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
                     "                  and sourcetype in ('ES', 'EM')" & _
                     "                  and targettype in ('SOC', 'LOC')" & _
                     "     ) ca inner join (select fcba, fccode from parameter where fcba = '" & Session("fcBA") & "' and parcode = 'ALLOCATION' and param_9 = '2') uct" & _
                     "     on (ca.allowdedcode = uct.fccode and ca.fcba = uct.fcba)" & _
                     "     union all" & _
                     "     select fcba" & _
                     "            , trunc(tdate, 'mm') tdate" & _
                     "            , targettype" & _
                     "            , targetcode" & _
                     "            , amount" & _
                     "     from cost_activity" & _
                     "     where fcba = '" & Session("fcBA") & "'" & _
                     "     and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
                     "     and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
                     "     and sourcetype not in ('ES', 'EM')" & _
                     "     and targettype in ('SOC', 'LOC') " & _
                     " )" & _
                     " group by fcba" & _
                     "          , to_number(to_char(tdate, 'yyyy'))" & _
                     "          , to_number(to_char(tdate, 'mm'))" & _
                     "          , targettype" & _
                     "          , targetcode"
	else	
		ExecuteSQL = " insert into costtotal_pertarget" & _
					 " (FCBA, YEAR, MONTH, TARGETTYPE, TARGETCODE, COSTTOTAL, BASIS)" & _
					 " select fcba" & _
					 "        , to_number(to_char(tdate,'yyyy')) year" & _
					 "        , to_number(to_char(tdate,'mm')) month" & _
					 "        , targettype" & _
					 "        , targetcode" & _
					 "        , costtotal" & _
					 "        , loc_closing_api.get_basis(fcba, tdate, gcmaster_api.get_gcgroup(fcba, 'LOC', targetcode)) basis" & _
					 " from (" & _
					 "      select fcba" & _
					 "             , trunc(tdate, 'mm') tdate" & _
					 "             , targettype" & _
					 "             , targetcode" & _
					 "             , nvl(sum(amount),0) as costtotal" & _
					 "      from cost_activity" & _
					 "      where fcba = '" & Session("fcBA") & "'" & _
					 "            and to_number(to_char(tdate,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					 "            and to_number(to_char(tdate,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					 "            and targettype = 'LOC'" & _
					 "      group by fcba" & _
					 "               , trunc(tdate, 'mm')" & _
					 "               , targettype" & _
					 "               , targetcode" & _
					 " )"
	end if
	
	connstr.execute(ExecuteSQL)
	
	If Err.Number = 0 Then
		connstr.CommitTrans
	Else
		connstr.RollbackTrans
		Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
		Response.Flush
	End If
	
	Response.write("<font face=Verdana size=1 color=Green>Generating Data, Success!.<br/><br/></font>")
	Response.Flush

	Response.Write "Generate Credit Allocation" & "<br/>"
	Response.Flush
	
	jobcodeNoBasis = "1150901"
	
	If Session("Estate_Mill")="E" then
		qCreditAlokasi = " select *" & _
                                 " from ( select fcba" & _
                                 "                    , year" & _
                                 "                    , month" & _
                                 "                    , role" & _
                                 "                    , sgroup" & _
                                 "                    , targettype" & _
                                 "                    , targetcode" & _
                                 "                    , costtotal" & _
                                 "                    , new_loc_closing_api.get_basisperallowded(fcba, year, month, role, sgroup, targettype, targetcode) basis" & _
                                 "          from ( select fcba" & _
                                 "                             , year" & _
                                 "                             , month" & _
                                 "                             , role" & _
                                 "                             , sgroup" & _
                                 "                             , targettype" & _
                                 "                             , targetcode" & _
                                 "                             , sum(costtotal) costtotal" & _
                                 "                    from ( select fcba" & _
                                 "                                       , year" & _
                                 "                                       , month" & _
                                 "                                       , allowdedcode" & _
                                 "                                       , targettype" & _
                                 "                                       , targetcode" & _
                                 "                                       , costtotal" & _
                                 "                                       , new_loc_closing_api.get_allowded_role(allowdedcode, fcba) role" & _
                                 "                                       , new_loc_closing_api.get_sgroup(allowdedcode, fcba) sgroup" & _
                                 "                              from allowances_loc_tab" & _
                                 "                              where fcba = '" & Session("fcBA") & "'" & _
                                 "                    )" & _
                                 "                    where role <> 'BORONGAN'" & _
                                 "                    group by fcba" & _
                                 "                                  , year" & _
                                 "                                  , month" & _
                                 "                                  , role" & _
                                 "                                  , sgroup" & _
                                 "                                  , targettype" & _
                                 "                                  , targetcode" & _
                                 "          )" & _
                                 " )" & _
                                 " where basis = 0"
						 
		arrRS = GetRSArray(qCreditAlokasi)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(7, i)) <> 0 Then
					temp_referenceno = arrRS(4, i)
					if IsNull(temp_referenceno) then
						temp_referenceno = "New-UPR"
					end if
					
					connstr.BeginTrans
					qInsertCR = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(6, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & Cdbl(arrRS(7, i)) * -1 & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(6, i) & "-Cr-NoAct-" & temp_referenceno & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
					
					qInsertDB = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(6, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & jobcodeNoBasis & "'" & _
								"         , '0'" & _
								"         , " & Cdbl(arrRS(7, i)) & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(6, i) & "-Db-NoAct-" & temp_referenceno & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
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
				End If
			Next
			Erase arrRS
		End If
	else
		qCreditAlokasi = " select fcba" & _
						 "        , targettype" & _
						 "        , targetcode" & _
						 "        , costtotal" & _
						 "        , basis" & _
						 " from costtotal_pertarget" & _
						 " where fcba = '" & Session("fcBA") & "'" & _
						 "       and year = " & Year(Session("v_General_Charges")) & "" & _
						 "       and month = " & Month(Session("v_General_Charges")) & "" & _
						 "       and targettype = 'LOC'"
						 
		arrRS = GetRSArray(qCreditAlokasi)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(4, i)) = 0 Then
					connstr.BeginTrans
					qInsertCR = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(2, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS(3, i)) * -1 & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-Cr-No Basis" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
					
					qInsertDB = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(2, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & jobcodeNoBasis & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS(3, i)) & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-Db-No Basis" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
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
				Else
					connstr.BeginTrans
					qInsertCR = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(2, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & cdbl(arrRS(3, i)) * -1 & "" & _
								"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-Cr" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
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
			Next
			Erase arrRS
		End If
						 
    end if
				  

	
	Response.write("<font face=Verdana size=1 color=Green> Generate Credit Allocation, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Generate Debet Allocation" & "<br/>"
	Response.Flush
	
	
	If Session("Estate_Mill")="E" then
		qDebetAlokasi = " select aa.fcba" & _
						"        , aa.transdate" & _
						"        , aa.targettype source_sourcetype" & _
						"        , aa.targetcode source_sourcecode" & _
						"        , aa.costtotal" & _
						"        , new_loc_closing_api.get_mandays_per_st(aa.fcba, aa.year, aa.month, aa.section, aa.grade, aa.sourcetype) basis" & _
						"        , bb.jobcode" & _
						"        , bb.targettype" & _
						"        , bb.targetcode       " & _
						"        , bb.mandays actual" & _
						"        , aa.section" & _
						" from ( select fcba" & _
						"               , year" & _
						"               , month" & _
						"               , last_day(to_date(month||'/1/'||year, 'mm/dd/yyyy')) transdate" & _
						"               , new_loc_closing_api.get_sourcetype(targettype) sourcetype" & _
						"               , new_loc_closing_api.get_section(targettype, targetcode) section" & _
						"               , new_loc_closing_api.get_grade(targettype, targetcode) grade" & _
						"               , targettype" & _
						"               , targetcode" & _
						"               , costtotal       " & _
						"        from costtotal_pertarget" & _
						"        where fcba = '" & Session("fcBA") & "'" & _
						" ) aa" & _
						" inner join ( select fcba" & _
						"                      , transdate" & _
						"                      , section" & _
						"                      , grade" & _
						"                      , jobcode" & _
						"                      , targettype" & _
						"                      , targetcode" & _
						"                      , sum(mandays) mandays" & _
						"              from employee_activities_es_tab" & _
						"              where fcba = '" & Session("fcBA") & "'" & _
						"              group by fcba" & _
						"                       , transdate" & _
						"                       , section" & _
						"                       , grade" & _
						"                       , jobcode" & _
						"                       , targettype" & _
						"                       , targetcode" & _
						" ) bb" & _
						" on (aa.fcba = bb.fcba and aa.transdate = bb.transdate and aa.section = bb.section and aa.grade = bb.grade)" & _
						" where aa.sourcetype = 'ES'" & _
						" union all" & _
						" select aa.fcba" & _
						"        , aa.transdate" & _
						"        , aa.targettype source_sourcetype" & _
						"        , aa.targetcode source_sourcecode" & _
						"        , aa.costtotal" & _
						"        , new_loc_closing_api.get_mandays_per_st(aa.fcba, aa.year, aa.month, aa.section, aa.grade, aa.sourcetype) basis" & _
						"        , bb.jobcode" & _
						"        , bb.targettype" & _
						"        , bb.targetcode       " & _
						"        , bb.mandays actual" & _
						"        , aa.section" & _
						" from ( select fcba" & _
						"               , year" & _
						"               , month" & _
						"               , last_day(to_date(month||'/1/'||year, 'mm/dd/yyyy')) transdate" & _
						"               , new_loc_closing_api.get_sourcetype(targettype) sourcetype" & _
						"               , new_loc_closing_api.get_section(targettype, targetcode) section" & _
						"               , new_loc_closing_api.get_grade(targettype, targetcode) grade" & _
						"               , targettype" & _
						"               , targetcode" & _
						"               , costtotal       " & _
						"        from costtotal_pertarget" & _
						"        where fcba = '" & Session("fcBA") & "'" & _
						" ) aa" & _
						" inner join ( select fcba" & _
						"                      , transdate" & _
						"                      , section" & _
						"                      , grade" & _
						"                      , jobcode" & _
						"                      , targettype" & _
						"                      , targetcode" & _
						"                      , sum(mandays) mandays" & _
						"              from employee_activities_em_tab" & _
						"              where fcba = '" & Session("fcBA") & "'" & _
						"              group by fcba" & _
						"                       , transdate" & _
						"                       , section" & _
						"                       , grade" & _
						"                       , jobcode" & _
						"                       , targettype" & _
						"                       , targetcode" & _
						" ) bb" & _
						" on (aa.fcba = bb.fcba and aa.transdate = bb.transdate and aa.section = bb.section and aa.grade = bb.grade)" & _
						" where aa.sourcetype = 'EM'"
		
		arrRS = GetRSArray(qDebetAlokasi)

		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(5, i)) <> 0 Then
					nFFAlokasi = 0
					nFFAlokasi = Round(Cdbl(arrRS(4, i))/Cdbl(arrRS(5, i))*Cdbl(arrRS(9, i)), 0)
					If Cdbl(nFFAlokasi) <> 0 Then
						connstr.BeginTrans
						
						qInsertCR = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(3, i) & "'" & _
									"         , 'CB'" & _
									"         , 'BS'" & _
									"         , '" & txtControlJob & "'" & _
									"         , '" & arrRS(9, i) & "'" & _
									"         , " & nFFAlokasi * -1 & "" & _
									"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(3, i) & "-Cr-New-UPR'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertCR)			  
						connStr.execute(qInsertCR)
						
						qInsertDB = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(3, i) & "'" & _
									"         , '" & arrRS(7, i) & "'" & _
									"         , '" & arrRS(8, i) & "'" & _
									"         , '" & arrRS(6, i) & "'" & _
									"         , " & cdbl(arrRS(9, i)) & "" & _
									"         , " & nFFAlokasi & "" & _
									"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(3, i) & "-Db-New-UPR'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
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
				End If
			Next
			Erase arrRS
		End If
	else
		qDebetAlokasi = " select ca.fcba" & _
						"        , ca.fddate" & _
						"        , targettype" & _
						"        , targetcode" & _
						"        , costtotal" & _
						"        , basis" & _
						"        , jobcode" & _
						"        , locationtype" & _
						"        , locationcode" & _
						"        , mandays" & _
						"        , loc_group" & _
						" from (" & _
						"     select fcba" & _
						"            , to_date(lpad(month, 2, 0)||'/01/'||year,'mm/dd/yyyy') fddate" & _
						"            , targettype" & _
						"            , targetcode" & _
						"            , costtotal" & _
						"            , basis" & _
						"     from costtotal_pertarget" & _
						"     where fcba = '" & Session("fcBA") & "'" & _
						"           and year = " & Year(Session("v_General_Charges")) & "" & _
						"           and month = " & Month(Session("v_General_Charges")) & "" & _
						"           and targettype = 'LOC'" & _
						" ) ca" & _
						" inner join loc_allocation lo" & _
						" on (ca.fcba = lo.fcba and ca.fddate = lo.fddate and gcmaster_api.get_gcgroup(ca.fcba, 'LOC', ca.targetcode) = lo.loc_group)"
		
		arrRS = GetRSArray(qDebetAlokasi)

		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(5, i)) <> 0 Then
					nFFAlokasi = Round(Cdbl(arrRS(9, i))/Cdbl(arrRS(5, i))*Cdbl(arrRS(4, i)), 0)
					If Cdbl(nFFAlokasi) <> 0 Then
						connstr.BeginTrans
						qInsertDB = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(3, i) & "'" & _
									"         , '" & arrRS(7, i) & "'" & _
									"         , '" & arrRS(8, i) & "'" & _
									"         , '" & arrRS(6, i) & "'" & _
									"         , " & cdbl(arrRS(9, i)) & "" & _
									"         , " & nFFAlokasi & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(3, i) & "-Db" & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
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
				End If
			Next
			Erase arrRS
		End If
	end if

	Response.write("<font face=Verdana size=1 color=Green> Generate Debet Allocation, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Generate Non UPR Allocation" & "<br/>"
	Response.Flush
	
	If Session("Estate_Mill")="E" then
		qnon_upr = " select aaa.fcba" & _
                         "           , aaa.year" & _
                         "           , aaa.month" & _
                         "           , aaa.sgroup" & _
                         "           , aaa.targettype" & _
                         "           , aaa.targetcode" & _
                         "           , aaa.sourcetype" & _
                         "           , aaa.section" & _
						 "           , aaa.grade" & _
                         "           , aaa.costtotal" & _
                         "           , aaa.basis" & _
                         "           , bbb.jobcode" & _
                         "           , bbb.tt" & _
                         "           , bbb.tc" & _
                         "           , new_loc_closing_api.get_premi_alokasi( aaa.fcba" & _
                         "                                                                        , aaa.year" & _
                         "                                                                        , aaa.month" & _
                         "                                                                        , aaa.sourcetype" & _
                         "                                                                        , aaa.section" & _
                         "                                                                        , aaa.grade" & _
                         "                                                                        , aaa.sgroup" & _
                         "                                                                        , bbb.jobcode" & _
                         "                                                                        , bbb.tt" & _
                         "                                                                        , bbb.tc) actual" & _
						 " from ( " & _
                         "         select fcba" & _
                         "                   , year" & _
                         "                   , month" & _
                         "                   , sgroup" & _
                         "                   , targettype" & _
                         "                   , targetcode" & _
                         "                   , sourcetype" & _
                         "                   , section" & _
                         "                   , grade" & _
                         "                   , costtotal" & _
                         "                   , new_loc_closing_api.get_premi_basis( fcba" & _
                         "                                                                             , year" & _
                         "                                                                             , month" & _
                         "                                                                             , sourcetype" & _
                         "                                                                             , section" & _
                         "                                                                             , grade" & _
                         "                                                                             , sgroup) basis" & _
                         "         from (" & _
                         "                 select fcba" & _
                         "                           , year" & _
                         "                           , month" & _
                         "                           , sgroup" & _
                         "                           , targettype" & _
                         "                           , targetcode" & _
                         "                           , sourcetype" & _
                         "                           , section" & _
                         "                           , grade" & _
                         "                           , sum(costtotal) costtotal" & _
                         "                 from (" & _
                         "                         select aa.fcba" & _
                         "                                   , aa.year" & _
                         "                                   , aa.month" & _
                         "                                   , aa.allowdedcode" & _
                         "                                   , aa.targettype" & _
                         "                                   , aa.targetcode" & _
                         "                                   , new_loc_closing_api.get_sourcetype(aa.targettype) sourcetype" & _
                         "                                   , new_loc_closing_api.get_section(aa.targettype, aa.targetcode) section" & _
                         "                                   , new_loc_closing_api.get_grade(aa.targettype, aa.targetcode) grade" & _
                         "                                   , bb.sgroup" & _
                         "                                   , aa.costtotal" & _
                         "                         from allowances_loc_tab aa" & _
                         "                         inner join ( select fcba, fccode, param_10 sgroup" & _
                         "                                         from parameter" & _
                         "                                         where fcba = '" & Session("fcBA") & "'" & _
                         "                                                  and parcode = 'ALLOCATION'" & _
                         "                                                  and param_9 = '3'" & _
                         "                         ) bb" & _
                         "                         on (aa.allowdedcode = bb.fccode and aa.fcba = bb.fcba)" & _
                         "                         where aa.fcba = '" & Session("fcBA") & "'" & _
                         "                 )" & _
                         "                 group by fcba" & _
                         "                           , year" & _
                         "                           , month" & _
                         "                           , sgroup" & _
                         "                           , targettype" & _
                         "                           , targetcode" & _
                         "                           , sourcetype" & _
                         "                           , section" & _
                         "                           , grade" & _
                         "         )" & _
                         " ) aaa" & _
                         " inner join (select distinct fcba" & _
                         "                          , sourcetype" & _
                         "                          , section" & _
						 "                          , grade" & _
                         "                          , jobcode" & _
                         "                          , targettype tt" & _
                         "                          , targetcode tc " & _
                         "                from employee_activities_em_tab" & _
                         "                where fcba = '" & Session("fcBA") & "') bbb" & _
                         " on (aaa.fcba = bbb.fcba and aaa.sourcetype = bbb.sourcetype and aaa.section = bbb.section and aaa.grade = bbb.grade)"
       
				  
		arrRS = GetRSArray(qnon_upr)
	
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(10, i)) <> 0 Then 'Jika Basis <> 0
					val_hasil = 0
					val_hasil = Round( (Cdbl(arrRS(9, i))/Cdbl(arrRS(10, i))) * Cdbl(arrRS(14, i)) , 0)
					
					if val_hasil <> 0 then 'Jika hasil proporsi <> 0
						connstr.BeginTrans
						qInsertCR = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(5, i) & "'" & _
									"         , 'CB'" & _
									"         , 'BS'" & _
									"         , '" & txtControlJob & "'" & _
									"         , '" & arrRS(14, i) & "'" & _
									"         , " & val_hasil * -1 & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(5, i) & "-Cr-" & arrRS(3, i) & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertCR)			  
						connStr.execute(qInsertCR)
						
						qInsertDB = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(5, i) & "'" & _
									"         , '" & arrRS(12, i) & "'" & _
									"         , '" & arrRS(13, i) & "'" & _
									"         , '" & arrRS(11, i) & "'" & _
									"         , '" & arrRS(14, i) & "'" & _
									"         , " & val_hasil & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(5, i) & "-Db-" & arrRS(3, i) & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
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
				End If
			Next
			Erase arrRS
		End If
		
		'ES
		qnon_upr = " select aaa.fcba" & _
                         "           , aaa.year" & _
                         "           , aaa.month" & _
                         "           , aaa.sgroup" & _
                         "           , aaa.targettype" & _
                         "           , aaa.targetcode" & _
                         "           , aaa.sourcetype" & _
                         "           , aaa.section" & _
						 "           , aaa.grade" & _
                         "           , aaa.costtotal" & _
                         "           , aaa.basis" & _
                         "           , bbb.jobcode" & _
                         "           , bbb.tt" & _
                         "           , bbb.tc" & _
                         "           , new_loc_closing_api.get_premi_alokasi( aaa.fcba" & _
                         "                                                                        , aaa.year" & _
                         "                                                                        , aaa.month" & _
                         "                                                                        , aaa.sourcetype" & _
                         "                                                                        , aaa.section" & _
                         "                                                                        , aaa.grade" & _
                         "                                                                        , aaa.sgroup" & _
                         "                                                                        , bbb.jobcode" & _
                         "                                                                        , bbb.tt" & _
                         "                                                                        , bbb.tc) actual" & _
						 " from ( " & _
                         "         select fcba" & _
                         "                   , year" & _
                         "                   , month" & _
                         "                   , sgroup" & _
                         "                   , targettype" & _
                         "                   , targetcode" & _
                         "                   , sourcetype" & _
                         "                   , section" & _
                         "                   , grade" & _
                         "                   , costtotal" & _
                         "                   , new_loc_closing_api.get_premi_basis( fcba" & _
                         "                                                                             , year" & _
                         "                                                                             , month" & _
                         "                                                                             , sourcetype" & _
                         "                                                                             , section" & _
                         "                                                                             , grade" & _
                         "                                                                             , sgroup) basis" & _
                         "         from (" & _
                         "                 select fcba" & _
                         "                           , year" & _
                         "                           , month" & _
                         "                           , sgroup" & _
                         "                           , targettype" & _
                         "                           , targetcode" & _
                         "                           , sourcetype" & _
                         "                           , section" & _
                         "                           , grade" & _
                         "                           , sum(costtotal) costtotal" & _
                         "                 from (" & _
                         "                         select aa.fcba" & _
                         "                                   , aa.year" & _
                         "                                   , aa.month" & _
                         "                                   , aa.allowdedcode" & _
                         "                                   , aa.targettype" & _
                         "                                   , aa.targetcode" & _
                         "                                   , new_loc_closing_api.get_sourcetype(aa.targettype) sourcetype" & _
                         "                                   , new_loc_closing_api.get_section(aa.targettype, aa.targetcode) section" & _
                         "                                   , new_loc_closing_api.get_grade(aa.targettype, aa.targetcode) grade" & _
                         "                                   , bb.sgroup" & _
                         "                                   , aa.costtotal" & _
                         "                         from allowances_loc_tab aa" & _
                         "                         inner join ( select fcba, fccode, param_10 sgroup" & _
                         "                                         from parameter" & _
                         "                                         where fcba = '" & Session("fcBA") & "'" & _
                         "                                                  and parcode = 'ALLOCATION'" & _
                         "                                                  and param_9 = '3'" & _
                         "                         ) bb" & _
                         "                         on (aa.allowdedcode = bb.fccode and aa.fcba = bb.fcba)" & _
                         "                         where aa.fcba = '" & Session("fcBA") & "'" & _
                         "                 )" & _
                         "                 group by fcba" & _
                         "                           , year" & _
                         "                           , month" & _
                         "                           , sgroup" & _
                         "                           , targettype" & _
                         "                           , targetcode" & _
                         "                           , sourcetype" & _
                         "                           , section" & _
                         "                           , grade" & _
                         "         )" & _
                         " ) aaa" & _
                         " inner join (select distinct fcba" & _
                         "                          , sourcetype" & _
                         "                          , section" & _
						 "                          , grade" & _
                         "                          , jobcode" & _
                         "                          , targettype tt" & _
                         "                          , targetcode tc " & _
                         "                from employee_activities_es_tab" & _
                         "                where fcba = '" & Session("fcBA") & "') bbb" & _
                         " on (aaa.fcba = bbb.fcba and aaa.sourcetype = bbb.sourcetype and aaa.section = bbb.section and aaa.grade = bbb.grade)"
				  
		arrRS = GetRSArray(qnon_upr)
	
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(10, i)) <> 0 Then 'Jika Basis <> 0
					val_hasil = 0
					val_hasil = Round( (Cdbl(arrRS(9, i))/Cdbl(arrRS(10, i))) * Cdbl(arrRS(14, i)) , 0)
					
					if val_hasil <> 0 then 'Jika hasil proporsi <> 0
						connstr.BeginTrans
						qInsertCR = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(5, i) & "'" & _
									"         , 'CB'" & _
									"         , 'BS'" & _
									"         , '" & txtControlJob & "'" & _
									"         , '" & arrRS(14, i) & "'" & _
									"         , " & val_hasil * -1 & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(5, i) & "-Cr-" & arrRS(3, i) & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
						 'response.Write(qInsertCR)			  
						connStr.execute(qInsertCR)
						
						qInsertDB = " insert into loccb (fieldcode" & _
									"                   , targettype" & _
									"                   , targetcode" & _
									"                   , jobcode" & _
									"                   , quantity" & _
									"                   , amount" & _
									"                   , transdate" & _
									"                   , referenceno" & _
									"                   , fcentry" & _
									"                   , fcedit" & _
									"                   , fcip" & _
									"                   , fcba" & _
									"                   , lastupdate" & _
									"                   , lasttime" & _
									"                   )" & _
									" values ('" & arrRS(5, i) & "'" & _
									"         , '" & arrRS(12, i) & "'" & _
									"         , '" & arrRS(13, i) & "'" & _
									"         , '" & arrRS(11, i) & "'" & _
									"         , '" & arrRS(14, i) & "'" & _
									"         , " & val_hasil & "" & _
									"         , " & " to_date('" & Session("v_General_Charges") & "','mm/dd/yyyy')" & _
									"         , '" & arrRS(5, i) & "-Db-" & arrRS(3, i) & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Session("CodeName") & "'" & _
									"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"         , '" & Session("fcBA") & "'" & _
									"         , sysdate" & _
									"         , to_char(sysdate, 'HH24:MI')" & _
									"         )"
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
				End If
			Next
			Erase arrRS
		End If
		
	end if
	
	Response.write("<font face=Verdana size=1 color=Green>Generate Non UPR Allocation, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Generate Borongan Allocation" & "<br/>"
	Response.Flush
	
	If Session("Estate_Mill")="E" then
		qborongan = " select fcba" & _
					"        , transdate" & _
					"        , allowdedcode" & _
					"        , alokasi_tt targettype" & _
					"        , alokasi_tc targetcode" & _
					"        , jobcode" & _
					"        , targettype tt" & _
					"        , targetcode tc" & _
					"        , sum(amount) amount" & _
					"        , new_loc_closing_api.get_sgroup(allowdedcode, fcba) sgroup" & _
					" from (" & _
					"     select fcba" & _
					"            , trunc(transdate, 'mm') transdate" & _
					"            , allowdedcode" & _
					"            , employeecode" & _
					"            , section" & _
					"            , grade" & _
					"            , amount" & _
					"            , sourcetype" & _
					"            , jobcode" & _
					"            , targettype" & _
					"            , targetcode" & _
					"            , alokasi_jobcode" & _
					"            , alokasi_tt" & _
					"            , alokasi_tc" & _
					"     from employee_borongan_tab" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and alokasi_tt = 'LOC'" & _
					" )" & _
					" group by fcba" & _
					"          , transdate" & _
					"          , allowdedcode" & _
					"          , alokasi_tt" & _
					"          , alokasi_tc" & _
					"          , jobcode" & _
					"          , targettype" & _
					"          , targetcode" & _
					"          , new_loc_closing_api.get_sgroup(allowdedcode, fcba)"
				  
		arrRS = GetRSArray(qborongan)
	
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				If Cdbl(arrRS(8, i)) <> 0 Then 'Jika Basis <> 0
					connstr.BeginTrans
					val_hasil = 0
					val_hasil = Round( Cdbl(arrRS(8, i)), 0)
					qInsertCR = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(4, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , '0'" & _
								"         , " & val_hasil * -1 & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(4, i) & "-Cr-" & arrRS(2, i) & "-" & arrRS(9, i) & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
					 'response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
					
					qInsertDB = " insert into loccb (fieldcode" & _
								"                   , targettype" & _
								"                   , targetcode" & _
								"                   , jobcode" & _
								"                   , quantity" & _
								"                   , amount" & _
								"                   , transdate" & _
								"                   , referenceno" & _
								"                   , fcentry" & _
								"                   , fcedit" & _
								"                   , fcip" & _
								"                   , fcba" & _
								"                   , lastupdate" & _
								"                   , lasttime" & _
								"                   )" & _
								" values ('" & arrRS(4, i) & "'" & _
								"         , '" & arrRS(6, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , '" & arrRS(5, i) & "'" & _
								"         , '0'" & _
								"         , " & val_hasil & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(4, i) & "-Db-" & arrRS(2, i) & "-" & arrRS(9, i) & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         )"
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
				End If
			Next
			Erase arrRS
		End If
	end if
	
	Response.write("<font face=Verdana size=1 color=Green>Generate Borongan Allocation, Success!.<br/><br/></font>")
	Response.Flush
	
	Response.Write "Normalisasi Rounding" & "<br/>"
	Response.Flush
	
	If Session("Estate_Mill")="E" then
		qSelisih = " select fcba" & _
                   "        , year" & _
                   "        , month" & _
                   "        , targettype" & _
                   "        , targetcode" & _
                   "        , referenceno" & _
                   "        , costtotal" & _
                   "        , alokasi" & _
                   "        , costtotal - alokasi var" & _
                   " from (" & _
                   "     select fcba" & _
                   "            , year" & _
                   "            , month" & _
                   "            , targettype" & _
                   "            , targetcode" & _
                   "            , referenceno" & _
                   "            , sum(costtotal) costtotal" & _
                   "            , NEW_LOC_CLOSING_API.get_totalalokasi(fcba, year, month, targetcode, referenceno) alokasi" & _
                   "     from (" & _
                   "         select fcba" & _
                   "                , year" & _
                   "                , month" & _
                   "                , allowdedcode" & _
                   "                , targettype" & _
                   "                , targetcode" & _
                   "                , new_loc_closing_api.get_sgroup(allowdedcode, fcba) referenceno" & _
                   "                , costtotal       " & _
                   "         from allowances_loc_tab " & _
                   "         where fcba = '" & Session("fcBA") & "'" & _
                   "     )" & _
                   "     group by fcba" & _
                   "              , year" & _
                   "              , month" & _
                   "              , targettype" & _
                   "              , targetcode" & _
                   "              , referenceno" & _
                   "              , NEW_LOC_CLOSING_API.get_totalalokasi(fcba, year, month, targetcode, referenceno)" & _
                   " )" & _
                   " where costtotal - alokasi <> 0"
				   
		arrRS = GetRSArray(qSelisih)
		
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
			
				Set objectRS = Server.CreateObject("ADODB.Recordset")
				objectSQL = " select fieldcode, transdate, jobcode, targettype, targetcode" & _
							" from (" & _
							"     select *" & _
							"     from loccb" & _
							"     where fcba = '" & Session("fcBA") & "'" & _
							"           and trunc(transdate, 'mm') = to_date('" & arrRS(2, i) & "/01/" & arrRS(1, i) & "', 'mm/dd/yyyy')" & _
							"           and jobcode <> '" & txtControlJob & "'" & _
							"           and fieldcode = '" & arrRS(4, i) & "'" & _
							"           and instr(referenceno, '" & arrRS(5, i) & "') <> 0" & _
							"           and abs(amount) > abs(" & Cdbl(arrRS(8, i)) & ")" & _
							"     order by lastupdate" & _
							" ) where rownum = 1"
				
				objectRS.Open objectSQL, connstr
				If Not objectRS.EOF Then
					first_fieldcode = objectRS("fieldcode")
					first_transdate = objectRS("transdate")
					first_jobcode = objectRS("jobcode")
					first_targettype = objectRS("targettype")
					first_targetcode = objectRS("targetcode")
				Else
					first_fieldcode = ""
					first_transdate = ""
					first_jobcode = ""
					first_targettype = ""
					first_targetcode = ""
				End If
				objectRS.close
				set objectRS = Nothing
				
				connstr.BeginTrans
				qInsertCR = " insert into loccb (fieldcode" & _
							"                   , targettype" & _
							"                   , targetcode" & _
							"                   , jobcode" & _
							"                   , quantity" & _
							"                   , amount" & _
							"                   , transdate" & _
							"                   , referenceno" & _
							"                   , fcentry" & _
							"                   , fcedit" & _
							"                   , fcip" & _
							"                   , fcba" & _
							"                   , lastupdate" & _
							"                   , lasttime" & _
							"                   )" & _
							" values ('" & first_fieldcode & "'" & _
							"         , 'CB'" & _
							"         , 'BS'" & _
							"         , '" & txtControlJob & "'" & _
							"         , '0'" & _
							"         , " & Cdbl(arrRS(8, i)) * -1 & "" & _
							"         , " & " to_date('" & first_transdate & "','mm/dd/yyyy')" & _
							"         , '" & first_fieldcode & "-Cr-" & arrRS(5, i) & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         )"
				 'response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)
				
				qInsertDB = " insert into loccb (fieldcode" & _
							"                   , targettype" & _
							"                   , targetcode" & _
							"                   , jobcode" & _
							"                   , quantity" & _
							"                   , amount" & _
							"                   , transdate" & _
							"                   , referenceno" & _
							"                   , fcentry" & _
							"                   , fcedit" & _
							"                   , fcip" & _
							"                   , fcba" & _
							"                   , lastupdate" & _
							"                   , lasttime" & _
							"                   )" & _
							" values ('" & first_fieldcode & "'" & _
							"         , '" & first_targettype & "'" & _
							"         , '" & first_targetcode & "'" & _
							"         , '" & first_jobcode & "'" & _
							"         , '0'" & _
							"         , " & Cdbl(arrRS(8, i)) & "" & _
							"         , " & " to_date('" & first_transdate & "','mm/dd/yyyy')" & _
							"         , '" & first_fieldcode & "-Db-" & arrRS(5, i) & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         )"
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
			
			Next
			Erase arrRS
		End If

	else
		qSelisih = " select fcba" & _
				   "        , fieldcode" & _
				   "        , transdate" & _
				   "        , amount" & _
				   " from (" & _
				   "     select fcba" & _
				   "            , fieldcode" & _
				   "            , transdate" & _
				   "            , sum(amount) amount" & _
				   "     from loccb" & _
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
				objectSQL = " select fieldcode, transdate, jobcode, targettype, targetcode" & _
							" from loccb" & _
							" where fcba = '" & Session("fcBA") & "'" & _
							"       and transdate = to_date('" & arrRS(2, i) & "', 'mm/dd/yyyy')" & _
							"       and fieldcode = '" & arrRS(1, i) & "'" & _
							"       and jobcode <> '" & txtControlJob & "'" & _
							"       and rownum = 1"
				objectRS.Open objectSQL, connstr
				If Not objectRS.EOF Then
					first_fieldcode = objectRS("fieldcode")
					first_transdate = objectRS("transdate")
					first_jobcode = objectRS("jobcode")
					first_targettype = objectRS("targettype")
					first_targetcode = objectRS("targetcode")
				Else
					first_fieldcode = ""
					first_transdate = ""
					first_jobcode = ""
					first_targettype = ""
					first_targetcode = ""
				End If
				objectRS.close
				set objectRS = Nothing
				
				UpdateAmoutSelisih = " update loccb" & _
									 " set amount = amount - " & cdbl(arrRS(3, i)) & "" & _
									 " where fcba = '" & Session("fcBA") & "'" & _
									 "       and transdate = to_date('" & first_transdate & "', 'mm/dd/yyyy')" & _
									 "       and fieldcode = '" & first_fieldcode & "'" & _
									 "       and jobcode = '" & first_jobcode & "'" & _
									 "       and targettype = '" & first_targettype & "'" & _
									 "       and targetcode = '" & first_targetcode & "'"
									 
				connStr.execute(UpdateAmoutSelisih)
			Next
			Erase arrRS
		End If
	end if
	
	Response.write("<font face=Verdana size=1 color=Green> Normalisasi Rounding, Success!.<br/><br/></font>")
	Response.Flush
	
	
	If Session("Estate_Mill")="E" then
		'----- Tranfer to Activity Costing
		response.Write "Transfer to Activity Based Costing..." & "<br>"
		Response.Flush

		' Get the Max LineNo of Cost Activity
		Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
		rsCostActivityMax.CursorType = 3
		rsCostActivityMax.CursorLocation = 3
										
		qCostActivityMax =  " select nvl(max(lineno), 0) maxLineNo from cost_activity " & _
							" where fcba = '" &  Session("fcBA") & "' " 
		rsCostActivityMax.Open qCostActivityMax, connstr
		
		If Not rsCostActivityMax.EOF Then
			maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))
		Else
			maxLineNo = 0
		End If

		rsCostActivityMax.close
		set rsCostActivityMax = Nothing
		' End of Get the Max LineNo of Cost Activity
		
		qABC = " INSERT INTO COST_ACTIVITY (" & _
			   " TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
			   " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,  " & _
										" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
										" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
										" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
										" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
										" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
										" LASTUPDATE, LASTTIME, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, LINENO ) " & _
			   " select TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
				  " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, sysdate PROCESS_DATE, " & _
						  " Cost_Activity_Info_Api.Get_Title_Reference_01 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),TARGETCODE, " & _
						  " Cost_Activity_Info_Api.Get_Title_Reference_02 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),referenceno, " & _
						  " '','', " & _
						  " '','', " & _
						  " '','', " & _
						  " decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr'), " & _
				  " sysdate LASTUPDATE, to_char(sysdate,'HH24:MI') LASTTIME, " & Month (Session("v_General_Charges")) & ", " & Year (Session("v_General_Charges")) & ", rownum + " & maxLineNo & " LINENO " & _
				  " FROM ( " & _
				  " SELECT TRANSDATE tdate, " & _
						   "JOBCODE jobcode, " & _ 
						   " GCMASTER_API.get_item(fcba, FIELDCODE) source, " & _
						   " GCMASTER_API.get_item(fcba, FIELDCODE) sourcetype, " & _
						   "FIELDCODE sourcecode, " & _
						   "'' sourcecode_status, " & _
						   "TARGETTYPE targettype, " & _
						   "TARGETCODE targetcode, " & _
						   "'' targetcode_status, " & _
						   "REFERENCENO referenceno, " & _
						   "SUM(QUANTITY) quantity, " & _
						   "0 output, " & _
						   "SUM(AMOUNT) amount, " & _ 
						   "FCBA," & _
						   "to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date " & _
				  " FROM LOCCB" & _
				  " WHERE fcBA='" & Session("fcBA") & "'" & _
				  "       and to_number(to_Char(TRANSDATE,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
				  "       and to_number(to_Char(TRANSDATE,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
				  " group by TRANSDATE, JOBCODE, GCMASTER_API.get_item(fcba, FIELDCODE), GCMASTER_API.get_item(fcba, FIELDCODE), FIELDCODE, '', TARGETTYPE, TARGETCODE, '', REFERENCENO,  FCBA)"
		'response.Write " qABC "& qABC
		response.Flush()		  
		connstr.Execute(qABC)
		
		Response.write("<font face=Verdana size=1 color=Green>Transfer to Activity Based Costing..., Success!.<br/><br/></font>")
		Response.Flush
	
	else
	
		Set objectRS = Server.CreateObject("ADODB.Recordset")
		objectSQL = " select nvl(sum(costtotal), 0) costtotal" & _
					"        , nvl(sum(alloc_db), 0) alloc_db" & _
					"        , nvl(sum(alloc_cr), 0)*-1 alloc_cr" & _
					" from (" & _
					"     select sum(costtotal) costtotal, 0 alloc_db , 0 alloc_cr" & _
					"     from costtotal_pertarget" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and year = " & Year(Session("v_General_Charges")) & "" & _
					"           and month = " & Month(Session("v_General_Charges")) & "" & _
					"           and targettype = 'LOC'" & _
					"     union all" & _
					"     select 0 costtotal, sum(amount) alloc_db , 0 alloc_cr" & _
					"     from loccb" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and to_number(to_char(transdate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					"           and to_number(to_char(transdate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					"           and jobcode <> '" & txtControlJob & "'" & _
					"     union all" & _
					"     select 0 costtotal,0 alloc_db ,  sum(amount) alloc_cr" & _
					"     from loccb" & _
					"     where fcba = '" & Session("fcBA") & "'" & _
					"           and to_number(to_char(transdate, 'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
					"           and to_number(to_char(transdate, 'mm')) = " & Month(Session("v_General_Charges")) & "" & _
					"           and jobcode = '" & txtControlJob & "'" & _
					" )"

			objectRS.Open objectSQL, connstr
			If Not objectRS.EOF Then
				Cost_In = Cdbl(objectRS("costtotal"))
				Alloc_Db = Cdbl(objectRS("alloc_db"))
				Alloc_Cr = Cdbl(objectRS("alloc_cr"))
			Else
			End If
			objectRS.close
			set objectRS = Nothing

			If Cost_In <> Alloc_Db Then
				err.number = 1001
				err.description = "Alokasi sisi debet tidak sesuai dengan Biaya yang Masuk! Cek Kembali Biaya yang masuk ke target LOC"
			ElseIf Cost_In <> Alloc_Cr Then
				err.number = 1001
				err.description = "Alokasi sisi credit tidak sesuai dengan Biaya yang Masuk! Cek Kembali Biaya yang masuk ke target LOC"
			ElseIf (Cost_In = Alloc_Db) And (Cost_In = Alloc_Cr) Then
				'----- Tranfer to Activity Costing
				response.Write "Transfer to Activity Based Costing..." & "<br>"
				Response.Flush

				' Get the Max LineNo of Cost Activity
				Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
				rsCostActivityMax.CursorType = 3
				rsCostActivityMax.CursorLocation = 3
												
				qCostActivityMax =  " select nvl(max(lineno), 0) maxLineNo from cost_activity " & _
									" where fcba = '" &  Session("fcBA") & "' " 
				rsCostActivityMax.Open qCostActivityMax, connstr
				
				If Not rsCostActivityMax.EOF Then
					maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))
				Else
					maxLineNo = 0
				End If

				rsCostActivityMax.close
				set rsCostActivityMax = Nothing
				' End of Get the Max LineNo of Cost Activity
				
				qABC = " INSERT INTO COST_ACTIVITY (" & _
					   " TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
					   " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,  " & _
												" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
												" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
												" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
												" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
												" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
												" LASTUPDATE, LASTTIME, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, LINENO ) " & _
					   " select TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
						  " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, sysdate PROCESS_DATE, " & _
								  " Cost_Activity_Info_Api.Get_Title_Reference_01 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),TARGETCODE, " & _
								  " Cost_Activity_Info_Api.Get_Title_Reference_02 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),referenceno, " & _
								  " '','', " & _
								  " '','', " & _
								  " '','', " & _
								  " decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr'), " & _
						  " sysdate LASTUPDATE, to_char(sysdate,'HH24:MI') LASTTIME, " & Month (Session("v_General_Charges")) & ", " & Year (Session("v_General_Charges")) & ", rownum + " & maxLineNo & " LINENO " & _
						  " FROM ( " & _
						  " SELECT TRANSDATE tdate, " & _
								   "JOBCODE jobcode, " & _ 
								   "'LOC' source, " & _
								   "'LOC' sourcetype, " & _
								   "FIELDCODE sourcecode, " & _
								   "'' sourcecode_status, " & _
								   "TARGETTYPE targettype, " & _
								   "TARGETCODE targetcode, " & _
								   "'' targetcode_status, " & _
								   "REFERENCENO referenceno, " & _
								   "SUM(QUANTITY) quantity, " & _
								   "0 output, " & _
								   "SUM(AMOUNT) amount, " & _ 
								   "FCBA," & _
								   "to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date " & _
						  " FROM LOCCB" & _
						  " WHERE fcBA='" & Session("fcBA") & "'" & _
						  "       and to_number(to_Char(TRANSDATE,'yyyy')) = " & Year(Session("v_General_Charges")) & "" & _
						  "       and to_number(to_Char(TRANSDATE,'mm')) = " & Month(Session("v_General_Charges")) & "" & _
						  " group by TRANSDATE, JOBCODE, 'LOC', 'LOC', FIELDCODE, '', TARGETTYPE, TARGETCODE, '', REFERENCENO,  FCBA)"
				'esponse.Write " qABC "& qABC
				response.Flush()		  
				connstr.Execute(qABC)
				
				Response.write("<font face=Verdana size=1 color=Green>Transfer to Activity Based Costing..., Success!.<br/><br/></font>")
				Response.Flush
			End If
		end if

End If
%>
<%
	Response.write "********** LOC Closing ***********<br/><br/>"
	Response.Flush
%>

<%if Session("Estate_Mill") = "E" then%>

	<%
		Response.Write "Closing Workshop" & "<br/>"
		Response.Flush
	%>
		<!--#include virtual="/p_loc_ws_closing.Asp"-->
	<%
		Response.Write "<font face=Verdana size=1 color=Green>" & "Closing Workshop Success!" & "<br/><br/></font>"
		Response.Flush
	%>

	<%
		Response.Write "Closing Vehicle" & "<br/>"
		Response.Flush
	%>
		<!--#include virtual="/p_loc_vh_closing.Asp"-->
	<%
		Response.Write "<font face=Verdana size=1 color=Green>" & "Closing Vehicle Success!" & "<br/><br/></font>"
		Response.Flush
	%>

	<%
		Response.Write "Closing Building" & "<br/>"
		Response.Flush
	%>
		<!--#include virtual="/p_loc_bm_closing.Asp"-->
	<%
		Response.Write "<font face=Verdana size=1 color=Green>" & "Closing Building Success!" & "<br/><br/></font>"
		Response.Flush
	%>

	<%
		Response.Write "Closing General Charges" & "<br/>"
		Response.Flush
	%>
		<!--#include virtual="/p_loc_gc_closing.Asp"-->
	<%
		Response.Write "<font face=Verdana size=1 color=Green>" & "Closing General Charges Success!" & "<br/><br/></font>"
		Response.Flush
	%>
	
	<%
		Response.Write "Delete Data Temporary" & "<br/>"
		Response.Flush
	%>
		<!--#include virtual="/p_deletedataupr_estate.asp"-->
	<%
		Response.Write "<font face=Verdana size=1 color=Green>" & "Delete Data Temporary Success!" & "<br/><br/></font>"
		Response.Flush
	%>
	
	<%Server.Execute("p_loc_mn_closing.Asp")%>


<%end if%>