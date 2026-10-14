<%

dim firstdate_, lastdate_

firstdate_ = textDate( V_Bulan & "/01/" & V_Tahun )
lastdate_ = textDate( V_Bulan & "/" & GetLastDay(V_Bulan, V_Tahun) & "/" & V_Tahun )

Response.Write "Process closing estate : " & firstdate_ & " s/d " & lastdate_ & "<br/>"

Response.Write "Generate Necessary Data<br/>"
Response.Flush

connstr.BeginTrans
ExecuteSQL = " delete" & _
			 " from EMPLOYEE_ALLOWANCES_TAB" & _
			 " where fcba = '" & Session("fcBA") & "'"

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " delete" & _
			 " from EMPLOYEE_ACTIVITIES_TAB" & _
			 " where fcba = '" & Session("fcBA") & "'"

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " delete from employee_activities_em_tab" & _
             " where fcba = '" & Session("fcBA") & "'"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " delete from employee_activities_es_tab" & _
             " where fcba = '" & Session("fcBA") & "'"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " insert into EMPLOYEE_ALLOWANCES_TAB" & _
			 " select fcba" & _
			 "        , transdate" & _
			 "        , allowdedcode" & _
			 "        , employeecode" & _
			 "        , employee_hist_api.get_sectionname(to_number(to_char(transdate, 'yyyy')), to_number(to_char(transdate, 'mm')), fcba, employeecode) section" & _
			 "        , employee_hist_api.get_presentgrade(to_number(to_char(transdate, 'yyyy')), to_number(to_char(transdate, 'mm')), fcba, employeecode) grade" & _
			 "        , amount" & _
			 " from (" & _
			 "     select fcba" & _
			 "            , last_day(adhocdate)transdate" & _
			 "            , allowdedcode" & _
			 "            , empcode employeecode" & _
			 "            , sum(amount) amount" & _
			 "     from empallded" & _
			 "     where fcba = '" & Session("fcBA") & "'" & _
			 "           and to_number(to_char(adhocdate, 'yyyy')) = " & V_Tahun & "" & _
			 "           and to_number(to_char(adhocdate, 'mm')) = " & V_Bulan & "" & _
			 "     group by fcba, last_day(adhocdate), allowdedcode, empcode" & _
			 "     having sum(amount) <> 0" & _
			 " ) where allowdedcode not in ('A45', 'A61', 'A59')"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " insert into EMPLOYEE_ACTIVITIES_TAB" & _
			 " select fcba" & _
			 "        , transdate" & _
			 "        , originaldate" & _
			 "        , sourcetype" & _
			 "        , originalgang" & _
			 "        , employeecode" & _
			 "        , employee_hist_api.get_sectionname(to_number(to_char(transdate, 'yyyy')), to_number(to_char(transdate, 'mm')), fcba, employeecode) section" & _
			 "        , employee_hist_api.get_presentgrade(to_number(to_char(transdate, 'yyyy')), to_number(to_char(transdate, 'mm')), fcba, employeecode) grade" & _
			 "        , jobcode" & _
			 "        , targettype" & _
			 "        , targetcode" & _
			 "        , mandays" & _
			 "        , unit" & _
			 "        , othrs" & _
			 " from (" & _
			 "     select fcba" & _
			 "            , last_day(fddate) transdate" & _
			 "            , fddate originaldate" & _
			 "            , 'ES' sourcetype" & _
			 "            , originalgang" & _
			 "            , employeecode" & _
			 "            , jobcode" & _
			 "            , locationtype targettype" & _
			 "            , locationcode targetcode" & _
			 "            , sum (mandays) mandays" & _
			 "            , sum (unit) unit" & _
			 "            , sum (othrs) othrs" & _
			 "     from ( select fcba" & _
			 "                   , fddate" & _
			 "                   , gangcode originalgang" & _
			 "                   , supervision_1 employeecode" & _
			 "                   , jobcode" & _
			 "                   , locationtype" & _
			 "                   , locationcode" & _
			 "                   , mandays" & _
			 "                   , unit" & _
			 "                   , othrs" & _
			 "            from attendance_gad_closing" & _
			 "            where fcba = '" & Session("fcBA") & "'" & _
			 "                  and supervision_1 is not null" & _
			 "            union all" & _
			 "            select fcba" & _
			 "                   , fddate" & _
			 "                   , gangcode originalgang" & _
			 "                   , supervision_2 employeecode" & _
			 "                   , jobcode" & _
			 "                   , locationtype" & _
			 "                   , locationcode" & _
			 "                   , mandays" & _
			 "                   , unit" & _
			 "                   , othrs" & _
			 "            from attendance_gad_closing" & _
			 "            where fcba = '" & Session("fcBA") & "'" & _
			 "                  and supervision_2 is not null" & _
			 "            union all" & _
			 "            select fcba" & _
			 "                   , fddate" & _
			 "                   , gangcode originalgang" & _
			 "                   , supervision_3 employeecode" & _
			 "                   , jobcode" & _
			 "                   , locationtype" & _
			 "                   , locationcode" & _
			 "                   , mandays" & _
			 "                   , unit" & _
			 "                   , othrs" & _
			 "            from attendance_gad_closing" & _
			 "            where fcba = '" & Session("fcBA") & "'" & _
			 "                  and supervision_3 is not null" & _
			 "            union all" & _
			 "            select fcba" & _
			 "                   , fddate" & _
			 "                   , gangcode originalgang" & _
			 "                   , supervision_4 employeecode" & _
			 "                   , jobcode" & _
			 "                   , locationtype" & _
			 "                   , locationcode" & _
			 "                   , mandays" & _
			 "                   , unit" & _
			 "                   , othrs" & _
			 "            from attendance_gad_closing" & _
			 "            where fcba = '" & Session("fcBA") & "'" & _
			 "                  and supervision_4 is not null" & _
			 "            )" & _
			 "     group by fcba" & _
			 "              , last_day(fddate)" & _
			 "              , fddate" & _
			 "              , originalgang" & _
			 "              , employeecode" & _
			 "              , jobcode" & _
			 "              , locationtype" & _
			 "              , locationcode" & _
			 "     union all" & _
			 "     select fcba" & _
			 "            , last_day(fddate) transdate" & _
			 "            , fddate originaldate" & _
			 "            , 'EM' sourcetype" & _
			 "            , gangcode originalgang" & _
			 "            , employeecode" & _
			 "            , jobcode" & _
			 "            , locationtype targettype" & _
			 "            , locationcode targetcode" & _
			 "            , sum(mandays) mandays" & _
			 "            , sum(unit) unit" & _
			 "            , sum(othrs) othrs" & _
			 "     from attendance_gad_closing" & _
			 "     where fcba = '" & Session("fcBA") & "'" & _
			 "     group by fcba" & _
			 "              , last_day(fddate)" & _
			 "              , fddate" & _
			 "              , gangcode" & _
			 "              , employeecode" & _
			 "              , jobcode" & _
			 "              , locationtype" & _
			 "              , locationcode" & _
			 " )"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " insert into employee_activities_em_tab" & _
             " select bb.*" & _
             " from ( select *" & _
             "        from ( select fcba" & _
             "                      , employeecode" & _
             "                      , listday" & _
             "                      , new_employee_closing_api.get_validsourcetype (employeecode, fcba, listday) sourcetype" & _
             "               from ( select distinct fcba" & _
             "                                      , employeecode" & _
             "                      from employee_allowances_tab" & _
             "                      where fcba = '" & Session("fcBA") & "'" & _
             "               ) a" & _
             "             , ( select to_date('" & firstdate_ & "','MM/DD/YYYY') + rownum -1 listday" & _
             "                 from all_objects" & _
             "                 where rownum <= to_date('" & lastdate_ & "','MM/DD/YYYY') - to_date('" & firstdate_ & "','MM/DD/YYYY') + 1" & _
             "               ) b" & _
             "        )" & _
             "        where sourcetype = 'EM'" & _
             " ) aa" & _
             " , ( select *" & _
             "     from employee_activities_tab" & _
             "     where fcba = '" & Session("fcBA") & "'" & _
             "           and sourcetype = 'EM') bb" & _
             " where aa.fcba = bb.fcba" & _
             "       and aa.employeecode = bb.employeecode" & _
             "       and bb.originaldate = aa.listday"



connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " insert into employee_activities_es_tab" & _
             " select bb.*" & _
             " from ( select *" & _
             "        from ( select fcba" & _
             "                      , employeecode" & _
             "                      , listday" & _
             "                      , new_employee_closing_api.get_validsourcetype (employeecode, fcba, listday) sourcetype" & _
             "               from ( select distinct fcba" & _
             "                                      , employeecode" & _
             "                      from employee_allowances_tab" & _
             "                      where fcba = '" & Session("fcBA") & "'" & _
             "               ) a" & _
             "             , ( select to_date('" & firstdate_ & "','MM/DD/YYYY') + rownum -1 listday" & _
             "                 from all_objects" & _
             "                 where rownum <= to_date('" & lastdate_ & "','MM/DD/YYYY') - to_date('" & firstdate_ & "','MM/DD/YYYY') + 1" & _
             "               ) b" & _
             "        )" & _
             "        where sourcetype = 'ES'" & _
             " ) aa" & _
             " , ( select *" & _
             "     from employee_activities_tab" & _
             "     where fcba = '" & Session("fcBA") & "'" & _
             "           and sourcetype = 'ES') bb" & _
             " where aa.fcba = bb.fcba" & _
             "       and aa.employeecode = bb.employeecode" & _
             "       and bb.originaldate = aa.listday"



connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " delete from EMPLOYEE_ALLOWANCES_OT_TAB" & _
			 " where fcba = '" & Session("fcBA") & "'"

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " delete from ATTENDANCE_GAD_OT_TAB" & _
			 " where fcba = '" & Session("fcBA") & "'"

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " delete from EMPLOYEE_ACTIVITIES_OT_TAB" & _
			 " where fcba = '" & Session("fcBA") & "'"

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " delete from EMPLOYEE_ACTIVITIES_OT_EM_TAB" & _
			 " where fcba = '" & Session("fcBA") & "'"

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " delete from EMPLOYEE_ACTIVITIES_OT_ES_TAB" & _
			 " where fcba = '" & Session("fcBA") & "'"

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If


connstr.BeginTrans
ExecuteSQL = " insert into EMPLOYEE_ALLOWANCES_OT_TAB" & _
			 " select fcba" & _
			 "        , transdate" & _
			 "        , originaldate" & _
			 "        , allowdedcode" & _
			 "        , employeecode" & _
			 "        , employee_hist_api.get_sectionname(to_number(to_char(transdate, 'yyyy')), to_number(to_char(transdate, 'mm')), fcba, employeecode) section" & _
			 "        , employee_hist_api.get_presentgrade(to_number(to_char(transdate, 'yyyy')), to_number(to_char(transdate, 'mm')), fcba, employeecode) grade" & _
			 "        , amount" & _
			 " from (" & _
			 "     select fcba" & _
			 "            , last_day(adhocdate)transdate" & _
			 "            , adhocdate originaldate" & _
			 "            , allowdedcode" & _
			 "            , empcode employeecode" & _
			 "            , sum(amount) amount" & _
			 "     from empallded" & _
			 "     where fcba = '" & Session("fcBA") & "'" & _
			 "           and to_number(to_char(adhocdate, 'yyyy')) = " & V_Tahun & "" & _
			 "           and to_number(to_char(adhocdate, 'mm')) = " & V_Bulan & "" & _
			 "			 and allowdedcode = 'A18' " & _
			 "     group by fcba, last_day(adhocdate), adhocdate, allowdedcode, empcode" & _
			 "     having sum(amount) <> 0" & _
			 " )  "


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " insert into ATTENDANCE_GAD_OT_TAB" & _
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
			 "       and to_number(to_char(fddate, 'yyyy')) = " & V_Tahun & "" & _
			 "       and to_number(to_char(fddate, 'mm')) = " & V_Bulan & "" & _
			 "       and attendance in ('KJ', 'WH', 'WS', 'KL')" & _
			 "       and othrs <> 0"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If


connstr.BeginTrans
ExecuteSQL = " insert into EMPLOYEE_ACTIVITIES_OT_TAB" & _
			 " select fcba" & _
			 "        , transdate" & _
			 "        , originaldate" & _
			 "        , sourcetype" & _
			 "        , originalgang" & _
			 "        , employeecode" & _
			 "        , employee_hist_api.get_sectionname(to_number(to_char(transdate, 'yyyy')), to_number(to_char(transdate, 'mm')), fcba, employeecode) section" & _
			 "        , employee_hist_api.get_presentgrade(to_number(to_char(transdate, 'yyyy')), to_number(to_char(transdate, 'mm')), fcba, employeecode) grade" & _
			 "        , jobcode" & _
			 "        , targettype" & _
			 "        , targetcode" & _
			 "        , mandays" & _
			 "        , unit" & _
			 "        , othrs" & _
			 " from (" & _
			 "     select fcba" & _
			 "            , last_day(fddate) transdate" & _
			 "            , fddate originaldate" & _
			 "            , 'ES' sourcetype" & _
			 "            , originalgang" & _
			 "            , employeecode" & _
			 "            , jobcode" & _
			 "            , locationtype targettype" & _
			 "            , locationcode targetcode" & _
			 "            , sum (mandays) mandays" & _
			 "            , sum (unit) unit" & _
			 "            , sum (othrs) othrs" & _
			 "     from ( select fcba" & _
			 "                   , fddate" & _
			 "                   , gangcode originalgang" & _
			 "                   , supervision_1 employeecode" & _
			 "                   , jobcode" & _
			 "                   , locationtype" & _
			 "                   , locationcode" & _
			 "                   , mandays" & _
			 "                   , unit" & _
			 "                   , othrs" & _
			 "            from ATTENDANCE_GAD_OT_TAB" & _
			 "            where fcba = '" & Session("fcBA") & "'" & _
			 "                  and supervision_1 is not null" & _
			 "            union all" & _
			 "            select fcba" & _
			 "                   , fddate" & _
			 "                   , gangcode originalgang" & _
			 "                   , supervision_2 employeecode" & _
			 "                   , jobcode" & _
			 "                   , locationtype" & _
			 "                   , locationcode" & _
			 "                   , mandays" & _
			 "                   , unit" & _
			 "                   , othrs" & _
			 "            from ATTENDANCE_GAD_OT_TAB" & _
			 "            where fcba = '" & Session("fcBA") & "'" & _
			 "                  and supervision_2 is not null" & _
			 "            union all" & _
			 "            select fcba" & _
			 "                   , fddate" & _
			 "                   , gangcode originalgang" & _
			 "                   , supervision_3 employeecode" & _
			 "                   , jobcode" & _
			 "                   , locationtype" & _
			 "                   , locationcode" & _
			 "                   , mandays" & _
			 "                   , unit" & _
			 "                   , othrs" & _
			 "            from ATTENDANCE_GAD_OT_TAB" & _
			 "            where fcba = '" & Session("fcBA") & "'" & _
			 "                  and supervision_3 is not null" & _
			 "            union all" & _
			 "            select fcba" & _
			 "                   , fddate" & _
			 "                   , gangcode originalgang" & _
			 "                   , supervision_4 employeecode" & _
			 "                   , jobcode" & _
			 "                   , locationtype" & _
			 "                   , locationcode" & _
			 "                   , mandays" & _
			 "                   , unit" & _
			 "                   , othrs" & _
			 "            from ATTENDANCE_GAD_OT_TAB" & _
			 "            where fcba = '" & Session("fcBA") & "'" & _
			 "                  and supervision_4 is not null" & _
			 "            )" & _
			 "     group by fcba" & _
			 "              , last_day(fddate)" & _
			 "              , fddate" & _
			 "              , originalgang" & _
			 "              , employeecode" & _
			 "              , jobcode" & _
			 "              , locationtype" & _
			 "              , locationcode" & _
			 "     union all" & _
			 "     select fcba" & _
			 "            , last_day(fddate) transdate" & _
			 "            , fddate originaldate" & _
			 "            , 'EM' sourcetype" & _
			 "            , gangcode originalgang" & _
			 "            , employeecode" & _
			 "            , jobcode" & _
			 "            , locationtype targettype" & _
			 "            , locationcode targetcode" & _
			 "            , sum(mandays) mandays" & _
			 "            , sum(unit) unit" & _
			 "            , sum(othrs) othrs" & _
			 "     from ATTENDANCE_GAD_OT_TAB" & _
			 "     where fcba = '" & Session("fcBA") & "'" & _
			 "     group by fcba" & _
			 "              , last_day(fddate)" & _
			 "              , fddate" & _
			 "              , gangcode" & _
			 "              , employeecode" & _
			 "              , jobcode" & _
			 "              , locationtype" & _
			 "              , locationcode" & _
			 " )"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If


connstr.BeginTrans
ExecuteSQL = " insert into employee_activities_ot_em_tab" & _
             " select bb.*" & _
             " from ( select *" & _
             "        from employee_allowances_ot_tab" & _
             "        where fcba = '" & Session("fcBA") & "'" & _
             "              and NEW_EMPLOYEE_CLOSING_API.get_validsourcetype(employeecode, fcba, originaldate) = 'EM'" & _
             " )aa" & _
             " inner join employee_activities_ot_tab bb" & _
             " on (aa.fcba = bb.fcba and aa.originaldate = bb.originaldate and aa.employeecode = bb.employeecode)" & _
             " where bb.fcba = '" & Session("fcBA") & "'" & _
			 "       and bb.sourcetype = 'EM'"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " insert into employee_activities_ot_em_tab" & _
             " select bb.fcba" & _
             "        , bb.transdate" & _
             "        , bb.originaldate" & _
             "        , bb.sourcetype" & _
             "        , bb.originalgang" & _
             "        , bb.employeecode" & _
             "        , bb.section" & _
             "        , bb.grade" & _
             "        , bb.jobcode" & _
             "        , bb.targettype" & _
             "        , bb.targetcode" & _
             "        , bb.mandays" & _
             "        , bb.unit" & _
             "        , bb.mandays othrs" & _
             " from (" & _
             "     select *" & _
             "     from employee_allowances_ot_tab" & _
             "     where fcba||originaldate||employeecode in (" & _
             "         select fcba||originaldate||employeecode listdata" & _
             "         from employee_allowances_ot_tab" & _
             "         where fcba = '" & Session("fcBA") & "'" & _
             "               and new_employee_closing_api.get_validsourcetype(employeecode, fcba, originaldate) = 'EM'" & _
             "         minus" & _
             "         select fcba||originaldate||employeecode listdata" & _
             "         from employee_activities_ot_em_tab" & _
             "         where fcba = '" & Session("fcBA") & "'" & _
             "     )" & _
             " ) aa" & _
             " inner join employee_activities_em_tab bb" & _
             " on (aa.fcba = bb.fcba and aa.originaldate = bb.originaldate and aa.employeecode = bb.employeecode)" & _
             " where bb.fcba = '" & Session("fcBA") & "'"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " insert into employee_activities_ot_es_tab" & _
             " select bb.*" & _
             " from ( select *" & _
             "        from employee_allowances_ot_tab" & _
             "        where fcba = '" & Session("fcBA") & "'" & _
             "              and NEW_EMPLOYEE_CLOSING_API.get_validsourcetype(employeecode, fcba, originaldate) = 'ES'" & _
             " )aa" & _
             " inner join employee_activities_ot_tab bb" & _
             " on (aa.fcba = bb.fcba and aa.originaldate = bb.originaldate and aa.employeecode = bb.employeecode)" & _
             " where bb.fcba = '" & Session("fcBA") & "'" & _
			 "       and bb.sourcetype = 'ES'"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " insert into employee_activities_ot_es_tab" & _
             " select bb.fcba" & _
             "        , bb.transdate" & _
             "        , bb.originaldate" & _
             "        , bb.sourcetype" & _
             "        , bb.originalgang" & _
             "        , bb.employeecode" & _
             "        , bb.section" & _
             "        , bb.grade" & _
             "        , bb.jobcode" & _
             "        , bb.targettype" & _
             "        , bb.targetcode" & _
             "        , bb.mandays" & _
             "        , bb.unit" & _
             "        , bb.mandays othrs" & _
             " from (" & _
             "     select *" & _
             "     from employee_allowances_ot_tab" & _
             "     where fcba||originaldate||employeecode in (" & _
             "         select fcba||originaldate||employeecode listdata" & _
             "         from employee_allowances_ot_tab" & _
             "         where fcba = '" & Session("fcBA") & "'" & _
             "               and new_employee_closing_api.get_validsourcetype(employeecode, fcba, originaldate) = 'ES'" & _
             "         minus" & _
             "         select fcba||originaldate||employeecode listdata" & _
             "         from employee_activities_ot_es_tab" & _
             "         where fcba = '" & Session("fcBA") & "'" & _
             "     )" & _
             " ) aa" & _
             " inner join employee_activities_es_tab bb" & _
             " on (aa.fcba = bb.fcba and aa.originaldate = bb.originaldate and aa.employeecode = bb.employeecode)" & _
             " where bb.fcba = '" & Session("fcBA") & "'"


connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If


'**********************************'
'            LEMBUR                '
'**********************************'

Response.write("<font face=Verdana size=1 color=Green>Generate Necessary Data, Success! " & " Time: " & time() & " .<br/><br/></font>")
Response.Flush


Response.Write "Borongan" & "<br/>"
Response.Flush

q_borongan = " select fcba" & _
             "        , transdate" & _
             "        , allowdedcode" & _
             "        , employeecode" & _
             "        , section" & _
             "        , grade" & _
             "        , amount" & _
             "        , 'EM' sourcetype" & _
             "        , jobcode" & _
             "        , targettype" & _
             "        , targetcode" & _
             "        , new_employee_closing_api.get_detailalokasi(allowdedcode, targettype, jobcode, 1, fcba, section, grade, 'EM') alokasi_jobcode" & _
             "        , new_employee_closing_api.get_detailalokasi(allowdedcode, targettype, targettype, 2, fcba, section, grade, 'EM') alokasi_tt" & _
             "        , new_employee_closing_api.get_detailalokasi(allowdedcode, targettype, targetcode, 3, fcba, section, grade, 'EM') alokasi_tc" & _
             " from ( " & _
             "     select fcba" & _
             "            , fddate transdate" & _
             "            , employeecode" & _
             "            , employee_hist_api.get_sectionname(to_number(to_char(fddate, 'yyyy')), to_number(to_char(fddate, 'mm')), fcba, employeecode) section" & _
             "            , employee_hist_api.get_presentgrade(to_number(to_char(fddate, 'yyyy')), to_number(to_char(fddate, 'mm')), fcba, employeecode) grade" & _
             "            , jobcode" & _
             "            , locationtype targettype" & _
             "            , locationcode targetcode" & _
             "            , decode(inclpayprocess, 'TRUE', decode(postingstatus, 'TRUE', 'A45', 'A45'), 'A45') allowdedcode" & _
             "            , sum(totalamount) amount" & _
			 "            , in_data" & _
             "     from ( select fcba" & _
             "                   , employeecode" & _
             "                   , locationtype" & _
             "                   , locationcode" & _
             "                   , jobcode" & _
             "                   , fddate" & _
             "                   , (nvl(unit,0) * nvl(rate,0))+(nvl(kg_brondolan,0) * nvl(rate,0)) totalamount" & _
             "                   , process_posting_api.getpostingvalue('PAYROLL', 'A45', 'JOB', jobcode, fcba) postingstatus" & _
             "                   , allowded_api.getincludedpayrollprocessing('A45', fcba) inclpayprocess" & _
			 "                   , employee_hist_api.in_data(to_number(to_char(fddate, 'yyyy')), to_number(to_char(fddate, 'mm')), fcba, employeecode) in_data" & _
             "            from attendance_gad" & _
             "            where " & _
             "                  rate <> 0" & _
             "                  and mandays = 0" & _
             "                  and attendance = 'KB'" & _
             "                  and to_number(to_char(fddate,'mm')) =  " & V_Bulan & "" & _
             "                  and to_number(to_char(fddate,'yyyy')) =  " & V_Tahun & "" & _
             "                  and fcba = '" & Session("fcBA") & "'" & _
             "     )     " & _
             "     group by fcba" & _
             "              , employeecode" & _
             "              , locationtype" & _
             "              , locationcode" & _
             "              , jobcode" & _
             "              , fddate" & _
             "              , decode(inclpayprocess,'TRUE',decode(postingstatus,'TRUE','A45','A45'),'A45')" & _
			 "              , in_data" & _
             " ) where in_data = 'YES'" & _
			 "             and empallded_api.in_data(fcba, transdate, allowdedcode, employeecode) = 'YES'"
			 
			 
connstr.BeginTrans
ExecuteSQL = " delete from employee_borongan_tab where fcba = '" & Session("fcBA") & "'"

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If

connstr.BeginTrans
ExecuteSQL = " insert into employee_borongan_tab " & q_borongan

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If
			  
arrRS = GetRSArray(q_borongan)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		If Cdbl(arrRS(6, i)) <> 0 Then
			connstr.BeginTrans
			maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
			qInsertDB = " insert into employeecb (empcode" & _
						"                         , sourcetype" & _
						"                         , targettype" & _
						"                         , targetcode" & _
						"                         , jobcode" & _
						"                         , amount" & _
						"                         , transdate" & _
						"                         , referenceno" & _
						"                         , fcentry" & _
						"                         , fcedit" & _
						"                         , fcip" & _
						"                         , fcba" & _
						"                         , lastupdate" & _
						"                         , lasttime" & _
						"                         , lineno)" & _
						" values ('" & arrRS(3, i) & "'" & _
						"         , '" & arrRS(7, i) & "'" & _
						"         , '" & arrRS(12, i) & "'" & _
						"         , '" & arrRS(13, i) & "'" & _
						"         , '" & arrRS(11, i) & "'" & _
						"         , " & Cdbl(arrRS(6, i)) & "" & _
						"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
						"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
						"         , '" & Session("CodeName") & "'" & _
						"         , '" & Session("CodeName") & "'" & _
						"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						"         , '" & Session("fcBA") & "'" & _
						"         , sysdate" & _
						"         , to_char(sysdate, 'HH24:MI')" & _
						"         , " & maxLineNoEmployeeCB & ")"
			' response.Write(qInsertDB)			  
			connStr.execute(qInsertDB)
			
			maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
			qInsertCR = " insert into employeecb (empcode" & _
						"                         , sourcetype" & _
						"                         , targettype" & _
						"                         , targetcode" & _
						"                         , jobcode" & _
						"                         , amount" & _
						"                         , transdate" & _
						"                         , referenceno" & _
						"                         , fcentry" & _
						"                         , fcedit" & _
						"                         , fcip" & _
						"                         , fcba" & _
						"                         , lastupdate" & _
						"                         , lasttime" & _
						"                         , lineno)" & _
						" values ('" & arrRS(3, i) & "'" & _
						"         , '" & arrRS(7, i) & "'" & _
						"         , 'CB'" & _
						"         , 'BS'" & _
						"         , '" & txtControlJob & "'" & _
						"         , " & Cdbl(arrRS(6, i)) * -1 & "" & _
						"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
						"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
						"         , '" & Session("CodeName") & "'" & _
						"         , '" & Session("CodeName") & "'" & _
						"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						"         , '" & Session("fcBA") & "'" & _
						"         , sysdate" & _
						"         , to_char(sysdate, 'HH24:MI')" & _
						"         , " & maxLineNoEmployeeCB & ")"
			' response.Write(qInsertCR)			  
			connStr.execute(qInsertCR)

			If Err.Number = 0 Then
				connstr.CommitTrans
				Response.Flush
			Else
				connstr.RollbackTrans
				Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
				Response.Flush
			End If
		End If
	Next
	Erase arrRS
End If

Response.write("<font face=Verdana size=1 color=Green>Borongan, Success! " & " Time: " & time() & " ..<br/><br/></font>")
Response.Flush


Response.Write "Lembur" & "<br/>"
Response.Flush

qlembur = " select aa.fcba" & _
          "        , aa.transdate" & _
          "        , aa.allowdedcode" & _
          "        , aa.employeecode" & _
          "        , aa.section" & _
          "        , aa.grade" & _
          "        , new_employee_closing_api.get_amount_ot_monthly(aa.fcba" & _
          "                                                         , aa.employeecode" & _
          "                                                         , aa.transdate) amount" & _
          "        , aa.sourcetype" & _
          "        , new_employee_closing_api.get_basis_ot(aa.fcba" & _
          "                                                , aa.employeecode" & _
          "                                                , aa.transdate) basis" & _
          "        , new_employee_closing_api.get_basis_ot_per_activity (aa.fcba" & _
          "                                                              , aa.employeecode" & _
          "                                                              , aa.originaldate" & _
          "                                                              , aa.sourcetype" & _
          "                                                              , bb.jobcode" & _
          "                                                              , bb.targettype" & _
          "                                                              , bb.targetcode) act" & _
          "        , allowded_api.getallded (aa.allowdedcode, aa.fcba) allded" & _
          "        , new_employee_closing_api.get_detailalokasi (aa.allowdedcode" & _
          "                                                      , bb.targettype" & _
          "                                                      , bb.jobcode" & _
          "                                                      , 1" & _
          "                                                      , aa.fcba" & _
          "                                                      , aa.section" & _
          "                                                      , aa.grade" & _
          "                                                      , aa.sourcetype) alokasi_jobcode" & _
          "        , new_employee_closing_api.get_detailalokasi (aa.allowdedcode" & _
          "                                                      , bb.targettype" & _
          "                                                      , bb.targettype" & _
          "                                                      , 2" & _
          "                                                      , aa.fcba" & _
          "                                                      , aa.section" & _
          "                                                      , aa.grade" & _
          "                                                      , aa.sourcetype) alokasi_tt" & _
          "        , new_employee_closing_api.get_detailalokasi (aa.allowdedcode" & _
          "                                                      , bb.targettype" & _
          "                                                      , bb.targetcode" & _
          "                                                      , 3" & _
          "                                                      , aa.fcba" & _
          "                                                      , aa.section" & _
          "                                                      , aa.grade" & _
          "                                                      , aa.sourcetype) alokasi_tc" & _
          " from (" & _
          "     select fcba" & _
          "            , transdate" & _
          "            , originaldate" & _
          "            , allowdedcode" & _
          "            , employeecode" & _
          "            , section" & _
          "            , grade" & _
          "            , amount" & _
          "            , new_employee_closing_api.get_validsourcetype (employeecode, fcba, originaldate) sourcetype" & _
          "     from employee_allowances_ot_tab" & _
          "     where fcba = '" & Session("fcBA") & "'" & _
          "           and new_employee_closing_api.get_uprtype (allowdedcode, section, grade) = '01'" & _
          " ) aa" & _
          " , (   " & _
          "     select DISTINCT fcba" & _
          "                     , originaldate" & _
          "                     , employeecode" & _
          "                     , jobcode" & _
          "                     , targettype" & _
          "                     , targetcode" & _
          "                     , 'EM' sourcetype" & _
          "     from employee_activities_ot_em_tab" & _
          "     where fcba = '" & Session("fcBA") & "'" & _
          "     union all" & _
          "     select DISTINCT fcba" & _
          "                     , originaldate" & _
          "                     , employeecode" & _
          "                     , jobcode" & _
          "                     , targettype" & _
          "                     , targetcode" & _
          "                     , 'ES' sourcetype" & _
          "     from employee_activities_ot_es_tab" & _
          "     where fcba = '" & Session("fcBA") & "'" & _
          " ) bb" & _
          " where aa.fcba = bb.fcba" & _
          "       and aa.originaldate = bb.originaldate" & _
          "       and aa.employeecode = bb.employeecode" & _
          "       and aa.sourcetype = bb.sourcetype"
			  
arrRS = GetRSArray(qlembur)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		If Cdbl(arrRS(8, i)) <> 0 Then
			nAlokasi = 0
			nAlokasi = Round(Cdbl(arrRS(6, i))/Cdbl(arrRS(8, i))*Cdbl(arrRS(9, i)), 0)
			If nAlokasi <> 0 Then						
				connstr.BeginTrans
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertDB = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , '" & arrRS(12, i) & "'" & _
							"         , '" & arrRS(13, i) & "'" & _
							"         , '" & arrRS(11, i) & "'" & _
							"         , " & nAlokasi & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertDB)			  
				connStr.execute(qInsertDB)
				
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertCR = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , 'CB'" & _
							"         , 'BS'" & _
							"         , '" & txtControlJob & "'" & _
							"         , " & nAlokasi * -1 & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)

				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
					Response.Flush
				End If
			End If
		End If
	Next
	Erase arrRS
End If

Response.write("<font face=Verdana size=1 color=Green>Lembur, Success! " & " Time: " & time() & " ..<br/><br/></font>")
Response.Flush



Response.Write "UPR 02" & "<br/>"
Response.Flush

qupr_02 = " select aa.fcba" & _
		  "        , aa.transdate" & _
		  "        , aa.allowdedcode" & _
		  "        , aa.employeecode" & _
		  "        , aa.section" & _
		  "        , aa.grade" & _
		  "        , aa.amount" & _
		  "        , bb.sourcetype" & _
		  "        , new_employee_closing_api.get_basis(aa.fcba, aa.employeecode, aa.transdate, 'EMES') basis" & _
		  "        , new_employee_closing_api.get_basis_per_st(aa.fcba, aa.employeecode, aa.transdate, bb.sourcetype) act" & _
		  "        , allowded_api.getallded(aa.allowdedcode, aa.fcba) allded" & _
		  "        , new_employee_closing_api.get_detailalokasi(aa.allowdedcode, 1, aa.fcba, aa.section, aa.grade, bb.sourcetype) alokasi_jobcode" & _
		  "        , new_employee_closing_api.get_detailalokasi(aa.allowdedcode, 2, aa.fcba, aa.section, aa.grade, bb.sourcetype) alokasi_tt" & _
		  "        , new_employee_closing_api.get_detailalokasi(aa.allowdedcode, 3, aa.fcba, aa.section, aa.grade, bb.sourcetype) alokasi_tc" & _
		  " from (" & _
		  "     select * from employee_allowances_tab" & _
		  "     where fcba = '" & Session("fcBA") & "'" & _
		  "     and new_employee_closing_api.get_uprtype(allowdedcode, section, grade) = '02'" & _
		  " ) aa,( select 'EM' sourcetype from dual" & _
		  "        union all" & _
		  "        select 'ES' sourcetype from dual" & _
		  "      ) bb"
			  
arrRS = GetRSArray(qupr_02)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		If Cdbl(arrRS(8, i)) <> 0 Then
			nAlokasi = 0
			nAlokasi = Round(Cdbl(arrRS(6, i))/Cdbl(arrRS(8, i))*Cdbl(arrRS(9, i)), 0)
			If nAlokasi <> 0 Then						
				connstr.BeginTrans
				If Cstr(arrRS(10, i)) = "A" Then
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertDB = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , '" & arrRS(12, i) & "'" & _
								"         , '" & arrRS(13, i) & "'" & _
								"         , '" & arrRS(11, i) & "'" & _
								"         , " & nAlokasi & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)
					
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertCR = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , " & nAlokasi * -1 & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
				ElseIf Cstr(arrRS(10, i)) = "D" Then
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertDB = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , " & nAlokasi * -1 & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)
					
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertCR = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , '" & arrRS(12, i) & "'" & _
								"         , '" & arrRS(13, i) & "'" & _
								"         , '" & arrRS(11, i) & "'" & _
								"         , " & nAlokasi & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
				Else
				End If

				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
					Response.Flush
				End If
			End If
		End If
	Next
	Erase arrRS
End If

Response.write("<font face=Verdana size=1 color=Green> UPR 02, Success! " & " Time: " & time() & " ..<br/><br/></font>")
Response.Flush

Response.Write "UPR 01 - ES" & "<br/>"
Response.Flush

qupr_01 = " SELECT aaa.fcba," & _
		  "        aaa.transdate," & _
		  "        aaa.allowdedcode," & _
		  "        aaa.employeecode," & _
		  "        aaa.section," & _
		  "        aaa.grade," & _
		  "        aaa.amount," & _
		  "        aaa.sourcetype," & _
		  "        new_employee_closing_api.get_basis (aaa.fcba," & _
		  "                                            aaa.employeecode," & _
		  "                                            aaa.transdate," & _
		  "                                            aaa.sourcetype)" & _
		  "           basis," & _
		  "        new_employee_closing_api.get_basis_per_activity (aaa.fcba," & _
		  "                                                         aaa.employeecode," & _
		  "                                                         aaa.originaldate," & _
		  "                                                         aaa.sourcetype," & _
		  "                                                         bbb.jobcode," & _
		  "                                                         bbb.targettype," & _
		  "                                                         bbb.targetcode)" & _
		  "           act," & _
		  "        allowded_api.getallded (aaa.allowdedcode, aaa.fcba) allded," & _
		  "        new_employee_closing_api.get_detailalokasi (aaa.allowdedcode," & _
		  "                                                    bbb.targettype," & _
		  "                                                    bbb.jobcode," & _
		  "                                                    1," & _
		  "                                                    aaa.fcba," & _
		  "                                                    aaa.section," & _
		  "                                                    aaa.grade," & _
		  "                                                    aaa.sourcetype)" & _
		  "           alokasi_jobcode," & _
		  "        new_employee_closing_api.get_detailalokasi (aaa.allowdedcode," & _
		  "                                                    bbb.targettype," & _
		  "                                                    bbb.targettype," & _
		  "                                                    2," & _
		  "                                                    aaa.fcba," & _
		  "                                                    aaa.section," & _
		  "                                                    aaa.grade," & _
		  "                                                    aaa.sourcetype)" & _
		  "           alokasi_tt," & _
		  "        new_employee_closing_api.get_detailalokasi (aaa.allowdedcode," & _
		  "                                                    bbb.targettype," & _
		  "                                                    bbb.targetcode," & _
		  "                                                    3," & _
		  "                                                    aaa.fcba," & _
		  "                                                    aaa.section," & _
		  "                                                    aaa.grade," & _
		  "                                                    aaa.sourcetype)" & _
		  "           alokasi_tc" & _
		  "   FROM (SELECT bb.fcba," & _
		  "                bb.transdate," & _
		  "                aa.listday originaldate," & _
		  "                aa.sourcetype," & _
		  "                bb.employeecode," & _
		  "                bb.section," & _
		  "                bb.grade," & _
		  "                bb.allowdedcode," & _
		  "                bb.amount" & _
		  "           FROM (SELECT *" & _
		  "                   FROM (SELECT fcba," & _
		  "                                employeecode," & _
		  "                                listday," & _
		  "                                new_employee_closing_api.get_validsourcetype (" & _
		  "                                   employeecode," & _
		  "                                   fcba," & _
		  "                                   listday)" & _
		  "                                   sourcetype" & _
		  "                           FROM (SELECT DISTINCT fcba, employeecode" & _
		  "                                   FROM employee_allowances_tab" & _
		  "                                  WHERE fcba = '" & Session("fcBA") & "') a," & _
		  "                                (SELECT   TO_DATE ('" & firstdate_ & "', 'MM/DD/YYYY')" & _
		  "                                        + ROWNUM" & _
		  " -1" & _
		  "                                           listday" & _
		  "                                   FROM all_objects" & _
		  "                                  WHERE ROWNUM <=" & _
		  "                                             TO_DATE ('" & lastdate_ & "'," & _
		  "                                                      'MM/DD/YYYY')" & _
		  "                                           - TO_DATE ('" & firstdate_ & "'," & _
		  "                                                      'MM/DD/YYYY')" & _
		  "                                           + 1) b)" & _
		  "                  WHERE sourcetype = 'ES') aa," & _
		  "                (SELECT *" & _
		  "                   FROM employee_allowances_tab" & _
		  "                  WHERE     fcba = '" & Session("fcBA") & "'" & _
		  "                        AND new_employee_closing_api.get_uprtype (" & _
		  "                               allowdedcode," & _
		  "                               section," & _
		  "                               grade) = '01') bb" & _
		  "          WHERE aa.fcba = bb.fcba AND aa.employeecode = bb.employeecode and bb.allowdedcode <> 'A18') aaa," & _
		  "        (SELECT DISTINCT fcba," & _
		  "                         originaldate," & _
		  "                         employeecode," & _
		  "                         jobcode," & _
		  "                         targettype," & _
		  "                         targetcode" & _
		  "           FROM employee_activities_es_tab" & _
		  "          WHERE fcba = '" & Session("fcBA") & "') bbb" & _
		  "  WHERE     aaa.fcba = bbb.fcba" & _
		  "        AND aaa.employeecode = bbb.employeecode" & _
		  "        AND aaa.originaldate = bbb.originaldate"
			  
arrRS = GetRSArray(qupr_01)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		If Cdbl(arrRS(8, i)) <> 0 Then
			nAlokasi = 0
			nAlokasi = Round(Cdbl(arrRS(6, i))/Cdbl(arrRS(8, i))*Cdbl(arrRS(9, i)), 0)
			If nAlokasi <> 0 Then						
				connstr.BeginTrans
				If Cstr(arrRS(10, i)) = "A" Then
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertDB = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , '" & arrRS(12, i) & "'" & _
								"         , '" & arrRS(13, i) & "'" & _
								"         , '" & arrRS(11, i) & "'" & _
								"         , " & nAlokasi & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)
					
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertCR = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , " & nAlokasi * -1 & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
				ElseIf Cstr(arrRS(10, i)) = "D" Then
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertDB = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , " & nAlokasi * -1 & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)
					
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertCR = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , '" & arrRS(12, i) & "'" & _
								"         , '" & arrRS(13, i) & "'" & _
								"         , '" & arrRS(11, i) & "'" & _
								"         , " & nAlokasi & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
				Else
				End If

				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
					Response.Flush
				End If
			End If
		End If
	Next
	Erase arrRS
End If

Response.write("<font face=Verdana size=1 color=Green> UPR 01 - ES, Success! " & " Time: " & time() & " ..<br/><br/></font>")
Response.Flush

Response.Write "UPR 01 - EM" & "<br/>"
Response.Flush

qupr_01 = " SELECT aaa.fcba," & _
		  "        aaa.transdate," & _
		  "        aaa.allowdedcode," & _
		  "        aaa.employeecode," & _
		  "        aaa.section," & _
		  "        aaa.grade," & _
		  "        aaa.amount," & _
		  "        aaa.sourcetype," & _
		  "        new_employee_closing_api.get_basis (aaa.fcba," & _
		  "                                            aaa.employeecode," & _
		  "                                            aaa.transdate," & _
		  "                                            aaa.sourcetype)" & _
		  "           basis," & _
		  "        new_employee_closing_api.get_basis_per_activity (aaa.fcba," & _
		  "                                                         aaa.employeecode," & _
		  "                                                         aaa.originaldate," & _
		  "                                                         aaa.sourcetype," & _
		  "                                                         bbb.jobcode," & _
		  "                                                         bbb.targettype," & _
		  "                                                         bbb.targetcode)" & _
		  "           act," & _
		  "        allowded_api.getallded (aaa.allowdedcode, aaa.fcba) allded," & _
		  "        new_employee_closing_api.get_detailalokasi (aaa.allowdedcode," & _
		  "                                                    bbb.targettype," & _
		  "                                                    bbb.jobcode," & _
		  "                                                    1," & _
		  "                                                    aaa.fcba," & _
		  "                                                    aaa.section," & _
		  "                                                    aaa.grade," & _
		  "                                                    aaa.sourcetype)" & _
		  "           alokasi_jobcode," & _
		  "        new_employee_closing_api.get_detailalokasi (aaa.allowdedcode," & _
		  "                                                    bbb.targettype," & _
		  "                                                    bbb.targettype," & _
		  "                                                    2," & _
		  "                                                    aaa.fcba," & _
		  "                                                    aaa.section," & _
		  "                                                    aaa.grade," & _
		  "                                                    aaa.sourcetype)" & _
		  "           alokasi_tt," & _
		  "        new_employee_closing_api.get_detailalokasi (aaa.allowdedcode," & _
		  "                                                    bbb.targettype," & _
		  "                                                    bbb.targetcode," & _
		  "                                                    3," & _
		  "                                                    aaa.fcba," & _
		  "                                                    aaa.section," & _
		  "                                                    aaa.grade," & _
		  "                                                    aaa.sourcetype)" & _
		  "           alokasi_tc" & _
		  "   FROM (SELECT bb.fcba," & _
		  "                bb.transdate," & _
		  "                aa.listday originaldate," & _
		  "                aa.sourcetype," & _
		  "                bb.employeecode," & _
		  "                bb.section," & _
		  "                bb.grade," & _
		  "                bb.allowdedcode," & _
		  "                bb.amount" & _
		  "           FROM (SELECT *" & _
		  "                   FROM (SELECT fcba," & _
		  "                                employeecode," & _
		  "                                listday," & _
		  "                                new_employee_closing_api.get_validsourcetype (" & _
		  "                                   employeecode," & _
		  "                                   fcba," & _
		  "                                   listday)" & _
		  "                                   sourcetype" & _
		  "                           FROM (SELECT DISTINCT fcba, employeecode" & _
		  "                                   FROM employee_allowances_tab" & _
		  "                                  WHERE fcba = '" & Session("fcBA") & "') a," & _
		  "                                (SELECT   TO_DATE ('" & firstdate_ & "', 'MM/DD/YYYY')" & _
		  "                                        + ROWNUM" & _
		  " -1" & _
		  "                                           listday" & _
		  "                                   FROM all_objects" & _
		  "                                  WHERE ROWNUM <=" & _
		  "                                             TO_DATE ('" & lastdate_ & "'," & _
		  "                                                      'MM/DD/YYYY')" & _
		  "                                           - TO_DATE ('" & firstdate_ & "'," & _
		  "                                                      'MM/DD/YYYY')" & _
		  "                                           + 1) b)" & _
		  "                  WHERE sourcetype = 'EM') aa," & _
		  "                (SELECT *" & _
		  "                   FROM employee_allowances_tab" & _
		  "                  WHERE     fcba = '" & Session("fcBA") & "'" & _
		  "                        AND new_employee_closing_api.get_uprtype (" & _
		  "                               allowdedcode," & _
		  "                               section," & _
		  "                               grade) = '01') bb" & _
		  "          WHERE aa.fcba = bb.fcba AND aa.employeecode = bb.employeecode and bb.allowdedcode <> 'A18') aaa," & _
		  "        (SELECT DISTINCT fcba," & _
		  "                         originaldate," & _
		  "                         employeecode," & _
		  "                         jobcode," & _
		  "                         targettype," & _
		  "                         targetcode" & _
		  "           FROM employee_activities_em_tab" & _
		  "          WHERE fcba = '" & Session("fcBA") & "') bbb" & _
		  "  WHERE     aaa.fcba = bbb.fcba" & _
		  "        AND aaa.employeecode = bbb.employeecode" & _
		  "        AND aaa.originaldate = bbb.originaldate"
			  
arrRS = GetRSArray(qupr_01)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		If Cdbl(arrRS(8, i)) <> 0 Then
			nAlokasi = 0
			nAlokasi = Round(Cdbl(arrRS(6, i))/Cdbl(arrRS(8, i))*Cdbl(arrRS(9, i)), 0)
			If nAlokasi <> 0 Then						
				connstr.BeginTrans
				If Cstr(arrRS(10, i)) = "A" Then
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertDB = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , '" & arrRS(12, i) & "'" & _
								"         , '" & arrRS(13, i) & "'" & _
								"         , '" & arrRS(11, i) & "'" & _
								"         , " & nAlokasi & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)
					
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertCR = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , " & nAlokasi * -1 & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
				ElseIf Cstr(arrRS(10, i)) = "D" Then
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertDB = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , 'CB'" & _
								"         , 'BS'" & _
								"         , '" & txtControlJob & "'" & _
								"         , " & nAlokasi * -1 & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertDB)			  
					connStr.execute(qInsertDB)
					
					maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
					qInsertCR = " insert into employeecb (empcode" & _
								"                         , sourcetype" & _
								"                         , targettype" & _
								"                         , targetcode" & _
								"                         , jobcode" & _
								"                         , amount" & _
								"                         , transdate" & _
								"                         , referenceno" & _
								"                         , fcentry" & _
								"                         , fcedit" & _
								"                         , fcip" & _
								"                         , fcba" & _
								"                         , lastupdate" & _
								"                         , lasttime" & _
								"                         , lineno)" & _
								" values ('" & arrRS(3, i) & "'" & _
								"         , '" & arrRS(7, i) & "'" & _
								"         , '" & arrRS(12, i) & "'" & _
								"         , '" & arrRS(13, i) & "'" & _
								"         , '" & arrRS(11, i) & "'" & _
								"         , " & nAlokasi & "" & _
								"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
								"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Session("CodeName") & "'" & _
								"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								"         , '" & Session("fcBA") & "'" & _
								"         , sysdate" & _
								"         , to_char(sysdate, 'HH24:MI')" & _
								"         , " & maxLineNoEmployeeCB & ")"
					' response.Write(qInsertCR)			  
					connStr.execute(qInsertCR)
				Else
				End If

				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
					Response.Flush
				End If
			End If
		End If
	Next
	Erase arrRS
End If

Response.write("<font face=Verdana size=1 color=Green> UPR 01 - EM, Success! " & " Time: " & time() & " ..<br/><br/></font>")
Response.Flush

Response.Write "No Activity 01" & "<br/>"
Response.Flush

q_noact_01 = " select fcba" & _
             "        , transdate" & _
             "        , allowdedcode" & _
             "        , employeecode" & _
             "        , section" & _
             "        , grade" & _
             "        , amount" & _
             "        , 'EM' sourcetype" & _
             "        , allowded_api.getallded(allowdedcode, fcba) allded" & _
			 "        , new_employee_closing_api.get_InfoParameterNoActivity(allowdedcode, 1, fcba) alokasi_jobcode" & _
			 "        , new_employee_closing_api.get_InfoParameterNoActivity(allowdedcode, 2, fcba) alokasi_tt" & _
			 "        , new_employee_closing_api.get_InfoParameterNoActivity(allowdedcode, 3, fcba) alokasi_tc" & _
             " from employee_allowances_tab" & _
             " where fcba = '" & Session("fcBA") & "'" & _
             "       and new_employee_closing_api.get_uprtype(allowdedcode, section, grade) = '01'" & _
             "       and NEW_EMPLOYEE_CLOSING_API.get_basis(fcba, employeecode, transdate, 'EM') = 0"
			  
arrRS = GetRSArray(q_noact_01)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		If Cdbl(arrRS(6, i)) <> 0 Then
			connstr.BeginTrans
			If Cstr(arrRS(8, i)) = "A" Then
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertDB = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , '" & arrRS(10, i) & "'" & _
							"         , '" & arrRS(11, i) & "'" & _
							"         , '" & arrRS(9, i) & "'" & _
							"         , " & Cdbl(arrRS(6, i)) & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertDB)			  
				connStr.execute(qInsertDB)
				
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertCR = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , 'CB'" & _
							"         , 'BS'" & _
							"         , '" & txtControlJob & "'" & _
							"         , " & Cdbl(arrRS(6, i)) * -1 & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)
			ElseIf Cstr(arrRS(8, i)) = "D" Then
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertDB = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , 'CB'" & _
							"         , 'BS'" & _
							"         , '" & txtControlJob & "'" & _
							"         , " & Cdbl(arrRS(6, i)) * -1 & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertDB)			  
				connStr.execute(qInsertDB)
				
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertCR = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , '" & arrRS(10, i) & "'" & _
							"         , '" & arrRS(11, i) & "'" & _
							"         , '" & arrRS(9, i) & "'" & _
							"         , " & Cdbl(arrRS(6, i)) & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)
			Else
			End If

			If Err.Number = 0 Then
				connstr.CommitTrans
				Response.Flush
			Else
				connstr.RollbackTrans
				Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
				Response.Flush
			End If
		End If
	Next
	Erase arrRS
End If

Response.write("<font face=Verdana size=1 color=Green>No Activity 01, Success! " & " Time: " & time() & " ..<br/><br/></font>")
Response.Flush


Response.Write "No Activity 02" & "<br/>"
Response.Flush

q_noact_02 = " select fcba" & _
             "        , transdate" & _
             "        , allowdedcode" & _
             "        , employeecode" & _
             "        , section" & _
             "        , grade" & _
             "        , amount" & _
             "        , sourcetype" & _
             "        , allded" & _
             "        , new_employee_closing_api.get_detailalokasi(allowdedcode, 1, fcba, section, grade, sourcetype) alokasi_jobcode" & _
             "        , new_employee_closing_api.get_detailalokasi(allowdedcode, 2, fcba, section, grade, sourcetype) alokasi_tt" & _
             "        , new_employee_closing_api.get_detailalokasi(allowdedcode, 3, fcba, section, grade, sourcetype) alokasi_tc" & _
             " from (" & _
             " select fcba" & _
             "        , transdate" & _
             "        , allowdedcode" & _
             "        , employeecode" & _
             "        , section" & _
             "        , grade" & _
             "        , amount" & _
             "        , 'EM' sourcetype" & _
             "        , allowded_api.getallded(allowdedcode, fcba) allded" & _
             " from employee_allowances_tab" & _
             " where fcba = '" & Session("fcBA") & "'" & _
             "       and new_employee_closing_api.get_uprtype(allowdedcode, section, grade) = '02'" & _
             "       and new_employee_closing_api.get_basis(fcba, employeecode, transdate, 'EM') = 0" & _
             " )"
			  
arrRS = GetRSArray(q_noact_02)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		If Cdbl(arrRS(6, i)) <> 0 Then
			connstr.BeginTrans
			If Cstr(arrRS(8, i)) = "A" Then
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertDB = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , '" & arrRS(10, i) & "'" & _
							"         , '" & arrRS(11, i) & "'" & _
							"         , '" & arrRS(9, i) & "'" & _
							"         , " & Cdbl(arrRS(6, i)) & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertDB)			  
				connStr.execute(qInsertDB)
				
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertCR = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , 'CB'" & _
							"         , 'BS'" & _
							"         , '" & txtControlJob & "'" & _
							"         , " & Cdbl(arrRS(6, i)) * -1 & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)
			ElseIf Cstr(arrRS(8, i)) = "D" Then
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertDB = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , 'CB'" & _
							"         , 'BS'" & _
							"         , '" & txtControlJob & "'" & _
							"         , " & Cdbl(arrRS(6, i)) * -1 & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertDB)			  
				connStr.execute(qInsertDB)
				
				maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
				qInsertCR = " insert into employeecb (empcode" & _
							"                         , sourcetype" & _
							"                         , targettype" & _
							"                         , targetcode" & _
							"                         , jobcode" & _
							"                         , amount" & _
							"                         , transdate" & _
							"                         , referenceno" & _
							"                         , fcentry" & _
							"                         , fcedit" & _
							"                         , fcip" & _
							"                         , fcba" & _
							"                         , lastupdate" & _
							"                         , lasttime" & _
							"                         , lineno)" & _
							" values ('" & arrRS(3, i) & "'" & _
							"         , '" & arrRS(7, i) & "'" & _
							"         , '" & arrRS(10, i) & "'" & _
							"         , '" & arrRS(11, i) & "'" & _
							"         , '" & arrRS(9, i) & "'" & _
							"         , " & Cdbl(arrRS(6, i)) & "" & _
							"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
							"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT" & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Session("CodeName") & "'" & _
							"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							"         , '" & Session("fcBA") & "'" & _
							"         , sysdate" & _
							"         , to_char(sysdate, 'HH24:MI')" & _
							"         , " & maxLineNoEmployeeCB & ")"
				' response.Write(qInsertCR)			  
				connStr.execute(qInsertCR)
			Else
			End If

			If Err.Number = 0 Then
				connstr.CommitTrans
				Response.Flush
			Else
				connstr.RollbackTrans
				Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
				Response.Flush
			End If
		End If
	Next
	Erase arrRS
End If

Response.write("<font face=Verdana size=1 color=Green>No Activity 02, Success! " & " Time: " & time() & " ..<br/><br/></font>")
Response.Flush


Response.Write "Normalisasi Rounding" & "<br/>"
Response.Flush

qSelisih = " select fcba" & _
		   "        , transdate" & _
		   "        , allowdedcode" & _
		   "        , allowded_api.getallded(allowdedcode, fcba) allded" & _
		   "        , employeecode" & _
		   "        , amount" & _
		   " from ( select fcba, transdate, allowdedcode, employeecode, sum(amount) amount" & _
		   "        from ( select fcba, last_day(adhocdate) transdate, allowdedcode, empcode employeecode, amount" & _
		   "               from empallded" & _
		   "               where fcba = '" & Session("fcBA") & "'" & _
		   "               and to_number(to_char(adhocdate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
		   "               and to_number(to_char(adhocdate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
		   "               AND allowdedcode NOT IN ('A61', 'A59')" & _
		   "               union all" & _
		   "               select fcba, last_day(transdate) transdate, substr(referenceno, 1, 3) allowdedcode, empcode employeecode, amount" & _
		   "               from employeecb" & _
		   "               where fcba = '" & Session("fcBA") & "'" & _
		   "               and to_number(to_char(transdate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
		   "               and to_number(to_char(transdate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
		   "               and jobcode = '" & txtControlJob & "'" & _
		   "               AND SUBSTR(referenceno, 1, 3) NOT IN ('A61', 'A59')" & _
		   "        )" & _
		   "        group by fcba, transdate, allowdedcode, employeecode" & _
		   " )" & _
		   " where amount <> 0"
		   
arrRS = GetRSArray(qSelisih)

If IsArray(arrRS) Then
	For i = LBound(arrRS, 2) To UBound(arrRS, 2)
		If Cstr(arrRS(3, i)) = "A" Then
			Set objectRS = Server.CreateObject("ADODB.Recordset")
			objectSQL = " select min (lineno) minlineno" & _
						" from employeecb" & _
						" where fcba = '" & Session("fcBA") & "'" & _
						"       and transdate = to_date('" & arrRS(1, i) & "', 'mm/dd/yyyy')" & _
						"       and empcode = '" & arrRS(4, i) & "'" & _
						"       and substr(referenceno, 1, 3) = '" & arrRS(2, i) & "'" & _
						"       and jobcode <> '" & txtControlJob & "'"
			objectRS.Open objectSQL, connstr
			If Not objectRS.EOF Then
				FirstRecord = objectRS("minlineno")
			End If
			objectRS.close
			set objectRS = Nothing
			
			connstr.BeginTrans
			maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
			InsertNewRecord = " insert into employeecb ( empcode" & _
							  "                          , sourcetype" & _
							  "                          , targettype" & _
							  "                          , targetcode" & _
							  "                          , jobcode" & _
							  "                          , mandays" & _
							  "                          , units" & _
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
							  " select empcode" & _
							  "        , sourcetype" & _
							  "        , targettype" & _
							  "        , targetcode" & _
							  "        , jobcode" & _
							  "        , mandays" & _
							  "        , units" & _
							  "        , " & cdbl(arrRS(5, i)) & " amount" & _
							  "        , transdate" & _
							  "        , referenceno||'-RND'" & _
							  "        , fcentry" & _
							  "        , fcedit" & _
							  "        , fcip" & _
							  "        , fcba" & _
							  "        , sysdate lastupdate" & _
							  "        , to_char(sysdate, 'hh24:mi') lasttime" & _
							  "        , " & maxLineNoEmployeeCB & " lineno" & _
							  " from employeecb" & _
							  " where lineno = '" & FirstRecord & "'" & _
							  "       and fcba = '" & Session("fcBA") & "'"
							  
			connstr.execute(InsertNewRecord)
			
			maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
			qInsertCR = " insert into employeecb (empcode" & _
						"                         , sourcetype" & _
						"                         , targettype" & _
						"                         , targetcode" & _
						"                         , jobcode" & _
						"                         , amount" & _
						"                         , transdate" & _
						"                         , referenceno" & _
						"                         , fcentry" & _
						"                         , fcedit" & _
						"                         , fcip" & _
						"                         , fcba" & _
						"                         , lastupdate" & _
						"                         , lasttime" & _
						"                         , lineno)" & _
						" values ('" & arrRS(4, i) & "'" & _
						"         , 'EM'" & _
						"         , 'CB'" & _
						"         , 'BS'" & _
						"         , '" & txtControlJob & "'" & _
						"         , " & cdbl(arrRS(5, i)) * -1 & "" & _
						"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
						"         , '" & arrRS(2, i) & "-CR-NEWCONCEPT-RND" & "'" & _
						"         , '" & Session("CodeName") & "'" & _
						"         , '" & Session("CodeName") & "'" & _
						"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						"         , '" & Session("fcBA") & "'" & _
						"         , sysdate" & _
						"         , to_char(sysdate, 'HH24:MI')" & _
						"         , " & maxLineNoEmployeeCB & ")"
			' response.Write(qInsertCR)			  
			connStr.execute(qInsertCR)

			If Err.Number = 0 Then
				connstr.CommitTrans
			Else
				connstr.RollbackTrans
				Response.write("Error Save ["&InsertNewRecord&"]["&qInsertCR&"] ("&Err.Description&")")
				Response.Flush
			End If
			
		ElseIf Cstr(arrRS(3, i)) = "D" Then
			Set objectRS = Server.CreateObject("ADODB.Recordset")
			objectSQL = " select min (lineno) minlineno" & _
						" from employeecb" & _
						" where fcba = '" & Session("fcBA") & "'" & _
						"       and transdate = to_date('" & arrRS(1, i) & "', 'mm/dd/yyyy')" & _
						"       and empcode = '" & arrRS(4, i) & "'" & _
						"       and substr(referenceno, 1, 3) = '" & arrRS(2, i) & "'" & _
						"       and jobcode <> '" & txtControlJob & "'"
			objectRS.Open objectSQL, connstr
			If Not objectRS.EOF Then
				FirstRecord = objectRS("minlineno")
			End If
			objectRS.close
			set objectRS = Nothing
			
			connstr.BeginTrans
			maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
			InsertNewRecord = " insert into employeecb ( empcode" & _
							  "                          , sourcetype" & _
							  "                          , targettype" & _
							  "                          , targetcode" & _
							  "                          , jobcode" & _
							  "                          , mandays" & _
							  "                          , units" & _
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
							  " select empcode" & _
							  "        , sourcetype" & _
							  "        , targettype" & _
							  "        , targetcode" & _
							  "        , jobcode" & _
							  "        , mandays" & _
							  "        , units" & _
							  "        , " & cdbl(arrRS(5, i)) & " amount" & _
							  "        , transdate" & _
							  "        , referenceno||'-RND'" & _
							  "        , fcentry" & _
							  "        , fcedit" & _
							  "        , fcip" & _
							  "        , fcba" & _
							  "        , sysdate lastupdate" & _
							  "        , to_char(sysdate, 'hh24:mi') lasttime" & _
							  "        , " & maxLineNoEmployeeCB & " lineno" & _
							  " from employeecb" & _
							  " where lineno = '" & FirstRecord & "'" & _
							  "       and fcba = '" & Session("fcBA") & "'"
			
			connstr.execute(InsertNewRecord)
			
			maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
			qInsertCR = " insert into employeecb (empcode" & _
						"                         , sourcetype" & _
						"                         , targettype" & _
						"                         , targetcode" & _
						"                         , jobcode" & _
						"                         , amount" & _
						"                         , transdate" & _
						"                         , referenceno" & _
						"                         , fcentry" & _
						"                         , fcedit" & _
						"                         , fcip" & _
						"                         , fcba" & _
						"                         , lastupdate" & _
						"                         , lasttime" & _
						"                         , lineno)" & _
						" values ('" & arrRS(4, i) & "'" & _
						"         , 'EM'" & _
						"         , 'CB'" & _
						"         , 'BS'" & _
						"         , '" & txtControlJob & "'" & _
						"         , " & cdbl(arrRS(5, i)) * -1 & "" & _
						"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
						"         , '" & arrRS(2, i) & "-DB-NEWCONCEPT-RND" & "'" & _
						"         , '" & Session("CodeName") & "'" & _
						"         , '" & Session("CodeName") & "'" & _
						"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						"         , '" & Session("fcBA") & "'" & _
						"         , sysdate" & _
						"         , to_char(sysdate, 'HH24:MI')" & _
						"         , " & maxLineNoEmployeeCB & ")"
			' response.Write(qInsertCR)			  
			connStr.execute(qInsertCR)

			If Err.Number = 0 Then
				connstr.CommitTrans
			Else
				connstr.RollbackTrans
				Response.write("Error Save ["&InsertNewRecord&"]["&qInsertCR&"] ("&Err.Description&")")
				Response.Flush
			End If
		Else
		End If
	
	 Next
	 Erase arrRS
End If

Response.write("<font face=Verdana size=1 color=Green> Normalisasi Rounding, Success! " & " Time: " & time() & " ..<br/><br/></font>")
Response.Flush
%>