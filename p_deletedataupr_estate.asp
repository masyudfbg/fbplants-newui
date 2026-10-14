<%
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
ExecuteSQL = " delete from allowances_loc_tab" & _
			 " where fcba = '" & Session("fcBA") & "'"

connstr.execute(ExecuteSQL)

If Err.Number = 0 Then
	connstr.CommitTrans
Else
	connstr.RollbackTrans
	Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
	Response.Flush
End If
%>