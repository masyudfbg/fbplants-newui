<%
		tglnow = Session("tglnow")
		
		varKodeAllowdedAttendanceEmpallded = "A99"
		response.Write "Calculate Allowances From Activities..." & "<br>"
		response.Flush()
		' Proses Memasukkan data premi BKM ke dalam List Data Payroll
		IsIncludeProcessPayroll = "FALSE"
		qSelection = 	" SELECT includedpayrollprocessing" & _
						" FROM allowded" & _
						" WHERE fccode = '" & varKodeAllowdedAttendanceEmpallded & "'" & _
						"       AND fcba = '" & Session("FCBA") & "'"
		
		Set rsGeneralSelection = Server.CreateObject("ADODB.Recordset")
		rsGeneralSelection.Open qSelection, connstr
		
		if not rsGeneralSelection.EOF then
			IsIncludeProcessPayroll = rsGeneralSelection("includedpayrollprocessing")
		else
			IsIncludeProcessPayroll = "FALSE"
		end if
		rsGeneralSelection.close: Set rsGeneralSelection = Nothing
		
		if IsIncludeProcessPayroll = "TRUE" then
			connstr.BeginTrans
			ExecuteSQL = " delete" & _
                         " from empallded" & _
                         " where fcba = '" & Session("fcBA") & "'" & _
						 " adhocdate= TO_DATE('" & tglnow & "','mm/dd/yyyy')" & _
						 " AND allowdedcode = '"& varKodeAllowdedAttendanceEmpallded &"'"
			connstr.execute(ExecuteSQL)
			If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
			End If
			
			qSelection = 	" SELECT ag.employeecode empcode" & _
							"        , ag.fddate adhocdate" & _
							"        , ag.jobcode" & _
							"        , SUM(ae.amount) amount" & _
							" FROM attendance_empallded ae" & _
							" LEFT JOIN attendance_gad ag ON (ag.fcba = ae.fcba AND ag.linenokey = ae.linenokey)" & _
							" WHERE ae.fcba = '" & Session("FCBA") & "'" & _
							" AND ag.FDDATE= TO_DATE('" & tglnow & "','mm/dd/yyyy')" & _
							" GROUP BY ag.employeecode" & _
							"          , ag.fddate, ag.jobcode"
			
			Set rsGeneralSelection = Server.CreateObject("ADODB.Recordset")
			rsGeneralSelection.Open qSelection, connstr
			
			If Not rsGeneralSelection.EOF then
				Do While Not rsGeneralSelection.EOF
					If cdbl(to_nol(rsGeneralSelection("amount"))) <> 0 Then
					
						maxLineNo	=	maxLineNo + 1				
						qAllowance	=	" INSERT INTO empallded ( empcode" & _
										"                         , allowdedcode" & _
										"                         , adhocdate" & _
										"                         , amount" & _
										"                         , fixedallded" & _
										"                         , remarks" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" VALUES ( '" & rsGeneralSelection("empcode") & "'" & _
										"          , '" & varKodeAllowdedAttendanceEmpallded & "'" & _
										"          , TO_DATE('" & rsGeneralSelection("adhocdate") & "','mm/dd/yyyy')" & _
										"          , '" & cdbl(to_nol(rsGeneralSelection("amount"))) & "'" & _
										"          , ''" & _
										"          , '" & rsGeneralSelection("jobcode") & "'" & _
										"          , '" & Session("CodeName") & "'" & _
										"          , '" & Session("CodeName") & "'" & _
										"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"          , '" & Session("fcBA") & "'" & _
										"          , SYSDATE" & _
										"          , TO_CHAR(SYSDATE,'hh24:mi')" & _
										"          , '" & maxLineNo & "')"
						connstr.execute(qAllowance)

					End If
					
					rsGeneralSelection.MoveNext
				Loop
			End If
			rsGeneralSelection.close: Set rsGeneralSelection = Nothing
			
		End If
%>