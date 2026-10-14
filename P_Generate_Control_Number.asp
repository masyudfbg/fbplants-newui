<%@Language=VBScript%>
<% 
'Option Explicit 
%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!--#include virtual="/ajax/aspJSON1.17.asp" -->

<% Dim textReportName : textReportName = "Control Number "%>

<%
Function GetRSArray(strSQL)
		'Response.Write(strSQL)
		set objRSArray = connStr.Execute(strSQL)
		If Not objRSArray.EOF Then
			arrRS = objRSArray.GetRows()
			GetRSArray = arrRS
		Else
			'Response.Write("There are currently no records returned.<br/>")
        Exit Function
	End If
		objRSArray.Close()
	Set objRSArray = Nothing
End Function
		
Dim TimeAwal,TimeAkhir,Hitung,Lama

TimeAwal=Time()
Dim command : command = Request("command")

Dim Startingtime, Endingtime
Startingtime = FormatDateTime(CDate(Now))
	

Dim v_tahun : v_tahun = Request("txtYear")
Dim v_Bulan : v_Bulan = Request("txtMonth")

stringQuery = "SELECT" &_
              "    Lu_name ModulName," &_
              "    Accountingyear," &_
              "    AccountingPeriod," &_
              "    Prefix," &_
              "    Suffix," &_
              "    AutoNumber CurrentNumber," &_
              "    Lengthpad" &_
              "  FROM" &_
              "     controlnumber" &_
              "  WHERE" &_
              "     accountingyear = " & v_tahun & " AND" &_
              "     accountingperiod = " & v_Bulan & " AND" &_
              "     fcba = '" & Session("fcBA") & "'" 
	arrRS = GetRSArray(stringQuery)
	If IsArray(arrRS) Then
		'//JSON type author Atwal Arifin
		Set oJSON = New aspJSON
			With oJSON.data
			.Add "queryname", "Accounting Year"                      'Create value
			.Add "response", "selected"
			.Add "time", 0
			.Add "longtime", 0
			.Add "item", oJSON.Collection()
			With oJSON.data("item")
				For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				'RESPONSE.WRITE(arrRS(0, i))
				.Add i, oJSON.Collection()                  'Create object
					With .item(i)
						.Add "modulename", arrRS(0, i)
						.Add "Accountingyear", ""&arrRS(1, i)&""
						.Add "AccountingPeriod", ""&arrRS(2, i)&""
						.Add "Prefix", arrRS(3, i)
						.Add "Suffix", arrRS(4, i)
						.Add "CurrentNumber", ""&arrRS(5, i)&""
						.Add "Lengthpad", ""&arrRS(6, i)&""
					End With				
				Next
				Erase arrRS
			End With
		End With
	else	
		if command <> "select" then
		    Dim CheckYear,CheckPeriod
			CheckingAccountingYear = "select sum(accountingyear) accountingyear, sum(accountingperiod) accountingperiod from(" &_
									 "	select count(accountingyear) accountingyear , 0 accountingperiod from accountingyear where fcba = 'KAB'" &_
									 "	and accountingyear = " & v_tahun &_
									 "	union all" &_
									 "	select 0 accountingyear,count(accountingperiod) accountingperiod from accountingperiod " &_
									 "	where fcba = '" & Session("fcBA") & "'" &_
									 "	and accountingperiod = " & v_Bulan & ")"
			check = GetRSArray(CheckingAccountingYear)
			If IsArray(check) Then
				CheckYear = check(0, 0)
				CheckPeriod = check(1, 0)
				
				if CheckYear <> "0" and  CheckPeriod <> "0" then
					Response.Flush()
					Dim strSQL
					Dim tempYear, tempMonth, strDate, tempDate
					Dim qInsertData, doContinueProcess


					Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl

					Dim rsGetParameterTrue, qGetParameterTrue, tempDetail, strSQLNormalization, txtLastDay, txtNextMonth, q, qLineMax, rsqLineMax, tempMaxLine, strClosingLog, strLenghtime, rsLenghtime

					connstr.BeginTrans

					doContinueProcess = "TRUE"

						' First we go for Report Detail SAP - FICO
						'Response.Write "Start Normalization for Financial Controlller Data..." 

						IF v_tahun <> ""  And  v_Bulan <> "" Then
						   
								tempPrefix = Session("fcBA") & Right(v_tahun,2) &  v_Bulan
								tempSuffix = ""
								tempLengthPad = 3
								strSQL = "begin ControlNumber_Api.Generate_Data ('PURCHASEREQUEST'," & v_tahun & ", " & v_Bulan  & ",'" & _ 
																			tempPrefix & "','" & tempSuffix & "'," & _
																			tempLengthPad & ",'" &  Session("fcBA") & "','" & Session("CodeName") & "','" & _
																			Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "'); end;" 
								response.write strSQL
								connStr.execute(strSQL)			
								 
								tempPrefix = "DOC" & Session("fcBA") & v_tahun &  v_Bulan
								tempSuffix = ""
								tempLengthPad = 3
								strSQL = "begin ControlNumber_Api.Generate_Data ('UPLOADPO'," & v_tahun & ", " & v_Bulan  & ",'" & _ 
																			tempPrefix & "','" & tempSuffix & "'," & _
																			tempLengthPad & ",'" &  Session("fcBA") & "','" & Session("CodeName") & "','" & _
																			Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "'); end;" 
								'response.write strSQL
								connStr.execute(strSQL)	 
								 
								
								' Format Grn No adalah GRNKSBTTMMXXXX
								tempPrefix = "GRN" &  Session("fcBA") &  v_tahun &  v_Bulan
								tempSuffix = ""
								tempLengthPad = 4
								strSQL = "begin ControlNumber_Api.Generate_Data ('RECEIVENOTE'," & v_tahun & ", " & v_Bulan  & ",'" & _ 
																			tempPrefix & "','" & tempSuffix & "'," & _
																			tempLengthPad & ",'" &  Session("fcBA") & "','" & Session("CodeName") & "','" & _
																			Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "'); end;" 
								connStr.execute(strSQL)	

                                ' Format Upload Finger No adalah KTP201601
								tempPrefix = "GRN" &  Session("fcBA") &  v_tahun &  Lpad(v_Bulan,2,'0')
								tempSuffix = ""
								tempLengthPad = 4
								strSQL = "begin ControlNumber_Api.Generate_Data ('UPLOAD_MF'," & v_tahun & ", " & v_Bulan  & ",'" & _ 
																			tempPrefix & "','" & tempSuffix & "'," & _
																			tempLengthPad & ",'" &  Session("fcBA") & "','" & Session("CodeName") & "','" & _
																			Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "'); end;" 
								connStr.execute(strSQL)									
						End if
						'Response.Write "<font face=Verdana size=1 color=blue>" & "Generate Control Number is successfully...." & "<br><br></font>"
					if err.number <> 0 or doContinueProcess = "FALSE" then
							connStr.RollbackTrans
							'response.Clear()
							'response.Write strSQL
							'response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
							Set oJSON = New aspJSON
								With oJSON.data
									.Add "queryname", "Accounting Year" '//Create value
									.Add "response", "error"
									.Add "item", oJSON.Collection()
									With oJSON.data("item")
										.Add 0, oJSON.Collection()	'//Create object
											With .item(0)
												.Add "command", command
												.Add "query", strSQL
												.Add "description", err.description
											End With				
									End With
								End With
					Else
							connStr.CommitTrans
							
							TimeAkhir=Time()
							Lama =TimeAkhir-TimeAwal
							'Response.Write "<br>" & "Finish ... " & time() & "<br><br>"
							'Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."
							
						arrRS = GetRSArray(stringQuery)
							If IsArray(arrRS) Then
							'//JSON type author Atwal Arifin
								Set oJSON = New aspJSON
								With oJSON.data
									.Add "queryname", "Accounting Year" '//Create value
									.Add "response", "successfully"
									.Add "time", ""&TimeAkhir&""
									.Add "longtime", ""&Lama&""
									.Add "item", oJSON.Collection()
									With oJSON.data("item")
										For i = LBound(arrRS, 2) To UBound(arrRS, 2)
										'RESPONSE.WRITE(arrRS(0, i))
										.Add i, oJSON.Collection()	'//Create object
											With .item(i)
												.Add "modulename", arrRS(0, i)
												.Add "Accountingyear", ""&arrRS(1, i)&""
												.Add "AccountingPeriod", ""&arrRS(2, i)&""
												.Add "Prefix", arrRS(3, i)
												.Add "Suffix", arrRS(4, i)
												.Add "CurrentNumber", ""&arrRS(5, i)&""
												.Add "Lengthpad", ""&arrRS(6, i)&""
											End With				
										Next
										Erase arrRS
									End With
								End With
							End If
					End if
				else

				Set oJSON = New aspJSON
					With oJSON.data
						.Add "queryname", "Checking" '//Create value
						.Add "response", "error"
						.Add "item", oJSON.Collection()
						With oJSON.data("item")
							.Add 0, oJSON.Collection()	'//Create object
							With .item(0)
								.Add "command", command
								.Add "query", "Account year" 
								.Add "description", ""&CheckYear&""
							End With
							.Add 1, oJSON.Collection()	'//Create object
							With .item(1)
								.Add "command", command
								.Add "query", "Account Period" 
								.Add "description", ""&CheckPeriod&""
							End With							
						End With
					End With
				end if
			else
			
			
			end if
			
		else
			Set oJSON = New aspJSON
						With oJSON.data
							.Add "queryname", "Accounting Year" '//Create value
							.Add "response", "error"
							.Add "item", oJSON.Collection()
							With oJSON.data("item")
								.Add 0, oJSON.Collection()	'//Create object
									With .item(0)
										.Add "command", command
										.Add "query", "no-Query"
										.Add "description", "no-Description"
									End With				
							End With
						End With
		End If
	End If
		
Response.Write oJSON.JSONoutput()  	
%>
