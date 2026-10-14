
<%
Response.CharSet="euc-kr"
Session.codepage="949"
Response.codepage="949"
Response.ContentType="text/html;charset=euc-kr"
%>
<%
Response.LCID = 1046 
%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->

<!--#includes file="include/JSON.asp" --> 
<!--#includes file="include/json.objectClass.asp"-->
<!--#includes file="include/aspJSON.asp" -->

<%


dim txtcari,fcba,tahun,bulan,msg,msgerror,company,filtercompany1,fcbaloop,ScriptName
Dim TimeAwal,TimeAkhir,Hitung,Lama, sqlpt
    ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
    company = Request("company")
    fcba  = Request("fcba")
	tahun = Request("tahun")
    bulan = Request("bulan")
	msgerror = ""
	msg = ""
	
	
	
	if trim(fcba) = "-" then 
		sqlpt = "select fccode from BUSINESSUNIT WHERE isactive='TRUE' AND fccompanycode = '" & company & "' AND fccode like '%' ORDER BY nourut"
	else 
		sqlpt = "select fccode from BUSINESSUNIT WHERE isactive='TRUE' AND fccompanycode = '" & company & "' AND fccode ='" & fcba & "' ORDER BY nourut"
	end if

	set rspt = connStr.Execute(sqlpt)
	if not rspt.eof then
		While Not rspt.EOF
			fcbaloop = rspt("fccode") 
			TimeAwal=Time()

			qCheckStatus = "select EMPALLDEDSTATUS_API.GetStatusClosing_acc (" & tahun & "," & bulan & ",'" & fcbaloop & "') closingStatus from dual "
			Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
			rsCheckStatus.CursorType = 3
			rsCheckStatus.CursorLocation = 3
			rsCheckStatus.Open qCheckStatus, connstr
			if not rsCheckStatus.eof then
			txtClosingStatus = rsCheckStatus("closingStatus")
			else
			msgerror =  "EmpalldedStatus are not defined "
			msg = msg &  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
			txtClosingStatus="FALSE"
			end if
			rsCheckStatus.close
			set rsCheckStatus = Nothing
			msg = msg &   "<br>" 
			msg = msg &   "<font face=Verdana size=1 color=green>" & " Begin Proses Closing " & fcbaloop & " Accounting Period: " & bulan & "/" & tahun & " Time: " & TimeAwal & "</font><br>" 

		
			IF txtClosingStatus = "FALSE" THEN  
				qProcessPayroll = "BEGIN " & _ 
								"fac_closing_accounting_api.ProsesClosingMonth (" & tahun & "," & bulan & ",'" & fcbaloop & "','" & Session("CodeName") & "'); " & _ 
								"END;"		
				'On Error Resume Next
				' response.write qProcessPayroll
				connstr.execute(qProcessPayroll)
				'ExecuteAndLog ScriptName, qProcessPayroll
				
				TimeAkhir=Time()
				Lama =TimeAkhir-TimeAwal

				If Err.Number = 0 Then
				msg = msg &   "<font face=Verdana size=1 color=green>" &  " End Proses Closing " & fcbaloop & " Success,  Time: " & TimeAkhir & " Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds." & "</font><br>" 
				Else
					msgerror =  "<font face=Verdana size=1 color=red>" & " Proses Error !!!" & "<br>" & Err.Description & "</font><br>" 
					msg = msg & "<font face=Verdana size=1 color=red>" & " Proses Error !!!" & "<br>" & Err.Description & "</font><br>" 
				End If
			'On Error GoTo 0
			ELSE
			msgerror =  "<font face=Verdana size=1 color=red>" & "BUSINESSUNIT : " & fcbaloop & " Accounting Period " & bulan & "/" & tahun & " LOCKED !!! " & "</font><br>" 
			msg = msg &  "<font face=Verdana size=1 color=red>" & "BUSINESSUNIT : " & fcbaloop & " Accounting Period " & bulan & "/" & tahun & "LOCKED !!! " & "</font><br>" 
			msg = msg &  " End Proses Closing " & fcbaloop & " ,Time: " & TimeAkhir & " Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds." & "<br>" 
			END IF	

		rspt.MoveNext
		Wend
		rspt.close()
		set rspt = Nothing
	end if


	'Set List = Server.CreateObject("ADODB.Recordset")
    'sqlview = "select  '" & msg & "' as msg from dual"
	
	
    'List.Open sqlview, connStr, 1
    'ListCount = List.RecordCount 

    'Set caljson = new JSON  'JSON.asp가 필요함
    'Set rs = connStr.Execute(sqlview)

    'Dim List_jsonStringgl
    'List_jsonStringgl = caljson.toJSON("results",rs, false)

    'Dim jsonObjgl, outputObjgl
    'set jsonObjgl = new JSONobject
    'set outputObjgl = jsonObjgl.parse(List_jsonStringgl)  'json.objectClass.asp가 필요함(파싱하려면)


    'Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    'oJSON.loadJSON(List_jsonStringgl)

Response.write msg
	
    'if msgerror = "" then 
    'Response.write List_jsonStringgl
    'Response.Flush
	'else
	'Response.write msgerror
	'end if
	
	
    
     
%>