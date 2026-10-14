
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


dim txtcari,fcba,tahun,bulan,msg,ScriptName
Dim TimeAwal,TimeAkhir,Hitung,Lama
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
	TimeAwal=Time()
    fcba  = session("fcba")
	tahun = Request("tahun")
    bulan = Request("bulan")
	txtTable = "Process_closing_Opr"
	
	
	'-- Check for Empallded PreProcess
qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreProcess (" & tahun & "," & bulan & ",'" & fcba & "') closingStatus from dual "
Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
rsCheckStatus.CursorType = 3
rsCheckStatus.CursorLocation = 3
rsCheckStatus.Open qCheckStatus, connstr
if not rsCheckStatus.eof then
   txtClosingStatus = rsCheckStatus("closingStatus")
else
   msg =  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
   txtClosingStatus="FALSE"
end if
rsCheckStatus.close
set rsCheckStatus = Nothing
	
IF txtClosingStatus = "TRUE" THEN  
	
	sqlexec = "BEGIN " & _ 
              "bkmcalculate_payroll_api.proses('" & bulan & "','" & tahun & "','" & fcba & "','" & Session("CodeName") & "');" & _
			  "END; " 
    On Error Resume Next
    'connStr.Execute(sqlexec)
	ExecuteAndLog ScriptName, sqlexec
	
	If Err.Number <> 0 Then
		msg = msg & "<font face=Verdana size=1 color=red>" & " Proses Payroll Error !!!" & "<br>" & Err.Description & "</font><br>" 
	End If
	On Error GoTo 0
  
    
	
	qProcessPayroll = "BEGIN " & _ 
	                  "EMPALLDEDSTATUS_API.FinishProcessPayroll (" & tahun & "," & bulan & ",'" & Session("fcBA") & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "'); " & _ 
					  "END;"		
    On Error Resume Next
    'connstr.execute(qProcessPayroll)
	ExecuteAndLog ScriptName, qProcessPayroll
    
	If Err.Number <> 0 Then
		msg = msg & "<font face=Verdana size=1 color=red>" & " Proses Status Error !!!" & "<br>" & Err.Description & "</font><br>" 
	End If
	On Error GoTo 0
  

TimeAkhir=Time()
Lama =TimeAkhir-TimeAwal


msg =  "Proses Finish ... " & time() & "<br>" & _
       " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."

   
   
ELSE
    msg = "<font face=Verdana size=1 color=red>" & "System are UNABLE to continue calculate the Payroll , Status Payroll has been Locked !!!" & "</font><br>" 
END IF	
Set List = Server.CreateObject("ADODB.Recordset")

 sqlview = "select  '" & msg & "' as msg from dual"
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonStringgl
    List_jsonStringgl = caljson.toJSON("results",rs, false)

    Dim jsonObjgl, outputObjgl
    set jsonObjgl = new JSONobject
    set outputObjgl = jsonObjgl.parse(List_jsonStringgl)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonStringgl)

    Response.write List_jsonStringgl

    
     
%>