
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
    fcba  = Request("fcba")
	tahun = Request("tahun")
    bulan = Request("bulan")
	company = Request("company")
	prd = Request("prd")
	

    if prd = "12" then
	qProcessPayroll = "BEGIN " & _ 
	                  "fac_closing_accounting_api.ProsesClosingYear12 (" & tahun & ",'" & company & "','" & fcba & "','" & Session("CodeName") & "'); " & _ 
					  "END;"
    ELSE
	qProcessPayroll = "BEGIN " & _ 
	                  "fac_closing_accounting_api.ProsesClosingYear (" & tahun & ",'" & company & "','" & fcba & "','" & Session("CodeName") & "'); " & _ 
					  "END;"
    
	end if	
    'connstr.execute(qProcessPayroll)
	ExecuteAndLog ScriptName, qProcessPayroll
    

TimeAkhir=Time()
Lama =TimeAkhir-TimeAwal


msg =  " Proses Finish ... " & time() & "<br>" & _
       " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."

   
   
'ELSE
'    msg = "<font face=Verdana size=1 color=red>" & "System are UNABLE to continue calculate the Payroll Please Un-Close Financial Process"
'END IF	

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