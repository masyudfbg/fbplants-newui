
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


dim txtcari,fcba,tgl,msg
    fcba  = session("fcba")
	tgl   = Request("tglbkm")
 	
	txtClosingStatus = StatusAccClosingTgl(tgl,fcba)
	msg =  txtClosingStatus

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