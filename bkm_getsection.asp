
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


dim txtcari,fcba,txtsupplier 
txtcari = Request("keyword")
section = Request("section")
rkbno = Request("rkbno")



Set List = Server.CreateObject("ADODB.Recordset")

if section = 1 then 
    sqlview =  "select fccode as id, fccode as text from sectionname " & _
    "  where fccode like UPPER('%" & txtcari & "%') order by fccode asc"
end if 

if section = 2 then 
    sqlview =  "select fccode as id, fccode as text from sectionname " & _
    "  where fccode like UPPER('%" & txtcari & "%') order by fccode asc"
end if

if section = 9 then 
    sqlview =  "select sectioncode as id, sectioncode as text from rkbhdr " & _
    "  where fcba='" & Session("fcba") & "' and rkbno='" & rkbno &"' and  sectioncode like UPPER('%" & txtcari & "%') order by sectioncode asc"
end if

if section = 10 then 
    sqlview =  "select fccode as id, fccode || ' - ' || fcname as text from vehicle " & _
    "  where fcba='" & Session("fcba") & "'  and  fccode || ' - ' || fcname like UPPER('%" & txtcari & "%') and activation='Y' order by fccode || ' - ' || fcname asc"
end if

List.Open sqlview, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.asp가 필요함
Set rs = connStr.Execute(sqlview)

Dim List_jsonString
List_jsonString = caljson.toJSON("results",rs, false)

Dim jsonObj, outputObj
set jsonObj = new JSONobject
set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
oJSON.loadJSON(List_jsonString)

List.Close
set List = Nothing

Response.write List_jsonString
%>