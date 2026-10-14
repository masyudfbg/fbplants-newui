
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
fcba = session("fcba") 
company = session("COMPANYCODE")
viewdata = request("viewdata")

Set List = Server.CreateObject("ADODB.Recordset")

if viewdata = "fcba" then
    sqlview =  "select FCCODE as id, FCCODE as text from  BUSINESSUNIT  where FCCOMPANYCODE= '" & company & "' " & _
    " and isactive='TRUE' and fccode like UPPER('%" & txtcari & "%') order by FCCODE asc"
else

sqlview =  "select FCCODE as id, FCCODE as text from  PARAMETER  where PARCODE= 'E0102'  and fcba='" & fcba & "'  " & _
"  and fccode like UPPER('%" & txtcari & "%') order by FCCODE asc"
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

Response.write List_jsonString
%>