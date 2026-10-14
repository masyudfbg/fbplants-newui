
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


dim txtcari,fcba,gangcode,upahedit,bkmno
fcba = session("fcba")
txtcari = Request("keyword")
upahedit    = Request("jmlitem")
bkmno       = Request("bkmno")






if upahedit = "1" then 
    sqlview = "Select B.EMPLOYEECODE,E.FCNAME, B.LINENOKEY, B.REMARKS, " & _
    " B.TBSA, B.TBSBA, B.TBSTP, B.TBSBK, B.TBSO, " & _
    " B.TBSAB, B.TBSE, B.ANCAKBT, B.ANCAKPS, B.ANCAKSP, " & _
    " B.ANCAKBRD, B.ANCAKOP, B.ANCAKBM, B.ANCAKBDT " & _
    " From   BKMDENDA B " & _
    " LEFT JOIN EMPLOYEE E ON (B.EMPLOYEECODE=E.FCCODE AND B.FCBA=E.FCBA) " & _
    " WHERE B.BKMNO='" & bkmno & "' order by e.fcname asc"

end if
Set List = Server.CreateObject("ADODB.Recordset")
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