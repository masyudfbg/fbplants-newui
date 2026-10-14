
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


dim company,idresta,fcba,r1,fctype,vehiclegroup
r1      = Request("r1")
company = Request("fccompanycode")
idresta = Request("idresta")
fcba    = Request("fcba")
txtcari = Request("keyword")
param   = Request("param")

Set List = Server.CreateObject("ADODB.Recordset")

If trim(fcba) = "-" Then
    filters = " NVL(fcba, fcba) IN ( " & _
                " SELECT fccode " & _
                " FROM businessunit " & _
                " WHERE fccompanycode = '" & company & "' " & _
                " AND isactive = 'TRUE' AND fctype='E' " & _
                " ) "
Else
    filters = " fcba ='" & fcba & "' "
End If

sqlview = " SELECT '-' id, '--All--' text FROM Dual " & _
          " UNION " & _ 
          " SELECT nurserycode id, nurserycode text " & _
          " FROM nursery WHERE "& filters &""

'  response.write sqlview

List.Open sqlview, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.asp가 필요함
Set rs = connStr.Execute(sqlview)

Dim List_jsonString
List_jsonString = caljson.toJSON("results", rs, false)

Dim jsonObj, outputObj
set jsonObj = new JSONobject
set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
oJSON.loadJSON(List_jsonString)

List.Close
set List = Nothing

Response.write List_jsonString
%>