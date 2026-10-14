
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
fcba = Request("fcBA")

Set List = Server.CreateObject("ADODB.Recordset")


sqlview =  "select h.pocode as id,h.pocode as text " & _
" from purchaseorder h,purchaseorderdetails d " & _
" where h.POCODE=d.POCODE and d.quantity-d.qtypayreq>0 and h.fcba='" & fcba & "' and to_char(h.podate,'yyyy')>2024 " & _
" and h.pocode like UPPER('%" & txtcari & "%') group by h.pocode order by h.pocode asc"


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