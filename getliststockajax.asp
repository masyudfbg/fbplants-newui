
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


dim txtcari,fcba,id 
id = Request("id")
txtcari = Request("keyword")
txtgroup    = Request("group")
txtsubgroup = Request("subgroup")
fcba = Request("fcba")


Set List = Server.CreateObject("ADODB.Recordset")

' All Stock
if id =  1 then 
    sqlview = " SELECT fccode id,fcname text " & _
			  " FROM bankstock " & _
			  " WHERE stockgroupcode like '" & txtgroup & "%' AND stocksubgroupcode like '" & txtsubgroup & "%' order by fccode"
elseif id =  2 then 
    sqlview = " SELECT DISTINCT a.itemcode id,a.itemcode||' - '||c.fcname text" & _ 
			  " FROM purchaseorderdetails a " & _ 
			  " INNER JOIN bankstock c ON (a.itemcode=c.fccode) " & _ 
			  " WHERE c.stockgroupcode like '" & txtgroup & "%' AND c.stocksubgroupcode like '" & txtsubgroup & "%'" & _ 
			  " ORDER BY a.itemcode " 
elseif id =  3 then 
    sqlview = " SELECT '-' id,'--All--' text FROM DUAL UNION SELECT fccode id,fccode||' - '||fcname text" & _ 
			  " FROM stockgroup  " & _ 
			  " WHERE fccode like '" & txtgroup & "%' AND fcba='" & fcba & "'" & _ 
			  " ORDER BY id " 

elseif id =  4 then 
    sqlview = " SELECT '' id,'--All--' text FROM DUAL UNION SELECT fccode id,fccode||' - '||fcname text" & _ 
			  " FROM stocksubgroup  " & _ 
			  " WHERE fccode like '" & txtgroup & "%' AND fcba='" & fcba & "'" & _ 
			  " ORDER BY id " 
	
else
    sqlview = " SELECT '-'  id,'None' text " & _
			  " FROM dual " 
end if

'Response.write sqlview

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

Response.write List_jsonString
%>