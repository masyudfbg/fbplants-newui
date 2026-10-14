
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


dim txtcari,fcba,id, rkbno 
txtcari = Request("tglbkm")
id      = Request("id")
doc_cls = Request("doc")
fcba    = Request("fcba")
rkbno    = Request("rkbno")


Set List = Server.CreateObject("ADODB.Recordset")

if id = 1 then 
    sqlview =  	"	select rkbno as id, rkbno as text from rkbhdr " & _
				"	where fcba = '" & fcba & "' AND rowstate='Approved' AND document_classification='" & doc_cls & "' " & _ 
				"	AND period_year*100+period_month  IN (SELECT  YEAR*100+month FROM empalldedstatus " & _
				"	WHERE fcba= '" & fcba & "' AND To_Date('" & txtcari & "','MM/DD/YYYY') BETWEEN start_date AND end_date) " & _ 
				"   UNION ALL " & _
				" 	select '-' as id,'--Pilih RKB--' as text from dual " & _
				" 	order by id asc" 
	' response.write sqlView
end if 

if section = 2 then 
    sqlview =  "select fccode as id, fccode as text from sectionname " & _
    "  where fccode like UPPER('%" & txtcari & "%') order by fccode asc"
end if

' RKB EDIT
if id = 3 then 
    sqlview =  	"	select rkbno as id, rkbno as text from rkbhdr " & _
				"	where rowstate='Approved' AND document_classification='" & doc_cls & "' " & _ 
				"	AND period_year*100+period_month  IN (SELECT  YEAR*100+month FROM empalldedstatus " & _
				"	WHERE rkbno='" & rkbno & "' AND To_Date('" & txtcari & "','MM/DD/YYYY') BETWEEN start_date AND end_date) " & _ 
				"   UNION ALL " & _
				" 	select '-' as id,'--Pilih RKB--' as text from dual " & _
				" 	order by id asc" 
end if 

if id = 4 then 
    sqlview =  	"	select rkbno as id, rkbno as text from rkbhdr " & _
				"	where fcba = '" & fcba & "' AND rowstate='Approved' AND document_classification in ('501','502','503') " & _ 
				"	AND period_year*100+period_month  IN (SELECT  YEAR*100+month FROM empalldedstatus " & _
				"	WHERE fcba= '" & fcba & "' AND To_Date('" & txtcari & "','MM/DD/YYYY') BETWEEN start_date AND end_date) " & _ 
				"   UNION ALL " & _
				" 	select '-' as id,'--Pilih RKB--' as text from dual " & _
				" 	order by id asc" 
end if 
'Response.Write sqlview

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