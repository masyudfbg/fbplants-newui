<% Response.Buffer = False %>
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

fcba = session("fcba")
wipdate = Request("wipdate")

Dim wipMonth, wipYear
wipMonth = CInt(Month(CDate(wipdate)))
wipYear  = CInt(Year(CDate(wipdate)))

Set List = Server.CreateObject("ADODB.Recordset")


sqlview =  "Select START_DATE, END_DATE, END_DATE + INTERVAL '3' DAY AS END_DATE_PLUS3 FROM EMPALLDEDSTATUS WHERE FCBA = '"& fcba &"' and year = "&wipYear&" and month = "&wipMonth&""

' response.write sqlview

List.Open sqlview, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.aspê°€ í•„ìš”í•¨
Set rs = connStr.Execute(sqlview)

Dim List_jsonStringgllhp
List_jsonStringgllhp = caljson.toJSON("results",rs, false)

Dim jsonObjgllhp, outputObjgllhp
set jsonObjgllhp = new JSONobject
set outputObjgllhp = jsonObjgllhp.parse(List_jsonStringgllhp)  'json.objectClass.aspê°€ í•„ìš”í•¨(íŒŒì‹±í•˜ë ¤ë©´)


Set oJSON = New aspJSON     'aspJSON.asp í•„ìš”í•¨(forë¬¸)
oJSON.loadJSON(List_jsonStringgllhp)

Response.write List_jsonStringgllhp

%>