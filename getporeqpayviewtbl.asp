
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


dim txtcari,pocode 
txtcari = Request("keyword")
pocode = Request("pocode")


Set List = Server.CreateObject("ADODB.Recordset")



    sqlview =  "select d.pocode,d.itemcode,s.fcname,d.quantity,d.amount " & _
    " from purchaseorderdetails d,stock s " & _
    " where d.itemcode=s.fccode and d.pocode='" & pocode & "' " & _
    " and d.fcba=s.fcba"
   
    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString
    List_jsonString = caljson.toJSON("dataview1",rs, false)

    Dim jsonObj, outputObj
    set jsonObj = new JSONobject
    set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)

    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString)

    Response.write List_jsonString


%>