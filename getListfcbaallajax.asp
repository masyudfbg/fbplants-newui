
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

if idresta =  1 then 
    sqlview = " SELECT '-' id,'--All--' text FROM Dual " & _
            " UNION " & _ 
            " SELECT fccode id,fccode text " & _
            " FROM businessunit a inner join user_fcba_tab b on a.fccode=b.fcba " & _
            " WHERE fcusercode='" & Session("CodeName") & "' AND isactive='TRUE' AND fccompanycode='" & company & "'  order by text"
elseif idresta =  2 then 
    sqlview = " SELECT '-' id,'--All--' text FROM Dual " & _
            " UNION " & _ 
            " SELECT fccode id,fccode text " & _
            " FROM businessunit a inner join user_fcba_tab b on a.fccode=b.fcba " & _
            " WHERE fcusercode='" & Session("CodeName") & "' AND isactive='TRUE' AND fccompanycode='" & company & "'  order by text"
elseif idresta =  3 then 
    sqlview = " SELECT 'all' id,'--All--' text FROM Dual UNION select distinct sectionname as id, sectionname_api.get_fcname(sectionname) text from employee where fcba= '" & fcba & "' order by text"
          '  Response.Write sqlview
elseif idresta =  4 then 
    sqlview = " SELECT 'all' id,'--All--' text FROM Dual UNION select distinct fccode as id, fccode text from sectionname where fccode LIKE UPPER('%" & txtcari & "%') order by text"
          '  Response.Write sqlview
elseif idresta =  5 then 
    sqlview =  " SELECT '-' id,'--All--' text FROM Dual " & _
            " UNION " & _ 
            " SELECT fccode id,fccode text " & _
            " FROM businessunit a inner join user_fcba_tab b on a.fccode=b.fcba " & _
            " WHERE fcusercode='" & Session("CodeName") & "' AND isactive='TRUE' AND fccompanycode like '" & company & "%'  order by text"
elseif idresta =  6 then 
    if r1 = "1" then 
       fctype  = "M"
    else 
       fctype  = "E"
    end if

    sqlview =   " SELECT '-' id,'--All--' text FROM Dual " & _
            " UNION " & _ 
            " SELECT fccode id,fccode text " & _
            " FROM businessunit a inner join user_fcba_tab b on a.fccode=b.fcba " & _
            " WHERE fcusercode='" & Session("CodeName") & "' AND isactive='TRUE' AND fccompanycode like '" & company & "%' AND a.fctype='" & fctype & "'  order by text"
elseif idresta = 7 then
    sqlview = " SELECT '-' id,'--All--' text FROM Dual " & _
              " UNION ALL " & _
              " SELECT COMPANYCODE as id, COMPANYCODE as text " & _
              " FROM company " & _
              " WHERE REGION = '"&company&"' "
elseif idresta = 8 then
    sqlview = " SELECT '-' id,'--All--' text FROM Dual " & _
              " UNION ALL " & _
              " SELECT fccode id,fccode text " & _
              " FROM businessunit a inner join user_fcba_tab b on a.fccode=b.fcba " & _
              " WHERE a.fccompanycode like '" & company & "%' AND a.isactive='TRUE' AND b.fcusercode='" & Session("CodeName") & "' " & _ 
              " ORDER BY text"
elseif idresta = 9 then
    sqlview = " SELECT '-' id,'--All--' text FROM Dual " & _
              " UNION ALL " & _
              " SELECT fccode id,fccode text " & _
              " FROM businessunit a inner join user_fcba_tab b on a.fccode=b.fcba " & _
              " WHERE a.fccompanycode like '" & company & "%' AND a.isactive='TRUE' AND b.fcusercode='" & Session("CodeName") & "' " & _ 
              " ORDER BY text"
elseif idresta = 10 then
    sqlview = " SELECT '-' id,'--All--' text FROM Dual " & _
              " UNION ALL " & _
              " SELECT fccode id,fcname text " & _
              " FROM vehicle where vehiclegroupcode = '"& param &"'" & _
              " group by fccode, fcname ORDER BY text ASC"
elseif idresta = 11 then
    sqlview = " SELECT '-' id,'--All--' text FROM Dual " & _
              " UNION ALL " & _
              " SELECT fccode id, fcname text FROM vehicleactivitycategory where fcba = '"& param &"' ORDER BY text ASC "
elseif idresta = 12 then
    sqlview = " SELECT '-' id,'--All--' text FROM Dual " & _
              " UNION ALL " & _
              " SELECT DISTINCT fccode id, fccode || ' - ' || fcname text FROM GANG where fcba = '"& param &"' ORDER BY text ASC "
else
    sqlview =   " SELECT '-' id,'--All--' text FROM Dual " & _
            " UNION " & _ 
            " SELECT fccode id,fccode text " & _
            " FROM businessunit a inner join user_fcba_tab b on a.fccode=b.fcba " & _
            " WHERE fcusercode='" & Session("CodeName") & "' AND isactive='TRUE' AND fccompanycode like '" & company & "%'  order by text"
end if

 'response.write sqlview



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