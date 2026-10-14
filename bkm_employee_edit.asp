
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
gangcode = Request("gangcode")
upahedit = Request("upahedit")
bkmno   = Request("bkmno")
tglbkm  = Request("tglbkm")
emp     = Request("emp")
fccode     = Request("fccode")
empcode     = Request("empcode")



Set List = Server.CreateObject("ADODB.Recordset")
 
if emp = "1" then 
    sqlview = "select e.fccode id, e.fcname || '-' ||  e.fccode as text, e.fcba, e.PRESENTGRADE,timein timeinx,timeout timeoutx,NVL(to_char(a.timein,'HH24:MI:SS'),'-')timein,NVL(to_char(a.timeout,'HH24:MI:SS'),'-')timeout " & _
    " from employee e left join USERFINGER.ATTENDANCE_TAB a ON a.EMPNO=e.FCCODE AND a.FCBA=e.FCBA AND a.TDATE=TO_DATE('" & tglbkm & "','MM/DD/YYYY') " & _
    "  where e.fcba='" & fcba & "' and e.fccode='" & fccode & "' and e.dateterminate is null order by e.fcname asc"
elseif emp = "3" then 
     sqlview = "select  e.fccode id, e.fcname || '-' ||  e.fccode as text " & _
    " from employee e  " & _
    " inner join bkmupahdtl b ON (b.EMPLOYEECODE =e.FCCODE AND b.FCBA=e.FCBA)  " & _
    " inner join bkmhdr h on (h.bkmno=b.bkmno and h.fcba=b.fcba) " & _
    " WHERE h.bkmdate=TO_DATE('" & tglbkm & "','MM/DD/YYYY') and h.DOCUMENT_CLASSIFICATION='502' " & _
    " AND h.fcba='" & fcba & "' and h.bkmno='" & bkmno & "' " & _
    " GROUP BY e.fccode , e.fcname || '-' ||  e.fccode"
   '' Response.write sqlview

elseif emp = "4" then 
     sqlview = "select  e.fccode id, e.fcname || '-' ||  e.fccode as text " & _
    " from employee e  " & _
    " inner join bkmupahdtl b ON (b.EMPLOYEECODE =e.FCCODE AND b.FCBA=e.FCBA)  " & _
    " inner join bkmhdr h on (h.bkmno=b.bkmno and h.fcba=b.fcba) " & _
    " WHERE h.bkmdate=TO_DATE('" & tglbkm & "','MM/DD/YYYY') and h.DOCUMENT_CLASSIFICATION='503' " & _
    " AND h.fcba='" & fcba & "' and h.bkmno='" & bkmno & "' " & _
    " order by e.fcname asc"
'Response.Write sqlview
elseif emp = "5" then 
     sqlview = "select  b.linenokey as linedenda " & _
    " from employee e  " & _
    " inner join bkmupahdtl b ON (b.EMPLOYEECODE =e.FCCODE AND b.FCBA=e.FCBA)  " & _
    " inner join bkmhdr h on (h.bkmno=b.bkmno and h.fcba=b.fcba) " & _
    " WHERE h.fcba='" & fcba & "' and h.bkmno='" & bkmno & "' and b.employeecode='" & empcode & "' " & _
    " order by b.linenokey asc"
elseif emp = "6" then 
      sqlview = "select  distinct e.fccode id, e.fcname || '-' ||  e.fccode as text " & _
    " from employee e  " & _
    " inner join mbkmupahdtl b ON (b.EMPLOYEECODE =e.FCCODE AND b.FCBA=e.FCBA)  " & _
    " inner join mbkmhdr h on (h.bkmno=b.bkmno and h.fcba=b.fcba) " & _
    " WHERE h.bkmdate=TO_DATE('" & tglbkm & "','MM/DD/YYYY') and h.DOCUMENT_CLASSIFICATION='503' " & _
    " AND h.fcba='" & fcba & "' and h.bkmno='" & bkmno & "' " & _
    " order by e.fcname || '-' || e.fccode asc"
'  Response.write sqlview
else
    sqlview = "select e.fccode id, e.fcname || '-' ||  e.fccode as text, e.fcba, e.PRESENTGRADE,timein timeinx,timeout timeoutx,NVL(to_char(a.timein,'HH24:MI:SS'),'-')timein,NVL(to_char(a.timeout,'HH24:MI:SS'),'-')timeout " & _
    " from employee e left join USERFINGER.ATTENDANCE_TAB a ON a.EMPNO=e.FCCODE AND a.FCBA=e.FCBA AND a.TDATE=TO_DATE('" & tglbkm & "','MM/DD/YYYY') " & _
    "  where e.fcba='" & fcba & "' AND e.fcname || '-' ||  e.fccode LIKE UPPER('%" & txtcari & "%')  and e.dateterminate is null order by e.fcname asc FETCH FIRST 100 ROWS ONLY"

 end if

'   Response.write sqlview

List.Open sqlview, connStr, 1
ListCount = List.RecordCount 

Set caljson = new JSON  'JSON.asp가 필요함
Set rs = connStr.Execute(sqlview)

Dim List_jsonString2
List_jsonString2 = caljson.toJSON("results",rs, false)

Dim jsonObj2, outputObj2
set jsonObj2 = new JSONobject
set outputObj2 = jsonObj2.parse(List_jsonString2)  'json.objectClass.asp가 필요함(파싱하려면)


Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
oJSON.loadJSON(List_jsonString2)

Response.write List_jsonString2
%>