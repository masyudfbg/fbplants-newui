
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
upahedit    = Request("upahedit")
bkmno       = Request("bkmno")






if upahedit = "2" then 
  
   sqlview = "SELECT u.BKMNO, u.EMPLOYEECODE,e.fccode || '-' ||  e.fcname as fcname, u.ACTCODE,u.ACTCODE || '-' ||  a.fcname as activityname, u.MANDAYS, " & _
   " u.ALLOWANCE_CODE, u.ALLOWANCE_AMT, u.ATTENDANCE absen,  " & _
   " u.LOCATIONCODE, u.OUTPUT  " & _
   " FROM BKMUPAHDTL u " & _
   " left join employee e on (u.EMPLOYEECODE=e.fccode and u.fcba=e.fcba) " & _
   " left join ACTMASTER a on (u.ACTCODE=a.FCCODE) " & _
   " left join BKMPROSES p on (u.bkmno=p.bkmno and u.LINENOKEY=p.LINENOKEY) " & _
   " where u.BKMNO='" & bkmno & "' "
  
 
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

end if

%>