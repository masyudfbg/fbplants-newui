
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
validasi = Request("validasi")
businessunit = Request("businessunit")
fcba = Request("fcBA")

Set List = Server.CreateObject("ADODB.Recordset")

if validasi = 1 then 
    sqlview =  "select h.suppliercode,s.fcname,h.totalcharge" & _
    " from purchaseorder h,supplier s,purchaseorderdetails d " & _
    " where h.POCODE=d.POCODE and h.suppliercode=s.fccode and h.fcba=s.fcba " & _ 
    " and d.quantity-d.qtypayreq>0 and h.pocode='" & pocode & "'"

    List.Open sqlview, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString
    List_jsonString = caljson.toJSON("dataview",rs, false)

    Dim jsonObj, outputObj
    set jsonObj = new JSONobject
    set outputObj = jsonObj.parse(List_jsonString)  'json.objectClass.asp가 필요함(파싱하려면)


    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString)

    Response.write List_jsonString
end if
if validasi = 2 then  

    sqlview1 =  "select d.prcode,d.pocode,d.itemcode,s.fcname,d.quantity,d.unitprice,d.amount,d.qtypayreq,d.quantity-d.qtypayreq as qtysisa, s.flaq " & _
    " from purchaseorderdetails d,stock s " & _
    " where d.itemcode=s.fccode and d.pocode='" & pocode & "' " & _
    " and d.fcba=s.fcba and d.quantity-d.qtypayreq>0 AND d.fcba='" & fcba & "'"
    
    List.Open sqlview1, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview1)

    Dim List_jsonString2
    List_jsonString2 = caljson.toJSON("dataview",rs, false)

    Dim jsonObj2, outputObj2
    set jsonObj2 = new JSONobject
    set outputObj2 = jsonObj2.parse(List_jsonString2)  'json.objectClass.asp가 필요함(파싱하려면)

    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString2)

    Response.write List_jsonString2

end if

'query monitoring timbang TBS'
dim tglwbnow,sqlwb,rswb,fcbasession,wbdate,tglsplit,blnsplit,thnsplit,bulanthnskrg,tglbulanthnawal
fcbasession = Session("fcBA")
sqlwb = "select WBDATE FROM configmst"
Set rswb = Server.CreateObject("ADODB.Recordset")
rswb.CursorType = 3
rswb.CursorLocation = 3
rswb.Open sqlwb, connstr
if not rswb.eof then
    wbdate         = rswb("WBDATE")
    tglsplit       = left(wbdate,2)
    blnsplit       = mid(wbdate,4,2)
    thnsplit       = right(wbdate,4)
    bulanthnskrg        = "01" & "/" & blnsplit & "/" & thnsplit
    tglbulanthnawal     = "01" & "/" & "01" & "/" & thnsplit
    tglsplit    = wbdate
end if

if validasi = 3 then  
   '' tglwbnow = FormatDateTime(Now, vbShortDate)
    
  '   sqlview1 =  "SELECT ar.AreaTypeName, SUM(CASE WHEN d.DocketDate=TO_DATE('" & wbdate & "','dd/mm/yyyy') AND d.FCBA='ARM' then f.FFBWeight else 0 end) AS TonaseARM, " & _
  '  " SUM(CASE WHEN d.DocketDate=TO_DATE('" & wbdate & "','dd/mm/yyyy') AND d.FCBA='R30' then f.FFBWeight else 0 end) AS TonaseRAM," & _ 
  '  " SUM(CASE WHEN d.DocketDate=TO_DATE('" & wbdate & "','dd/mm/yyyy') AND d.FCBA='BRE' then f.FFBWeight else 0 end) AS TonaseBRE " & _ 
  '  " FROM LHPTIKETHDR  d, LHPFFBWBTRN  f, LHPAREAMST  a, LHPAREATYPEREF  ar " & _
  '  " WHERE d.DocketID=f.DocketID AND f.AreaID=a.AreaID AND f.fcba=a.fcba AND a.AreaTypeID=ar.AreaTypeID  " & _
  '  " AND d.IsInsideFlg=0 AND d.DocketDate=to_date('" & wbdate & "','dd/mm/yyyy') and d.fcba=f.fcba" & _
  '  " GROUP BY ar.AREATYPEID,ar.AreaTypeName order by ar.AREATYPEID asc"

  sqlview1 = "SELECT t.AreaTypeName, sum(t.TonaseARMHI)TonaseARMHI,sum(t.TonaseARMSHI)TonaseARMSHI,sum(t.TonaseARMSBI)TonaseARMSBI,sum(t.TonaseRAMHI)TonaseRAMHI, " & _
    " sum(t.TonaseRAMSHI)TonaseRAMSHI,sum(t.TonaseRAMSBI)TonaseRAMSBI,SUM(t.TonaseBREHI)TonaseBREHI,sum(t.TonaseBRESHI)TonaseBRESHI,sum(t.TonaseBRESBI)TonaseBRESBI " & _
    " FROM( " & _
    "    SELECT ar.AreaTypeName, " & _
    "    SUM(CASE WHEN d.DocketDate= TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.FFBWeight else 0 end) AS TonaseARMHI,  " & _
    "    SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & bulanthnskrg & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.FFBWeight else 0 end) AS TonaseARMSHI, " & _
    "    SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & tglbulanthnawal & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.FFBWeight else 0 end) AS TonaseARMSBI, " & _
    "    0 TonaseRAMHI,0 TonaseRAMSHI,0 TonaseRAMSBI,0 TonaseBREHI,0 TonaseBRESHI,0 TonaseBRESBI " & _
    "    FROM LHPTIKETHDR  d " & _
    "    INNER JOIN LHPFFBWBTRN  f ON d.DOCKETID=f.DOCKETID AND d.fcba=f.fcba  " & _
    "    INNER JOIN LHPAREAMST  a ON a.AREAID=f.AREAID AND a.FCBA=f.FCBA " & _
    "    INNER JOIN LHPAREATYPEREF  ar ON ar.AREATYPEID=a.AREATYPEID " & _
    "    WHERE d.fcba='ARM' AND d.IsInsideFlg=0  " & _
    "    GROUP BY ar.AREATYPENAME " & _
    "       union all " & _
    "    SELECT ar.AreaTypeName, " & _
    "    0 TonaseARMHI,0 TonaseARMSHI,0 TonaseARMSBI, " & _
    "    SUM(CASE WHEN d.DocketDate= TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.FFBWEIGHT else 0 end) AS TonaseRAMHI,  " & _
    "    SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & bulanthnskrg & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy') then f.FFBWEIGHT else 0 end) AS TonaseRAMSHI,  " & _
    "    SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & tglbulanthnawal & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.FFBWEIGHT else 0 end) AS TonaseRAMSBI, " & _
    "    0 TonaseBREHI,0 TonaseBRESHI,0 TonaseBRESBI " & _
    "    FROM LHPTIKETHDR  d " & _
    "    INNER JOIN LHPFFBWBTRN  f ON d.DOCKETID=f.DOCKETID AND d.fcba=f.fcba " & _
    "    INNER JOIN LHPAREAMST  a ON a.AREAID=f.AREAID AND a.FCBA=f.FCBA " & _
    "    INNER JOIN LHPAREATYPEREF  ar ON ar.AREATYPEID=a.AREATYPEID " & _
    "    WHERE d.fcba='R30' " & _
    "    AND d.IsInsideFlg=0  " & _
    "    GROUP BY ar.AREATYPENAME " & _
    "    union all " & _
    "    SELECT ar.AreaTypeName, " & _
    "    0 TonaseARMHI,0 TonaseARMSHI,0 TonaseARMSBI,0 TonaseRAMHI,0 TonaseRAMSHI,0 TonaseRAMSBI, " & _
    "    SUM(CASE WHEN d.DocketDate= TO_DATE('" & wbdate & "','dd/mm/yyyy')  then d.Weightout-d.Weightin else 0 end) AS TonaseBREHI, " & _
    "    SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & bulanthnskrg & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy') then d.Weightout-d.Weightin else 0 end) AS TonaseBRESHI, " & _
    "    SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & tglbulanthnawal & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then d.Weightout-d.Weightin else 0 end) AS TonaseBRESBI " & _
    "    FROM LHPTIKETHDR  d " & _
    "    INNER JOIN LHPFFBWBTRN  f ON d.DOCKETID=f.DOCKETID AND d.fcba=f.fcba  " & _
    "    INNER JOIN LHPAREAMST  a ON a.AREAID=f.AREAID AND a.FCBA=f.FCBA " & _
    "    INNER JOIN LHPAREATYPEREF  ar ON ar.AREATYPEID=a.AREATYPEID " & _
    "    WHERE d.fcba='BRE' " & _
    "    AND d.IsInsideFlg=0   " & _
    "    GROUP BY ar.AREATYPENAME " & _
    " )t " & _
    " group by t.AREATYPENAME"

    'Response.Write sqlview1
   
    List.Open sqlview1, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview1)

    Dim List_jsonString3
    List_jsonString3 = caljson.toJSON("dataviewtbs",rs, false)

    Dim jsonObj3, outputObj3
    set jsonObj3 = new JSONobject
    set outputObj3 = jsonObj3.parse(List_jsonString3)  'json.objectClass.asp가 필요함(파싱하려면)

    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString3)

    Response.write List_jsonString3

end if

'query monitoring  timbang CPO'
if validasi = 4 then  
    tglwbnow = FormatDateTime(Now, vbShortDate)
    sqlview1 =  "SELECT sum(t.TonaseCPOARMHI)TonaseCPOARMHI,sum(t.TonaseCPOARMSHI)TonaseCPOARMSHI,sum(t.TonaseCPOARMSBI)TonaseCPOARMSBI " & _ 
    " FROM( " & _
    " SELECT " & _
    " SUM(CASE WHEN d.DocketDate= TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonaseCPOARMHI,  " & _
    " SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & bulanthnskrg & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonaseCPOARMSHI, " & _
    " SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & tglbulanthnawal & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonaseCPOARMSBI " & _
    " FROM LHPTIKETHDR  d " & _
    " INNER JOIN LHPTIKETCPO f ON d.DOCKETID=f.DOCKETID AND d.fcba=f.fcba  " & _
    " WHERE d.fcba='ARM' AND d.IsInsideFlg=0 AND d.Productid='1'  " & _  
    " )t"
   
    List.Open sqlview1, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview1)

    Dim List_jsonString4
    List_jsonString4 = caljson.toJSON("dataviewcpo",rs, false)

    Dim jsonObj4, outputObj4
    set jsonObj4 = new JSONobject
    set outputObj4 = jsonObj4.parse(List_jsonString4)  'json.objectClass.asp가 필요함(파싱하려면)

    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString4)

    Response.write List_jsonString4

end if

'query monitoring  timbang PK'
if validasi = 5 then  
    tglwbnow = FormatDateTime(Now, vbShortDate)
  ''  fcbasession = Session("fcBA")
    sqlview1 =  "SELECT sum(t.TonasePKARMHI)TonasePKARMHI,sum(t.TonasePKARMSHI)TonasePKARMSHI,sum(t.TonasePKARMSBI)TonasePKARMSBI " & _ 
    " FROM( " & _
    " SELECT " & _
    " SUM(CASE WHEN d.DocketDate= TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonasePKARMHI,  " & _
    " SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & bulanthnskrg & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonasePKARMSHI, " & _
    " SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & tglbulanthnawal & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonasePKARMSBI " & _
    " FROM LHPTIKETHDR  d " & _
    " LEFT JOIN LHPTIKETKERNEL f ON d.DOCKETID=f.DOCKETID AND d.fcba=f.fcba  " & _
    " WHERE d.fcba='ARM' AND d.IsInsideFlg=0 AND d.Productid='3'  " & _  
    " )t"
  ' response.write sqlview1
    List.Open sqlview1, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview1)

    Dim List_jsonString5
    List_jsonString5 = caljson.toJSON("dataviewpk",rs, false)

    Dim jsonObj5, outputObj5
    set jsonObj5 = new JSONobject
    set outputObj5 = jsonObj5.parse(List_jsonString5)  'json.objectClass.asp가 필요함(파싱하려면)

    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString5)

    Response.write List_jsonString5

end if

'query monitoring  timbang CKG'
if validasi = 6 then  
    tglwbnow = FormatDateTime(Now, vbShortDate)
   '' fcbasession = Session("fcBA")
    sqlview1 =  "SELECT sum(t.TonaseCKGARMHI)TonaseCKGARMHI,sum(t.TonaseCKGARMSHI)TonaseCKGARMSHI,sum(t.TonaseCKGARMSBI)TonaseCKGARMSBI " & _ 
    " FROM( " & _
    " SELECT " & _
    " SUM(CASE WHEN d.DocketDate= TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonaseCKGARMHI,  " & _
    " SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & bulanthnskrg & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonaseCKGARMSHI, " & _
    " SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & tglbulanthnawal & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonaseCKGARMSBI " & _
    " FROM LHPTIKETHDR  d " & _
    " LEFT JOIN LHPTIKETCKG f ON d.DOCKETID=f.DOCKETID AND d.fcba=f.fcba  " & _
    " WHERE d.fcba='ARM' AND d.IsInsideFlg=0 AND d.Productid='5'  " & _  
    " )t"
   
    List.Open sqlview1, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview1)

    Dim List_jsonString6
    List_jsonString6 = caljson.toJSON("dataviewckg",rs, false)

    Dim jsonObj6, outputObj6
    set jsonObj6 = new JSONobject
    set outputObj6 = jsonObj6.parse(List_jsonString6)  'json.objectClass.asp가 필요함(파싱하려면)

    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString6)

    Response.write List_jsonString6

end if

'query monitoring  timbang Lain-Lain'
if validasi = 7 then  
    tglwbnow = FormatDateTime(Now, vbShortDate)
   '' fcbasession = Session("fcBA")
    sqlview1 =  "SELECT sum(t.TonaseLainARMHI)TonaseLainARMHI,sum(t.TonaseLainARMSHI)TonaseLainARMSHI,sum(t.TonaseLainARMSBI)TonaseLainARMSBI " & _ 
    " FROM( " & _
    " SELECT " & _
    " SUM(CASE WHEN d.DocketDate= TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonaseLainARMHI,  " & _
    " SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & bulanthnskrg & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonaseLainARMSHI, " & _
    " SUM(CASE WHEN d.DocketDate BETWEEN TO_DATE('" & tglbulanthnawal & "','dd/mm/yyyy') AND TO_DATE('" & wbdate & "','dd/mm/yyyy')  then f.Netweight else 0 end) AS TonaseLainARMSBI " & _
    " FROM LHPTIKETHDR  d " & _
    " LEFT JOIN LHPTIKETRETAIL f ON d.DOCKETID=f.DOCKETID AND d.fcba=f.fcba  " & _
    " WHERE d.fcba='ARM' AND d.IsInsideFlg=0 AND d.Productid='10'  " & _  
    " )t"
   
    List.Open sqlview1, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview1)

    Dim List_jsonString7
    List_jsonString7 = caljson.toJSON("dataviewretail",rs, false)

    Dim jsonObj7, outputObj7
    set jsonObj7 = new JSONobject
    set outputObj7 = jsonObj7.parse(List_jsonString7)  'json.objectClass.asp가 필요함(파싱하려면)

    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString7)

    Response.write List_jsonString7

end if

'query monitoring  kendaraan di dalam
if validasi = 8 then  
    'fcbasession = Session("fcBA")
    sqlview1 =  "SELECT d.FCBA, d.DocketID, d.NoReg, d.ProductID, " & _
    " p.PRODUCTNAME,  " & _
    " d.InWBID, d.DriverName, d.TimeIn, d.WeightIn  " & _
    " FROM LHPTIKETHDR d,LHPPRODUCT p " & _
    " WHERE d.PRODUCTID=p.PRODUCTID and d.IsInsideFlg=1 " & _
    " ORDER BY d.FCBA,d.DocketID DESC"
   
    List.Open sqlview1, connStr, 1
    ListCount = List.RecordCount 

    Set caljson = new JSON  'JSON.asp가 필요함
    Set rs = connStr.Execute(sqlview1)

    Dim List_jsonString8
    List_jsonString8 = caljson.toJSON("datain",rs, false)

    Dim jsonObj8, outputObj8
    set jsonObj8 = new JSONobject
    set outputObj8 = jsonObj8.parse(List_jsonString8)  'json.objectClass.asp가 필요함(파싱하려면)

    Set oJSON = New aspJSON     'aspJSON.asp 필요함(for문)
    oJSON.loadJSON(List_jsonString8)

    Response.write List_jsonString8

end if
%>