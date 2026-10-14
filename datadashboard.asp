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


dim txtcari,fcba,bkmfrom,bkmto,idview,company,filtersection,filterrowstate
idview   = Request("idview")
'tampilkan data listing bkm'
if idview = "1" then
    tahun = Request("tahun")
    bulan = Request("month")

    Set List = Server.CreateObject("ADODB.Recordset")	

    sqlview =  " SELECT b.fccompanycode, " & _
                "     a.year, " & _
                "     a.month, " & _
                "     sc.site, " & _
                "     MIN(a.end_date_process), " & _
                "     SUM(CASE WHEN a.modulename = 'STOCK_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS STOCK_CLOSING,  " & _
                "     SUM(CASE WHEN a.modulename = 'WIP_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS WIP_CLOSING,  " & _
                "     SUM(CASE WHEN a.modulename = 'ARESTA_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS ARESTA_CLOSING,  " & _
                "     SUM(CASE WHEN a.modulename = 'KASBANK_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS KASBANK_CLOSING,  " & _
                "     SUM(CASE WHEN a.modulename = 'WSVRA_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS WSVRA_CLOSING, " & _
                "     ROUND(SUM(a.skor) / (NVL(sc.site,1) * 5), 0) AS SCORE " & _
                " FROM empallded_unit_closing a  " & _
                " LEFT JOIN businessunit b ON a.FCBA = b.FCCODE  " & _
                " LEFT JOIN ( " & _
                "         SELECT fccompanycode, COUNT(*) AS site " & _
                "         FROM businessunit " & _
                "         WHERE isactive = 'TRUE' " & _
                "           AND FCTYPE NOT IN ('HO','R') " & _
                "           AND fccode NOT IN ('ESRS','CAES','MRE','BLE','SGE','SMB','DLE','BRPS') " & _
                "         GROUP BY fccompanycode " & _
                " ) sc  ON b.fccompanycode = sc.fccompanycode " & _
                " WHERE a.year = "&tahun&" " & _
                "   AND a.month = "&bulan&" " & _
                "   AND b.fccode NOT IN ('ESRS','CAES','MRE','BLE','SGE','SMB','DLE','BRPS') and b.fccompanycode not in ('POM') " & _
                " GROUP BY b.fccompanycode, sc.site, a.year, a.month " & _
                " ORDER BY AVG(a.end_date_process - DATE '1970-01-01') ASC "

    ' Response.write sqlView

    Set caljson = new JSON
    Set rs = connStr.Execute(sqlview)

    Dim List_jsonString1
    List_jsonString1 = caljson.toJSON("results", rs, false)

    Dim jsonObj7, outputObj7
    set jsonObj7 = new JSONobject
    set outputObj7 = jsonObj7.parse(List_jsonString1)

    Set oJSON = New aspJSON
    oJSON.loadJSON(List_jsonString1)

    Response.write List_jsonString1
end if

if idview = "2" then
    tahun2 = Request("tahun")

    Set List = Server.CreateObject("ADODB.Recordset")	

    sqlview2 =  " SELECT b.fccompanycode, " & _
                "     a.year, " & _
                "     a.month, " & _
                "     sc.site, " & _
                "     SUM(CASE WHEN a.modulename = 'STOCK_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS STOCK_CLOSING,  " & _
                "     SUM(CASE WHEN a.modulename = 'WIP_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS WIP_CLOSING,  " & _
                "     SUM(CASE WHEN a.modulename = 'ARESTA_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS ARESTA_CLOSING,  " & _
                "     SUM(CASE WHEN a.modulename = 'KASBANK_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS KASBANK_CLOSING,  " & _
                "     SUM(CASE WHEN a.modulename = 'WSVRA_CLOSING' THEN a.skor ELSE 0 END) / NVL(sc.site,1) AS WSVRA_CLOSING, " & _
                "     ROUND(SUM(a.skor) / (NVL(sc.site,1) * 5), 0) AS SCORE " & _
                " FROM empallded_unit_closing a  " & _
                " LEFT JOIN businessunit b ON a.FCBA = b.FCCODE  " & _
                " LEFT JOIN ( " & _
                "         SELECT fccompanycode, COUNT(*) AS site " & _
                "         FROM businessunit " & _
                "         WHERE isactive = 'TRUE' " & _
                "           AND FCTYPE NOT IN ('HO','R') " & _
                "           AND fccode NOT IN ('ESRS','CAES','MRE','BLE','SGE','SMB','DLE','BRPS') " & _
                "         GROUP BY fccompanycode " & _
                " ) sc  ON b.fccompanycode = sc.fccompanycode " & _
                " WHERE a.year = '" & tahun2 & "' " & _
                "   AND b.fccode NOT IN ('ESRS','CAES','MRE','BLE','SGE','SMB','DLE','BRPS') and b.fccompanycode not in ('POM') " & _
                " GROUP BY b.fccompanycode, sc.site, a.year, a.month "

    ' response.write sqlView2

    Set caljson2 = new JSON
    Set rs1 = connStr.Execute(sqlview2)

    Dim List_jsonString2
    List_jsonString2 = caljson2.toJSON("results", rs1, false)

    Dim jsonObj2, outputObj2
    set jsonObj2 = new JSONobject
    set outputObj2 = jsonObj2.parse(List_jsonString2)

    Set oJSON2 = New aspJSON
    oJSON2.loadJSON(List_jsonString2)

    Response.write List_jsonString2
end if

if idview = "3" then
    tglbkm = Request("tglbkm")
    fcba = session("fcba")

    arrTgl = Split(tglbkm, "/")
    bulan = arrTgl(0)
    tahun = arrTgl(2)

    ' response.write tglbkm

    Set List = Server.CreateObject("ADODB.Recordset")	

    sqlview2 =  " SELECT STOCK_CLOSING FROM EMPALLDEDSTATUS WHERE YEAR = "& tahun &" AND MONTH = "& bulan &" AND fcba = '"& fcba &"'"

    ' response.write sqlView2

    Set caljson2 = New JSON
    Set rs3 = connStr.Execute(sqlview2)

    Dim List_jsonString3
    List_jsonString3 = caljson2.toJSON("results", rs3, False)

    Response.Write List_jsonString3
end if

if idview = "4" then
    tahun4 = Request("tahun")
    fcba4 = Request("fcba")

    If trim(fcba4) = "-" Then
		wherefcba4 = " "
	Else
		wherefcba4 = " AND fcba ='" & fcba4 & "' "
	End If

    Set List = Server.CreateObject("ADODB.Recordset")	

    sqlview4 =  "WITH PRD_HIST AS ( " &_
                "    SELECT " &_
                "        prcode, " &_
                "        fcba, " &_
                "        lineid, " &_
                "        fcentry, " &_
                "        lastupdate AS lupdate, " &_
                "        currentauthorizelevel, " &_
                "        ROW_NUMBER() OVER ( " &_
                "            PARTITION BY prcode, fcba, lineid, currentauthorizelevel " &_
                "            ORDER BY lastupdate DESC " &_
                "        ) AS rn " &_
                "    FROM purchaserequestdetails_hist " &_
                "), " &_
                "PR_NOT_IN_PO AS ( " &_
                "    SELECT DISTINCT " &_
                "        a.prcode, " &_
                "        b.prdate, " &_
                "        b.rowstateid, " &_
                "        a.itemcode, " &_
                "        x.fcname               AS stockname, " &_
                "        x.UNITOFMEASUREMENT    AS uom, " &_
                "        a.requestedqty, " &_
                "        a.approvedqty, " &_
                "        a.description, " &_
                "        a.keterangan, " &_
                "        a.keterangan1, " &_
                "        a.lineid, " &_
                "        a.fcba, " &_
                "        bu.FCCOMPANYCODE, " &_
                "        bu.region, " &_
                "        h.fcentry              AS fcentry_161, " &_
                "        h.lupdate              AS approve_date_161, " &_
                "        i.fcentry              AS fcentry_190, " &_
                "        i.lupdate              AS approve_date_190, " &_
                "        j.fcentry              AS fcentry_900, " &_
                "        j.lupdate              AS approve_date_900_raw, " &_
                "        TO_CHAR(j.lupdate, 'mm/dd/yyyy hh:mi:ss AM') AS approve_date_900, " &_
                "        dc.fcname              AS doc_classification_name, " &_
                "        CASE " &_
                "            WHEN j.lupdate IS NOT NULL " &_
                "                THEN ROUND(j.lupdate - TRUNC(b.prdate)) " &_
                "            ELSE " &_
                "                ROUND(TRUNC(SYSDATE) - TRUNC(b.prdate)) " &_
                "        END AS days_pr_to_approve_900, " &_
                "        CASE " &_
                "            WHEN j.lupdate IS NOT NULL THEN 'Sudah Approve' " &_
                "            ELSE 'Belum Approve' " &_
                "        END AS status_approve_900, " &_
                "        ROW_NUMBER() OVER ( " &_
                "            PARTITION BY bu.region, bu.fccompanycode " &_
                "            ORDER BY a.requestedqty DESC " &_
                "        ) AS rn_top " &_
                "    FROM bankstock x " &_
                "    INNER JOIN purchaserequestdetails a " &_
                "        ON x.fccode = a.itemcode " &_
                "    INNER JOIN purchaserequest b " &_
                "        ON a.prcode = b.prcode " &_
                "        AND a.fcba  = b.fcba " &_
                "    LEFT JOIN PRD_HIST h " &_
                "        ON b.prcode = h.prcode " &_
                "        AND a.fcba  = h.fcba " &_
                "        AND a.lineid = h.lineid " &_
                "        AND h.currentauthorizelevel = '161' " &_
                "        AND h.rn = 1 " &_
                "    LEFT JOIN PRD_HIST i " &_
                "        ON a.prcode = i.prcode " &_
                "        AND a.fcba  = i.fcba " &_
                "        AND a.lineid = i.lineid " &_
                "        AND i.currentauthorizelevel = '190' " &_
                "        AND i.rn = 1 " &_
                "    LEFT JOIN PRD_HIST j " &_
                "        ON b.prcode = j.prcode " &_
                "        AND a.fcba  = j.fcba " &_
                "        AND a.lineid = j.lineid " &_
                "        AND j.currentauthorizelevel = '900' " &_
                "        AND j.rn = 1 " &_
                "    LEFT JOIN ( " &_
                "        SELECT DISTINCT fccode, fcname " &_
                "        FROM document_classification " &_
                "        WHERE DOCUMENT_TYPE='PR' " &_
                "    ) dc ON dc.fccode = b.document_classification " &_
                "    LEFT JOIN BUSINESSUNIT bu " &_
                "        ON bu.FCCODE = b.fcba " &_
                "    WHERE NOT EXISTS ( " &_
                "        SELECT 1 " &_
                "        FROM purchaseorderdetails c " &_
                "        WHERE c.prcode   = a.prcode " &_
                "          AND c.prlineid = a.lineid " &_
                "          AND c.itemcode = a.itemcode " &_
                "    ) " &_
                "    AND b.prdate >= TRUNC(SYSDATE, 'YYYY') " &_
                "    AND b.prdate <  ADD_MONTHS(TRUNC(SYSDATE, 'YYYY'), 12) " &_
                "    AND j.fcentry NOT IN ('DEPT_HEAD_PURCH', 'RINTO_PURNAMASIDI') " &_
                "    AND b.rowstateid = 'Approved' AND dc.FCCODE IN (203, 204, 205, 211, 212, 213, 214, 215, 216, 217) " &_
                "), " &_
                "PR_DEDUP AS ( " &_
                "    SELECT " &_
                "        region, " &_
                "        fccompanycode, " &_
                "        prcode, " &_
                "        prdate, " &_
                "        rowstateid, " &_
                "        itemcode, " &_
                "        stockname, " &_
                "        uom, " &_
                "        requestedqty, " &_
                "        approvedqty, " &_
                "        description, " &_
                "        keterangan, " &_
                "        keterangan1, " &_
                "        lineid, " &_
                "        fcba, " &_
                "        fcentry_161, " &_
                "        approve_date_161, " &_
                "        fcentry_190, " &_
                "        approve_date_190, " &_
                "        fcentry_900, " &_
                "        approve_date_900, " &_
                "        status_approve_900, " &_
                "        days_pr_to_approve_900, " &_
                "        doc_classification_name, " &_
                "        ROW_NUMBER() OVER ( " &_
                "            PARTITION BY prcode " &_
                "            ORDER BY days_pr_to_approve_900 DESC " &_
                "        ) AS rn_dedup " &_
                "    FROM PR_NOT_IN_PO " &_
                ") " &_
                "SELECT " &_
                "    region, " &_
                "    fccompanycode, " &_
                "    prcode, " &_
                "    prdate, " &_
                "    rowstateid, " &_
                "    itemcode, " &_
                "    stockname, " &_
                "    uom, " &_
                "    requestedqty, " &_
                "    approvedqty, " &_
                "    description, " &_
                "    keterangan, " &_
                "    keterangan1, " &_
                "    lineid, " &_
                "    fcba, " &_
                "    fcentry_161, " &_
                "    approve_date_161, " &_
                "    fcentry_190, " &_
                "    approve_date_190, " &_
                "    fcentry_900, " &_
                "    approve_date_900, " &_
                "    status_approve_900, " &_
                "    days_pr_to_approve_900, " &_
                "    doc_classification_name " &_
                "FROM PR_DEDUP " &_
                "WHERE rn_dedup = 1 " & wherefcba4 &_
                "ORDER BY days_pr_to_approve_900 DESC " &_
                "FETCH FIRST 10 ROWS ONLY "

    'response.write sqlview4

    Set caljson4 = new JSON
    Set rs4 = connStr.Execute(sqlview4)

    Dim List_jsonString4
    List_jsonString4 = caljson4.toJSON("results", rs4, false)

    Dim jsonObj4, outputObj4
    set jsonObj4 = new JSONobject
    set outputObj4 = jsonObj4.parse(List_jsonString4)

    Set oJSON4 = New aspJSON
    oJSON4.loadJSON(List_jsonString4)

    Response.write List_jsonString4
end if

if idview = "5" then
    tahun5 = Request("tahun")
    fcba5 = Request("fcba")

    If trim(fcba5) = "-" Then
		wherefcba5 = " "
	Else
		wherefcba5 = " AND fcba ='" & fcba5 & "' "
	End If

    Set List = Server.CreateObject("ADODB.Recordset")	

    sqlview5 =  "WITH PRD_HIST AS (" & _
                "    SELECT" & _
                "        prcode," & _
                "        fcba," & _
                "        lineid," & _
                "        fcentry," & _
                "        lastupdate AS lupdate," & _
                "        currentauthorizelevel," & _
                "        ROW_NUMBER() OVER (" & _
                "            PARTITION BY prcode, fcba, lineid, currentauthorizelevel" & _
                "            ORDER BY lastupdate DESC" & _
                "        ) AS rn" & _
                "    FROM purchaserequestdetails_hist" & _
                ")," & _
                "PR_GRN_CHECK AS (" & _
                "    SELECT DISTINCT" & _
                "        a.prcode," & _
                "        b.prdate," & _
                "        b.rowstateid," & _
                "        a.itemcode," & _
                "        x.fcname                                        AS stockname," & _
                "        x.UNITOFMEASUREMENT                             AS uom," & _
                "        a.requestedqty," & _
                "        a.approvedqty," & _
                "        sp.fcname                                       AS suppliercode," & _
                "        d.pocode," & _
                "        TO_CHAR(d.podate, 'mm/dd/yyyy')                 AS podate," & _
                "        c.quantity," & _
                "        NVL(c.unitprice, 0)                             AS unitprice," & _
                "        e.receivenotecode," & _
                "        e.rndate," & _
                "        TO_CHAR(e.LASTUPDATE, 'mm/dd/yyyy')             AS grnlastupdate," & _
                "        f.quantity                                      AS qtygrn," & _
                "        dc.fcname                                       AS doc_classification_name," & _
                "        a.description," & _
                "        NVL(c.keterangan, a.keterangan)                 AS keterangan," & _
                "        NVL(c.keterangan1, a.keterangan1)               AS keterangan1," & _
                "        a.lineid," & _
                "        c.linenokey                                     AS linenokeypod," & _
                "        a.fcba," & _
                "        bu.FCCOMPANYCODE," & _
                "        bu.region," & _
                "        h.fcentry                                       AS fcentry_161," & _
                "        h.lupdate                                       AS approve_date_161," & _
                "        i.fcentry                                       AS fcentry_190," & _
                "        i.lupdate                                       AS approve_date_190," & _
                "        j.fcentry                                       AS fcentry_900," & _
                "        TO_CHAR(j.lupdate, 'mm/dd/yyyy hh:mi:ss AM')    AS approve_date_900," & _
                "        CASE" & _
                "            WHEN e.rndate IS NOT NULL THEN 'Sudah GRN'" & _
                "            ELSE 'Belum GRN'" & _
                "        END AS status_grn," & _
                "        CASE" & _
                "            WHEN e.rndate IS NOT NULL" & _
                "                THEN ROUND(TRUNC(e.rndate) - TRUNC(d.podate))" & _
                "            ELSE" & _
                "                ROUND(TRUNC(SYSDATE) - TRUNC(d.podate))" & _
                "        END AS days_po_to_grn, d.purchaseby, d.deliverydestination" & _
                "    FROM bankstock x" & _
                "    INNER JOIN purchaserequestdetails a" & _
                "        ON x.fccode = a.itemcode" & _
                "    INNER JOIN purchaserequest b" & _
                "        ON a.prcode = b.prcode" & _
                "        AND a.fcba  = b.fcba" & _
                "    LEFT JOIN PRD_HIST h" & _
                "        ON b.prcode = h.prcode" & _
                "        AND a.fcba  = h.fcba" & _
                "        AND a.lineid = h.lineid" & _
                "        AND h.currentauthorizelevel = '161'" & _
                "        AND h.rn = 1" & _
                "    LEFT JOIN PRD_HIST i" & _
                "        ON a.prcode = i.prcode" & _
                "        AND a.fcba  = i.fcba" & _
                "        AND a.lineid = i.lineid" & _
                "        AND i.currentauthorizelevel = '190'" & _
                "        AND i.rn = 1" & _
                "    LEFT JOIN PRD_HIST j" & _
                "        ON b.prcode = j.prcode" & _
                "        AND a.fcba  = j.fcba" & _
                "        AND a.lineid = j.lineid" & _
                "        AND j.currentauthorizelevel = '900'" & _
                "        AND j.rn = 1" & _
                "    LEFT JOIN purchaseorderdetails c" & _
                "        ON a.prcode    = c.prcode" & _
                "        AND a.lineid   = c.prlineid" & _
                "        AND a.itemcode = c.itemcode" & _
                "    LEFT JOIN purchaseorder d" & _
                "        ON c.pocode = d.pocode" & _
                "        AND c.fcba  = d.fcba" & _
                "    LEFT JOIN supplier sp" & _
                "        ON d.suppliercode = sp.fccode" & _
                "        AND d.fcba        = sp.fcba" & _
                "    LEFT JOIN ( " &_
                "       SELECT purchaselinenokey, " &_
                "                purchaseitemcode, " &_
                "                fcba, " &_
                "                RECEIVENOTECODE, " &_
                "                SUM(quantity) AS quantity " &_
                "       FROM receivenotedetail " &_ 
                "        GROUP BY purchaselinenokey, purchaseitemcode, fcba, RECEIVENOTECODE " &_
                "    ) f ON c.linenokey = f.purchaselinenokey " &_
                "    AND a.itemcode = f.purchaseitemcode " &_
                "    LEFT JOIN receivenote e ON f.receivenotecode = e.receivenotecode " & _
                "        AND f.fcba           = e.fcba" & _
                "    LEFT JOIN (" & _
                "        SELECT DISTINCT fccode, fcname" & _
                "        FROM document_classification" & _
                "        WHERE DOCUMENT_TYPE = 'PR'" & _
                "    ) dc ON dc.fccode = b.document_classification" & _
                "    LEFT JOIN BUSINESSUNIT bu" & _
                "        ON bu.FCCODE = b.fcba" & _
                "    WHERE b.prdate >= TRUNC(SYSDATE, 'YYYY') AND b.prdate < ADD_MONTHS(TRUNC(SYSDATE, 'YYYY'), 12)" & _
                "      AND d.rowstateid = 'Ordered' AND dc.FCCODE IN (203, 204, 205, 211, 212, 213, 214, 215, 216, 217)" & _
                ")," & _
                "PR_DEDUP AS (" & _
                "    SELECT" & _
                "        region," & _
                "        fccompanycode," & _
                "        prcode," & _
                "        prdate," & _
                "        rowstateid," & _
                "        itemcode," & _
                "        stockname," & _
                "        uom," & _
                "        requestedqty," & _
                "        approvedqty," & _
                "        suppliercode," & _
                "        pocode," & _
                "        podate," & _
                "        quantity," & _
                "        unitprice," & _
                "        receivenotecode," & _
                "        rndate," & _
                "        grnlastupdate," & _
                "        qtygrn," & _
                "        doc_classification_name," & _
                "        description," & _
                "        keterangan," & _
                "        keterangan1," & _
                "        lineid," & _
                "        linenokeypod," & _
                "        fcba," & _
                "        fcentry_161," & _
                "        approve_date_161," & _
                "        fcentry_190," & _
                "        approve_date_190," & _
                "        fcentry_900," & _
                "        approve_date_900," & _
                "        status_grn," & _
                "        days_po_to_grn, purchaseby, deliverydestination, " & _
                "        ROW_NUMBER() OVER (" & _
                "            PARTITION BY prcode, lineid" & _
                "            ORDER BY" & _
                "                CASE WHEN rndate IS NOT NULL THEN 0 ELSE 1 END," & _
                "                days_po_to_grn DESC" & _
                "        ) AS rn_dedup" & _
                "    FROM PR_GRN_CHECK" & _
                ") " & _
                "SELECT DISTINCT" & _
                "    fccompanycode," & _
                "    fcba," & _
                "    pocode," & _
                "    days_po_to_grn," & _
                "    status_grn," & _
                "    purchaseby," & _
                "    deliverydestination" & _
                " FROM PR_DEDUP" & _
                " WHERE rn_dedup = 1 AND status_grn = 'Belum GRN' AND approve_date_900 IS NOT NULL AND pocode IS NOT NULL " & wherefcba5 & _
                " ORDER BY POCODE ASC" & _
                " FETCH FIRST 10 ROWS ONLY"

    ' response.write sqlview5

    Set caljson5 = new JSON
    Set rs5 = connStr.Execute(sqlview5)

    Dim List_jsonString5
    List_jsonString5 = caljson5.toJSON("results", rs5, false)

    Dim jsonObj5, outputObj5
    set jsonObj5 = new JSONobject
    set outputObj5 = jsonObj5.parse(List_jsonString5)

    Set oJSON5 = New aspJSON
    oJSON5.loadJSON(List_jsonString5)

    Response.write List_jsonString5
end if
%>