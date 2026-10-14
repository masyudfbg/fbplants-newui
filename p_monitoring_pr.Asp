<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Monitoring PR</title>
<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<style>
	/* Isian Keterangan + tombol Update di tabel (tampilan saja) */
	#product-table td textarea { width: 220px; min-height: 38px; padding: 6px 9px; font: 12.5px/1.35 inherit; vertical-align: middle; resize: vertical; }
	#product-table td button {
		min-height: 30px; margin: 4px 0 0 4px; padding: 0 12px; font: 600 12px inherit; color: #fff;
		background: #1B7863; border: 1px solid #1B7863; border-radius: 7px; cursor: pointer; vertical-align: middle;
	}
	#product-table td button:hover { background: #12584a; border-color: #12584a; }
</style>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF;font-size:1px"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtINVOICEDATE_From= trim(request("CREATIONDATE_From"))
v_txtINVOICEDATE_To= trim(request("CREATIONDATE_To"))

dim company,bulan,tahun,filtercompany1,fcba,sqlpt,userid,statusgrn,filterStatusGRN,whereClause,doc_classification,filter_doc
dim filterfcba1,filterfcba2,qq,rsq,fcuserrole

company= trim(request("company"))
fcba= trim(request("businessunit"))
statusgrn= trim(request("statusgrn"))
statusgrn= trim(request("statusgrn"))
doc_classification= trim(request("document_classification"))

userid = session("CodeName")

 q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
	set rsview = connStr.Execute(q)
	if not rsview.eof then
	   fcbaname = rsview("fcName")
	   CompanyName=rsview("FCCOMPANYNAME")
	   MgrName = rsview("FCMANAGER")
	   KTUName = rsview("FCKTU")
	   CropName=rsview("FCCROP")
	   FcbaType=rsview("fcType")
	end if

Dim sqlview,rsview
IF UCase(userid) = "DIV_HEAD_COMM" or UCase(userid) ="DEPT_HEAD_PURCH" or UCase(userid) ="STAFF_OFFICER_PURCH_02" or UCase(userid) ="STAFF_OFFICER_PURCH" or UCase(userid) ="SISKA_MEILINA_SEMBIRING" or UCase(userid) ="RINTO_PURNAMASIDI" THEN
	
	whereClause = " WHERE b.prdate BETWEEN TO_DATE('" & v_txtINVOICEDATE_From & "','MM/DD/YYYY') AND TO_DATE('" & v_txtINVOICEDATE_To & "','MM/DD/YYYY') AND b.rowstateid not in ('Rejected') AND b.document_classification not in('206','210') "
ELSE 
	whereClause = " WHERE b.prdate BETWEEN TO_DATE('" & v_txtINVOICEDATE_From & "','MM/DD/YYYY') AND TO_DATE('" & v_txtINVOICEDATE_To & "','MM/DD/YYYY') AND b.rowstateid not in ('Rejected') "
	
END IF
 	if trim(company)="all" THEN
		filtercompany1 = " and a.fcba in(select fccode from businessunit b inner join USER_FCBA_TAB ut ON b.FCCODE=ut.FCBA where ut.fcusercode='" & userid & "' and b.plantation ='FBG') "
	else
		if trim(fcba) = "all" then 
			filtercompany1 = " and a.fcba in ( select fccode from businessunit where fccompanycode='" & company & "' )"
		else 
			filtercompany1 	= " and  a.fcba ='" & fcba & "' "
			filterfcba1 	= " and  f.fcba ='" & fcba & "' "
			filterfcba2 	= " and  e.fcba ='" & fcba & "' "
		end if
    end if
	If Trim(statusgrn) = "donegrn" Then
		filterStatusGRN = " AND e.receivenotecode IS NOT NULL " ' Sudah GRN
			Else
				If Trim(statusgrn) = "notgrn" Then
					filterStatusGRN = " AND e.receivenotecode IS NULL " ' Belum GRN
			Else
					filterStatusGRN = ""
	End If

	if doc_classification<>"all" then 
		filter_doc = " AND b.DOCUMENT_CLASSIFICATION= '" & doc_classification & "'"
	end if
end if
		If filtercompany1 <> "" Then whereClause = whereClause & filtercompany1
		If filterStatusGRN <> "" Then whereClause = whereClause & filterStatusGRN	
		

sqlview = " WITH PRD_HIST AS ( " & _
" SELECT prcode, fcba, lineid, fcentry, lastupdate AS lupdate, currentauthorizelevel, " & _
" ROW_NUMBER() OVER (PARTITION BY prcode, fcba, lineid, currentauthorizelevel ORDER BY lastupdate DESC) AS rn " & _
" FROM purchaserequestdetails_hist ) " & _
" SELECT DISTINCT a.prcode " & _
"      ,    b.prdate " & _
"      ,	b.rowstateid " & _
"      ,	a.itemcode " & _
"      ,	x.fcname AS stockname " & _
"      ,    x.UNITOFMEASUREMENT uom " & _
"      ,    a.requestedqty " & _
"      ,    a.approvedqty " & _    
"      ,    sp.fcname suppliercode  " & _ 
"      ,    d.pocode  " & _
"      ,    to_char(d.podate, 'mm/dd/yyyy') as podate  " & _
"      ,    c.quantity " & _
"      ,    nvl(c.unitprice,0)unitprice " & _
"      ,    e.receivenotecode  " & _
"      ,    e.rndate, TO_CHAR(e.LASTUPDATE, 'mm/dd/yyyy') AS GRNLASTUPDATE  " & _
"      ,    f.quantity qtyGRN " & _
"	   ,    dc.fcname" &  _
"      ,    a.description "  & _
"      ,	NVL(c.keterangan, a.keterangan) as keterangan  " & _
"      ,    nvl(c.keterangan1, a.keterangan1) as keterangan1 "  & _
"      ,    a.lineid "  & _
"      ,    c.linenokey linenokeypod "  & _
"      ,    a.fcba "  & _
"      ,    bu.FCCOMPANYCODE "  & _
"      ,    h.fcentry AS fcentry_161 "  & _
"	   ,	h.lupdate AS approve_date_161 "  & _
"	   ,	i.fcentry AS fcentry_190 "  & _
"	   ,	i.lupdate AS approve_date_190 "  & _
"	   ,	j.fcentry AS fcentry_900 "  & _
"	   ,	TO_CHAR(j.lupdate, 'mm/dd/yyyy hh:mi:ss AM') AS approve_date_900 "  & _
"	   ,	bu.region AS region "  & _
"	   ,	bu.fccompanycode AS fccompanycode "  & _
"			FROM bankstock x "  & _
"			INNER JOIN purchaserequestdetails a ON x.fccode = a.itemcode "  & _
"			INNER JOIN purchaserequest b ON a.prcode = b.prcode AND a.fcba = b.fcba "  & _
"			LEFT JOIN PRD_HIST h ON b.prcode = h.prcode AND a.fcba = h.fcba AND a.lineid = h.lineid "  & _
"		    AND h.currentauthorizelevel = '161' AND h.rn = 1"  & _
"			LEFT JOIN PRD_HIST i ON a.prcode = i.prcode AND a.fcba = i.fcba AND a.lineid = i.lineid "  & _
"		    AND i.currentauthorizelevel = '190' AND i.rn = 1"  & _
"			LEFT JOIN PRD_HIST j ON b.prcode = j.prcode AND a.fcba = j.fcba AND a.lineid = j.lineid "  & _
"    		AND j.currentauthorizelevel = '900' AND j.rn = 1"  & _
"			LEFT JOIN purchaseorderdetails c ON a.prcode = c.prcode AND a.lineid = c.prlineid AND a.itemcode = c.itemcode "  & _
"			LEFT JOIN purchaseorder d ON c.pocode = d.pocode AND c.fcba = d.fcba "  & _
"         	LEFT JOIN supplier sp" & _
"            ON (d.suppliercode = sp.fccode AND d.fcba = sp.fcba)" & _
"			LEFT JOIN receivenotedetail f ON c.linenokey = f.purchaselinenokey AND a.itemcode = f.purchaseitemcode   "  & _
filterfcba1 & _
"			LEFT JOIN receivenote e ON f.receivenotecode = e.receivenotecode AND f.fcba = e.fcba "  & _
filterfcba2 & _
"		    AND e.fcba = f.fcba "  & _
"		    AND c.linenokey = f.purchaselinenokey "  & _
"		    AND a.itemcode = f.purchaseitemcode"  & _
"			LEFT JOIN ( "  & _
"				SELECT DISTINCT fccode, fcname "  & _
"				FROM document_classification WHERE DOCUMENT_TYPE='PR' "  & _
"			) dc ON dc.fccode = b.document_classification "  & _
"			LEFT JOIN BUSINESSUNIT BU ON BU.FCCODE=b.fcba" & _
				whereClause & _
				filter_doc & _
"           ORDER BY bu.region,bu.fccompanycode, a.prcode,b.prdate,a.itemcode  "
set rsview = connStr.Execute(sqlview)
'response.Write(sqlview)
response.Flush()
%>
	<body>
		<div id="wrapper">
		<!--#include file="print_excel.asp"-->
        	<table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="28" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
					<h2>MONITORING PR</h2>
					<br/>
						<h4>
							PERIODE S/D &nbsp;:&nbsp;<%=textDate(v_txtINVOICEDATE_From)%>  TO <%=textDate(v_txtINVOICEDATE_To)%>
						</h4>
				</div>
			</div>
				</th>
			</tr>
			<%
			dim sqluser,rsuser,menu_po
			sqluser = "select menu_po from users_tab where fccode='" & userid & "'"
								set rsuser = connStr.Execute(sqluser)
								if not rsuser.eof then
									menu_po = rsuser("menu_po")
								end if
			%>
				<!--<div class="printhide">-->
						<tr>
							<th colspan="28" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
						
						
						</tr> 
						<tr >
							<th rowspan="2" >COMPANY</th>
							<th rowspan="2" >FCBA</th>
							<%
                                dim userlogins11
                                userlogins11 = Session("CodeName")
                                if userlogins11="STAFF_OFFICER_PURCH" OR userlogins11="STAFF_OFFICER_PURCH_02" OR userlogins11="DIV_HEAD_COMM" OR userlogins11="IT_HO6" OR userlogins11="IT_HO2" OR userlogins11="SISKA_MEILINA_SEMBIRING" OR userlogins11="RINTO_PURNAMASIDI" OR userlogins="DEPT_HEAD_PURCH" THEN
                            %>
                                <th colspan="15" >PR</th>
                            <%else%> 
                                <th colspan="15" >PR</th>
                            <% end if%>
							<%
                                dim userlogins111
                                userlogins111 = Session("CodeName")
                                if userlogins111="STAFF_OFFICER_PURCH" OR userlogins111="STAFF_OFFICER_PURCH_02" OR userlogins111="DIV_HEAD_COMM" OR userlogins111="IT_HO6" OR userlogins="IT_HO2" OR userlogins="SISKA_MEILINA_SEMBIRING" OR userlogins111="RINTO_PURNAMASIDI" OR userlogins111="DEPT_HEAD_PURCH" THEN
                            %>
                                <th colspan="<%if menu_po="TRUE" then%>6<%else%>5<%end if%>" >PO</th>
                            <%else%> 
                                <th colspan="<%if menu_po="TRUE" then%>6<%else%>5<%end if%>" >PO</th>
                            <% end if%>
							<%
                                dim userlogins
                                userlogins = Session("CodeName")
                                if userlogins="STAFF_OFFICER_PURCH" OR userlogins="STAFF_OFFICER_PURCH_02" OR userlogins="DIV_HEAD_COMM" OR userlogins="IT_HO6" OR userlogins="IT_HO2" OR userlogins="SISKA_MEILINA_SEMBIRING" OR userlogins="RINTO_PURNAMASIDI" OR userlogins="DEPT_HEAD_PURCH" THEN
                            %>
                                <th colspan="6" >GRN</th>
                            <%else%> 
                                <th colspan="3" >GRN</th>
                            <% end if%>
							<th rowspan="2" >KETERANGAN</th>
							<th rowspan="2" >HISTORI KETERANGAN</th>
                </tr>
                <tr>
                    <th >NO.PR</th>
                    <th >TGL.PP</th>
					<th >PR Line</th>
                    <th >Kode.Barang</th>
                    <th >Nama.Barang</th>
                    <th >Satuan</th>
					<th >Keterangan</th>
					<th >Qty Diminta</th>
                    <th >Qty Approve</th>
					<th >Classification</th>
					<th >Status</th>
                    <th >Approve 1</th>
					<th >Approve 2</th>
					<th >Approve 3</th>
					<th >Dura</th>
					<th >Supplier</th>
                    <th >NO.PO</th>
                    <th >Tgl.PO</th>
					<%
                        dim podura
                        podura = Session("CodeName")
                        if podura="STAFF_OFFICER_PURCH" OR podura="STAFF_OFFICER_PURCH_02" OR podura="DIV_HEAD_COMM" OR podura="IT_HO6" OR podura="IT_HO2" OR podura="SISKA_MEILINA_SEMBIRING" OR podura="RINTO_PURNAMASIDI" OR podura="DEPT_HEAD_PURCH" THEN
                    %>
                        <th style="color:#fff;">Dura</th>
                    <%else%> 
                        <th>Dura</th>
                    <% end if%>
                    <th >Qty PO</th>
					<%if menu_po="TRUE" then%>
                    <th >Price PO</th>
					<%END IF%>
					<th >NO.GRN</th>
                    <th >Tgl.GRN</th>
					<%
                        dim userlogins1
                        userlogins1 = Session("CodeName")
                        if userlogins1="STAFF_OFFICER_PURCH" OR userlogins1="STAFF_OFFICER_PURCH_02" OR userlogins1="DIV_HEAD_COMM" OR userlogins1="IT_HO6" OR userlogins1="IT_HO2" OR userlogins1="SISKA_MEILINA_SEMBIRING" OR userlogins1="RINTO_PURNAMASIDI" OR userlogins1="DEPT_HEAD_PURCH" THEN
                    %>
						<th style="color:#fff;">Dura</th>
                        <th style="color:#fff;">LU.GRN</th>
					    <th style="color:#fff;">Dura</th>
                    <%else%> 
                    <% end if%>
                    <th rowspan="1">Qty GRN</th>
                
                        </tr>
                
                       
                         </thead>
                         <tbody>

						 <%
 
							

						 	dim no,userapproved1,userapproved2,userapproved3,tglapp1,tglapp2,tglapp3,statusauto
							
							no = 1
							While Not rsview.EOF
								statusauto = rsview("rowstateid")
								userapproved1 = rsview("fcentry_161") 
								userapproved2 = rsview("fcentry_190")
								userapproved3 = rsview("fcentry_900")
								tglapp1 = rsview("approve_date_161") 
								tglapp2 = rsview("approve_date_190")
								tglapp3 = rsview("approve_date_900")

								if statusauto = "AuthorizeOnProgress" then 
									statusauto = "Authorize"
								end if

								if userapproved1 = "OKKY_DIV_HEAD_ENGINEERING" then 
									userapproved1 = "HOE"
								end if
								if userapproved2 = "OKKY_DIV_HEAD_ENGINEERING" OR userapproved2 ="DIV_HEAD_ENGINEERING" then 
									userapproved2 = "HOE"
								end if
								if userapproved3 = "DEPT_HEAD_PURCH" OR userapproved2 ="RINTO_PURNAMASIDI" then 
									userapproved3 = "PCH"
								end if

								userapproved1 = userapproved1 & " - " & tglapp1
								userapproved2 = userapproved2 & " - " & tglapp2
								userapproved3 = userapproved3 & " - " & tglapp3
						 %>

						<tr>
							<td align="center" style="<%=cmdstyle%>" ><font face="Verdana" size="0.5">
							<%=rsview("fccompanycode")%>
							<input type="hidden"  id="fccompanycode<%=no%>" value="<%=rsview("fccompanycode")%>">
							</td> 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="0.5">
							<%=rsview("fcba")%>
							<input type="hidden" id="fcba<%=no%>" value="<%=rsview("fcba")%>">
							<input type="hidden" id="pocode<%=no%>" value="<%=rsview("pocode")%>">
							<input type="hidden" id="linenokeypod<%=no%>" value="<%=rsview("linenokeypod")%>">
							</td>  
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="0.5">
							<%=rsview("prcode")%>
							<input type="hidden" id="prcode<%=no%>" value="<%=rsview("prcode")%>">
							</td> 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="0.5"><%=rsview("prdate")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="0.5">
							<%=rsview("lineid")%>
							<input type="hidden" id="lineid<%=no%>" value="<%=rsview("lineid")%>">
							</td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="0.5">
							<%=rsview("itemcode")%>
							<input type="hidden" id="itemcode<%=no%>" value="<%=rsview("itemcode")%>">
							</td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("stockname")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("uom")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("description")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("requestedqty")%> </td>
							<td align="Right"  style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("approvedqty")%></td> 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("fcname")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=statusauto%></td>
                            <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=userapproved1%></td>
                            <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=userapproved2%></td>
                            <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=userapproved3%></td>
                            <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
								<%
									dim prdura
									prdura = Session("CodeName")
									if prdura="STAFF_OFFICER_PURCH" OR prdura="STAFF_OFFICER_PURCH_02" OR prdura="DIV_HEAD_COMM" OR prdura="IT_HO2" OR prdura="SISKA_MEILINA_SEMBIRING" OR prdura="RINTO_PURNAMASIDI" OR prdura="DEPT_HEAD_PURCH" THEN
										if Not IsNull(tglapp3) and tglapp3 <> "" then
											Response.Write DateDiff("d", rsview("prdate"), tglapp3)
										else
											Response.Write "-"
										end if
									else 
										Response.Write "-"
									end if
								%>
							</td>
							<td align="center"   style="<%=cmdstyle%>"><font face="Verdana" size="1">
							<%
								dim userlogins2
								userlogins2 = Session("CodeName")
								if userlogins2="STAFF_OFFICER_PURCH" OR userlogins2="STAFF_OFFICER_PURCH_02" OR userlogins2="DIV_HEAD_COMM" OR userlogins2="IT_HO2" OR userlogins2="SISKA_MEILINA_SEMBIRING" OR userlogins2="RINTO_PURNAMASIDI" OR userlogins2="DEPT_HEAD_PURCH" THEN
							%>
							<%=rsview("suppliercode")%>
							<%else%> 
							-
							<% end if%>
							</td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("pocode")%></td>
                            <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("podate")%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
								<%
									dim podura1
									podura1 = Session("CodeName")
									if podura1="STAFF_OFFICER_PURCH" OR podura1="STAFF_OFFICER_PURCH_02" OR podura1="DIV_HEAD_COMM" OR podura1="IT_HO6" OR podura1="IT_HO2" OR podura1="SISKA_MEILINA_SEMBIRING" OR podura1="RINTO_PURNAMASIDI" OR podura1="DEPT_HEAD_PURCH" THEN
								%>
									<%=DateDiff("d", tglapp3, rsview("podate"))%>
								<%else%> 
									-
								<% end if%>
							</td>
                            <td align="Right"  style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("quantity")%></td>
							<%if menu_po="TRUE" then%>
                            <td align="Right"  style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rsview("unitprice"),0)%></td>
							<%end if%>
                            <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("receivenotecode")%></td>
                            <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("rndate")%> </td>
							<%
								dim userlogin1
								userlogin1 = Session("CodeName")
								if userlogin1="STAFF_OFFICER_PURCH" OR userlogin1="STAFF_OFFICER_PURCH_02" OR userlogin1="DIV_HEAD_COMM" OR userlogin1="IT_HO6" OR userlogin1="IT_HO2" OR userlogin1="SISKA_MEILINA_SEMBIRING" OR userlogin1="RINTO_PURNAMASIDI" OR userlogin1="DEPT_HEAD_PURCH" THEN
							%>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=DateDiff("d", rsview("podate"), rsview("rndate"))%> </td>
								<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("GRNLASTUPDATE")%> </td>
								<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=DateDiff("d", rsview("rndate"), rsview("GRNLASTUPDATE"))%> </td>
							<% end if%>
                            <td align="Right"  style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("qtyGRN")%></td>
							<td align="left"   style="<%=cmdstyle%>"><font face="Verdana" size="1">
							<%
								dim userlogin
								userlogin = Session("CodeName")
								if userlogin="STAFF_OFFICER_PURCH" OR userlogin="STAFF_OFFICER_PURCH_02" OR userlogin="SISKA_MEILINA_SEMBIRING" OR userlogin="IT_HO2" THEN
							%>
							<textarea type="text" name="keterangan" id="keterangan<%=no%>"><%=rsview("keterangan")%></textarea> <button type="button" onclick="getKeterangan(<%=no%>)">Update</button>
							<%else%>
							<%=rsview("keterangan")%>
							<% end if%>
							</td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1">
								<%=rsview("keterangan1")%>
							</td>
						</tr>
						<%
						no = no + 1
						rsview.MoveNext
						Wend
						rsview.close()
						%>
						</tbody>
						</table>
		</div>
	</BODY>
</HTML>	

 

<script type="text/javascript">
function getKeterangan(no) {
  
        var keterangan = $('#keterangan' + no).val();
        var prcode = $('#prcode' + no).val();
        var lineid = $('#lineid' + no).val();
        var itemcode = $('#itemcode' + no).val();
	 	var fcba = $('#fcba' + no).val();
	 	var pocode = $('#pocode' + no).val();
	 	var linenokeypod = $('#linenokeypod' + no).val();

        $.ajax({
            url: "monitoringprproses.asp",
            type: "GET",
            data: {
                keterangan: keterangan,
                prcode: prcode,
                lineid: lineid,
                itemcode: itemcode,
                fcba: fcba,  
                pocode: pocode,  
                linenokeypod: linenokeypod  
            },
			success: function(response) {
            // Menampilkan pesan jika berhasil
            alert("keterangan Berhasil diperbaharui!");
			//location.reload();
        },
        error: function(xhr, status, error) {
            // Menampilkan pesan jika terjadi error
            alert("Gagal memperbaharui data: " + error);
        }
        });
   
}


function ExportExcel(fcbaname) 

	{
		var statusgrn 	    = $("#statusgrn").val();
		var tahun 	    = $("#tahun").val();
		var pcompany 	= $("#company").val();
		var pfcba 	    = $("#businessunit").val();
		
		window.open('p_monitoring_pr1.asp?statusgrn=' + fcbaname, '_blank');
		
	}

   



</script>