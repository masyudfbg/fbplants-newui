<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>FBPlants Reporting - BKM Approval</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
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

</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtPRDATE_From=trim(request("PRDATE_From"))
v_txtPRDATE_To=trim(request("PRDATE_To"))
v_txtPRCODE=trim(request("PRCODE"))
v_txtITEMCODE=trim(request("ITEMCODE"))
v_txtRequestorID=trim(request("REQUESTORID"))
v_txtRowStateId=trim(request("ROWSTATEID"))
businessunit=trim(request("businessunit"))

		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if

 dim bkmfrom,bkmto,txtfcba1,txtafd,txtstatus1,filter
    bkmfrom     = Request("bkmdate_From")
    bkmto       = Request("bkmdate_To")
    txtfcba1    = Request("businessunit")
    txtafd      = Request("sectionname")
    txtstatus1  = Request("statusapproved")

    if trim(txtafd) = "" then 
		filter = " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & Session("CodeName") & "' ) "
	else
	    filter = " AND hdr.sectioncode in(select SECTIONNAME  from user_section_tab where FCUSERCODE='" & Session("CodeName") & "' AND sectionname='" & txtafd & "') "
	end if

	q = " SELECT hdr.fcba, hdr.bkmno,hdr.rkbno, hdr.bkmdate, hdr.document_classification, hdr.sectioncode, hdr.gangcode,  hdr.document_classification," & _
		" hdr.VEHICLECODE,hdr.fcentry,hdr.fcedit,aus.authorizelevelcode,hdr.rowstate, " & _
		" hdr.fcedit,hdr.lastupdate" & _
	    " FROM bkmhdr hdr" & _
	    " LEFT JOIN authorizetemplate at ON (at.document_type = 'AA' AND at.document_classification = hdr.document_classification)" & _
		" INNER JOIN authorizestructure aus ON (aus.fcba = hdr.fcba AND aus.authorizetemplatecode = at.fccode AND aus.fcusercode = '" & Session("CodeName") & "')" & _
	    " WHERE hdr.bkmdate between to_date('" & bkmfrom & "','mm/dd/yyyy') AND to_date('" & bkmto & "','mm/dd/yyyy') AND hdr.rowstate = '" & txtstatus1 & "'" & _
		" AND at.document_classification = hdr.document_classification AND hdr.fcba='" & txtfcba1 & "' " & _
        filter & _ 
		" ORDER BY hdr.bkmdate, hdr.fcba, hdr.bkmno"


	if len(trim(txtSelect)) > 0 then
	txtSelect = "Selected  by : " & txtSelect
	else
	txtSelect = "Selected by All Data"
	end if

 'response.Write(q)
response.Flush()
%>
	<body>
		<div id="wrapper">
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>BKM APPROVAL</h2>
				</div>
			</div>
					
		</th>
        </tr>
        		<!--<div class="printhide">-->
						<tr>
							<th colspan="11"  style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
							<%	
							
							Set rsCat = Server.CreateObject("ADODB.Recordset")
							rsCat.CursorType = 3
							rsCat.CursorLocation = 3
							rsCat.Open q, connstr

							
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr>
						<tr>
							<!--<th>No.</th>-->
							<th>Kebun/Pabrik</th>
							<th>BKM Date</th>
							<th>BKM No</th>
							<th>BKM RKB</th>
							<th>BKM Type</th>
							<th>Section Name</th>
							<th>Gang</th>
							<th>Vehicle Code</th>
							<th>Status</th>
							<th>User Approved</th>
							<th>Last Time Approved</th>
							
						</tr>
                        </thead>
                        <tbody>
							<%
							Dim bkmtype
							While Not rsCat.EOF
								bkmtype = rscat("document_classification")
								if bkmtype="501" then
									bkmtype = "BKM Perawatan"
								elseif bkmtype="502" then
									bkmtype = "BKM Traksi"
								elseif bkmtype="503" then
									bkmtype = "BKM Panen"
								elseif bkmtype="504" then
									bkmtype = "BKM WIP"
								end if
							%>
							<tr>
								<!--<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%> </td>-->
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("bkmdate")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("bkmno")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rkbno")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=bkmtype%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sectioncode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("gangcode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("vehiclecode")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rowstate")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcedit")%> </td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("lastupdate")%> </td>
								
							</tr>
							  <%

							  rsCat.MoveNext
							  'response.Flush()
							  no=no+1
							Wend
							rsCat.close()
							rsCat=null%>
							<input  type="hidden" value="<%=txtTable%>" name="txtTable">
							<input  type="hidden" value="<%=txtModule%>" name="txtModule">
							</form>
                            </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	