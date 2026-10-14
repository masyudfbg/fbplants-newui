<!--#include virtual="/connDb1.Inc"-->
<!--#include file="functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
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
<title>ePlaS Reporting - Project Acounting</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

dim v_txtSearch,v_qty

v_txtSearch = request("txtSearch")


txtSQL = 	" select FCCODE, FCNAME, PROJECTGROUPCODE, PROJECTTYPECODE, BUSINESSUNITREQUESTOR," & _
			" CONTROLJOBCODE, PLANSTARTDATE, VALID_FROM, VALID_UNTIL, PROJECTMANAGER,REMARKS,nvl(PROJECTQTYREQUEST,0) PROJECTQTYREQUEST," & _
			" FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME" & _
			" from PROJECTACC" & _
			" where fcba='" & session("fcBA") & "' "

	 
if  len(trim(v_txtSearch)) > 0 then
    txtSQL = txtSQL & " and fccode LIKE '%" & v_txtSearch & "%'"
end if


txtsql=(txtsql) & " order by PROJECTGROUPCODE,FCCODE"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
        <!--#include virtual="/print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PROJECT ACCOUNTING</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
						</tr>
						<!--</div>-->
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1
	Dim v_Quantity,v_qtyretrun,v_amountreceived,v_amountreturn
	v_HOURS=0
	v_UNIT=0
	v_OUTPUT=0
	v_AMOUNT=0
	%>
	</tr>
	<tr>
	    <th>PROJECT CODE</th>
		<th>PROJECT NAME</th>
		<th>PROJECT GROUP CODE</th>
		<th>PROJECT TYPE CODE</th>
		<th>BUSINESS UNIT REQUESTOR</th>
		<th>CONTROL JOB CODE</th>
		<th>PLAN START DATE</th>
		<th>VALID FROM</th>
		<th>VALID UNTIL</th>
		<th>PROJECT MANAGER</th>
		<th>REMARKS</th>
		<th>PROJECT QTY REQUEST</th>		
	</tr>
     </thead>
     <tbody>
	<%
	
	v_qty = 0
	While Not rsCat.EOF%>
	 	<tr>
		 <% for i=0 to 10 
				if i = 2 then
						q = "select FCNAME from PROJECTACCGROUP Where FCCODE = '" & rscat(i) & "' and fcba = '" & session("fcba") & "'"
							Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
							rsCallCombo.Open q, connstr
							if not rsCallCombo.eof then
							 Do While Not rsCallCombo.eof
							%>	
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat(i)%>&nbsp;.&nbsp;<%=rsCallCombo("FCNAME")%></td>
							<%rsCallCombo.movenext
							loop
						End if
						rsCallCombo.Close
						set rsCallCombo = Nothing
						q=""%>	

				<% elseif i = 6 or i = 7 or i = 8 then %>
				
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat(i)%></td>
				<% elseif i = 1 then %>
				
				<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat(i)%></td>
				
				<% else %>
				<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat(i)%></td>
				<% 
				end if
		 Next %>
		 <td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=rscat("PROJECTQTYREQUEST")%></td>
		</tr>
	  <%
	  v_qty = v_qty+cdbl(rscat("PROJECTQTYREQUEST"))
	  rsCat.MoveNext
	  response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="11"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_qty,0)%></td>

	</tr>
  </tbody>  
</table>
<!--#include virtual="/Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>