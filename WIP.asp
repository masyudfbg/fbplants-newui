<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Contract Work in Progress</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
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
<body>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <%
 '-------------awal-----------------------
Dim FCCODE,RETENTION,QTY,RATE,AGREEMENTLINENO,WIPDATE
Dim FDDATE, JOB
Dim LOCATIONTYPE
Dim LOCATIONCode,UNIT, OUTPUT, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA
Dim cmdView
Dim CheckQTY, rsCheckQTY, varSelisih, varTotalAgreement, rsCheckWIPQTY, CheckWIPQTY, varWIPQTY
Dim varError
Dim getAgreementCode, rsAgreementCode
Dim checkRate, rsCheckRate, varRate
Dim strCheckDate, rsCheckDate, varRemain, varDueDate

fcCode =request("fcCode")
AGREEMENTLINENO=request("AGREEMENTLINENO")
WIPDATE=REQUEST("WIPDATE")
QTY=request("QTY")
REFERENCE=request("REFERENCE")
REMARKS=request("REMARKS")
		
if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	" DELETE FROM WORKINPROGRESS  WHERE" & _
				" WIPDATE = TO_DATE('" & WIPDATE & "','mm/dd/yyyy')" & _
				" AND AGREEMENTCODE = '" & FCCODE & "'" & _
				" AND AGREEMENTLINENO = '" & AGREEMENTLINENO & "'" & _
				" AND FCBA = '" & Session("fcBA") & "'"
				
		'Response.Write q
		ConnStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")="New" then	
		'Validasi Qty
		CheckQTY = "SELECT   fcba, agreementcode, lineno, SUM (totalagreement) totalagreement," & _
					" SUM (totalwip) totalwip,   SUM (totalagreement) - SUM (totalwip) selisih" & _
					" FROM ((SELECT   fcba, agreementcode, lineno, SUM (qty) totalagreement, 0 totalwip" & _
					" FROM agreementdetail GROUP BY fcba, agreementcode, lineno) UNION (" &_
					" SELECT   fcba, agreementcode, agreementlineno lineno, 0 totalagreement, SUM (qty) totalwip" & _
					" FROM workinprogress GROUP BY fcba, agreementcode, agreementlineno)) dataall" & _
					" WHERE fcba = '" & Session("fcBA") & "' AND agreementcode = '" & FCCODE & "'" & _
					" AND lineno = '" & AGREEMENTLINENO & "' GROUP BY fcba, agreementcode, lineno"
		
		Set rsCheckQTY = Server.CreateObject("ADODB.Recordset")
		rsCheckQTY.CursorType = 3
		rsCheckQTY.CursorLocation = 3 
		rsCheckQTY.Open CheckQTY, connstr
		
		if Not rsCheckQTY.Eof then
			varSelisih = cdbl(rsCheckQTY("selisih"))
		Else
		   varSelisih = 0		   
		End if

		if (varSelisih  - cdbl(QTY) ) >=  0 then
		    checkRate = "select rate, duedate from agreementdetail " & _
						"where " & _
						"fcba = '" & Session("fcBA") & "' and " & _
						"agreementcode = '" & FCCODE & "' and " & _
						"lineno = '" & AGREEMENTLINENO &"'"
						
						
			Set rsCheckRate = Server.CreateObject("ADODB.Recordset")
			rsCheckRate.CursorType = 3
			rsCheckRate.CursorLocation = 3 
			rsCheckRate.Open checkRate, connstr

			if Not rsCheckRate.Eof then
				varRate = cdbl(rsCheckRate("rate"))
				varDueDate = rsCheckRate("duedate")
			Else
				varRate = 0
			End if		

			rsCheckRate.Close
			set rsCheckRate = Nothing		

			
            strCheckDate = "select to_number(to_date('" &  WIPDATE & "','MM/DD/YYYY') - to_date('" &  varDueDate & "','MM/DD/YYYY')) as RemainDate from dual "
     			
			Set rsCheckDate = Server.CreateObject("ADODB.Recordset")
			rsCheckDate.CursorType = 3
			rsCheckDate.CursorLocation = 3 
			rsCheckDate.Open strCheckDate, connstr

			if Not rsCheckDate.Eof then
				varRemain 	= 	cdbl(rsCheckDate("RemainDate"))
			Else
				varRemain 	= 	0
			End if	
			
			rsCheckDate.Close
			set rsCheckDate = Nothing		

            if varRemain <= 0 then		
				q	=	" Insert into WORKINPROGRESS (" & _
						" AGREEMENTCODE" & _
						" , WIPDATE" & _
						" , AGREEMENTLINENO" & _
						" , QTY" & _
						" , POSTEDFLAG" & _
						" , PROGRESSPAYMENTDATE" & _
						" , FCENTRY" & _
						" , FCEDIT" & _
						" , FCIP" & _
						" , FCBA" & _
						" , WIPAMT" & _
						" , PAID" & _
						" , NETBALANCE" & _
						" , REFERENCE" & _
						" , REMARKS" & _
						" , LASTUPDATE" & _
						" , LASTTIME" & _
						" )" & _
						" values (" & _
						" '" & FCCODE & "'" & _
						" , to_date('" & WIPDATE & "','mm/dd/yyyy')" & _
						" , " & AGREEMENTLINENO & "" & _
						" , " & QTY & "" & _
						" , 0" & _
						" , to_date('" & WIPDATE & "','mm/dd/yyyy')" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						" , '" & Session("fcBA") & "'" & _
						" , " & QTY * varRate & "" & _
						" , 0" & _
						" , 0" & _
						" , '" & REFERENCE & "'" & _
						" , '" & REMARKS & "'" & _
						" , sysdate" & _
						" , To_Char(sysdate,'HH24:MI')" & _
						" )"

					'Response.Write q	
					connStr.execute(q)
					cmdView	= "True"
					varError = ""
			else
				cmdView	= "True"
				varError = "Tanggal WIP harus lebih kecil dari Tanggal DUEDATE pada Contract Aggreement "  
			end if
	    else
			cmdView	= "True"
			varError = "Total Quantity yang anda masukan melebihi dari kontrak yang disepakati "  
        end if
		
		rsCheckQTY.Close
		Set rsCheckQTY = Nothing
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	   
		'Validasi Qty
		CheckQTY = "SELECT   fcba, agreementcode, lineno, SUM (totalagreement) totalagreement," & _
					" SUM (totalwip) totalwip,   SUM (totalagreement) - SUM (totalwip) selisih" & _
					" FROM ((SELECT   fcba, agreementcode, lineno, SUM (qty) totalagreement, 0 totalwip" & _
					" FROM agreementdetail GROUP BY fcba, agreementcode, lineno) UNION (" &_
					" SELECT   fcba, agreementcode, agreementlineno lineno, 0 totalagreement, SUM (qty) totalwip" & _
					" FROM workinprogress GROUP BY fcba, agreementcode, agreementlineno)) dataall" & _
					" WHERE fcba = '" & Session("fcBA") & "' AND agreementcode = '" & FCCODE & "'" & _
					" AND lineno = '" & AGREEMENTLINENO & "' GROUP BY fcba, agreementcode, lineno"
		
		Set rsCheckQTY = Server.CreateObject("ADODB.Recordset")
		rsCheckQTY.CursorType = 3
		rsCheckQTY.CursorLocation = 3 
		rsCheckQTY.Open CheckQTY, connstr
		
		if Not rsCheckQTY.Eof then
			varSelisih = cdbl(rsCheckQTY("selisih"))
		Else
			varSelisih = 0
		End if
		
		CheckWIPQTY = 		" SELECT nvl(SUM(qty),0) qty" & _
							" FROM workinprogress" & _
							" WHERE fcba = '" & Session("fcBA") & "' AND agreementcode = '" & FCCODE & "'" & _
							" AND AGREEMENTLINENO = '" & AGREEMENTLINENO & "' AND WIPDATE = TO_DATE('" & WIPDATE & "','mm/dd/yyyy')"
		
		Set rsCheckWIPQTY = Server.CreateObject("ADODB.Recordset")
		rsCheckWIPQTY.CursorType = 3
		rsCheckWIPQTY.CursorLocation = 3 
		rsCheckWIPQTY.Open CheckWIPQTY, connstr
		
		if Not rsCheckWIPQTY.Eof then
			varWIPQTY = cdbl(rsCheckWIPQTY("qty"))
		Else
			varWIPQTY = 0
		End if

		if ((varSelisih + varWIPQTY) - cdbl(QTY) ) >=  0 then
		
			q	=	" DELETE FROM WORKINPROGRESS  WHERE" & _
					" WIPDATE = TO_DATE('" & WIPDATE & "','mm/dd/yyyy')" & _
					" AND AGREEMENTCODE = '" & FCCODE & "'" & _
					" AND AGREEMENTLINENO = '" & AGREEMENTLINENO & "'" & _
					" AND FCBA = '" & Session("fcBA") & "'"
			ConnStr.execute(q)
		
			checkRate = "select rate from agreementdetail " & _
				"where " & _
				"fcba = '" & Session("fcBA") & "' and " & _
				"agreementcode = '" & FCCODE & "' and " & _
				"lineno = '" & AGREEMENTLINENO &"'"
						
						
			Set rsCheckRate = Server.CreateObject("ADODB.Recordset")
			rsCheckRate.CursorType = 3
			rsCheckRate.CursorLocation = 3 
			rsCheckRate.Open checkRate, connstr

			if Not rsCheckRate.Eof then
				varRate = cdbl(rsCheckRate("rate"))
			Else
				varRate = 0
			End if	
		
			rsCheckRate.Close
			set rsCheckRate = Nothing
			
			
		
				q	=	" Insert into WORKINPROGRESS (" & _
						" AGREEMENTCODE" & _
						" , WIPDATE" & _
						" , AGREEMENTLINENO" & _
						" , QTY" & _
						" , POSTEDFLAG" & _
						" , PROGRESSPAYMENTDATE" & _
						" , FCENTRY" & _
						" , FCEDIT" & _
						" , FCIP" & _
						" , FCBA" & _
						" , WIPAMT" & _
						" , PAID" & _
						" , NETBALANCE" & _
						" , REFERENCE" & _
						" , REMARKS" & _
						" , LASTUPDATE" & _
						" , LASTTIME" & _
						" )" & _
						" values (" & _
						" '" & FCCODE & "'" & _
						" , to_date('" & WIPDATE & "','mm/dd/yyyy')" & _
						" , " & AGREEMENTLINENO & "" & _
						" , " & QTY & "" & _
						" , 0" & _
						" , to_date('" & WIPDATE & "','mm/dd/yyyy')" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						" , '" & Session("fcBA") & "'" & _
						" , " & QTY * varRate & "" & _
						" , 0" & _
						" , 0" & _
						" , '" & REFERENCE & "'" & _
						" , '" & REMARKS & "'" & _
						" , sysdate" & _
						" , To_Char(sysdate,'HH24:MI')" & _
						" )"
						
					connStr.execute(q)
					cmdView	= "True"
					varError = ""
					
	    else
			cmdView	= "True"
			varError = "Total Quantity yang anda masukan melebihi dari kontrak yang disepakati "  
        end if
		
		rsCheckQTY.Close
		Set rsCheckQTY = Nothing
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

If Len(Trim(fcCode)) > 0 Then

	Dim rsCekAvailableCode
	set rsCekAvailableCode = server.CreateObject("Adodb.recordset")

	q = " SELECT DISTINCT AGREEMENTCODE FROM AGREEMENTDETAIL" & _
		" where agreementcode = '" & fcCode & "'" & _
		"       and fcBA = '" & session("fcBA") & "'"
		
	rsCekAvailableCode.Open q, connstr
	
	If Not rsCekAvailableCode.EOF then
		FCCODE = rsCekAvailableCode("AGREEMENTCODE")
	Else
		Response.Write "<div class = ""error"">" & "Kontrak yang di maksud, tidak ada dalam Master Kontrak. Cek Kembali Kontrak anda" & "</div>"
		CmdView = "False"
	End if

	rsCekAvailableCode.Close
	set rsCekAvailableCode = Nothing

End If

txtTable	= Request("txtTable")
txtModule	= Request("txtModule")
txtSearch	= Request("txtSearch")
txtField	= Request("txtField")
txtJenis	= request("txtJenis")

If Len(Trim(varError)) > 0 Then
	response.Write  "<div class = ""error"">" & varError & "</div>"
End If

if err.number <> 0 then
   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Error Number : " & err.number & " because : " & err.description & "</font>"
End if
 '-------------ISI Payment Form
 %>
 <form name="frmEntry" method="post" action="WIP.asp ">
	<table class="data" width="100%">
		<tr>
			<td  align="left" width="205">Agreement Code</td>
			<td align="left">
				<input type="text" maxlength="30" name="fcCODE" size="30" value="<%=fcCODE%>">
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="PrintWIP.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
				<!-- HADIR -->
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Contract"
				dataSubmodule 	= "Work in Progress"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td>
		</tr>
    </table>  
 </form>
<%
	If cmdView = "True" And Len(Trim(fcCode)) > 0 Then
		Dim rsWC,intRecordsPerPage,cColor
		set rsWC = server.CreateObject("Adodb.recordset")
		'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
		'Field Date Mesti to_char(WIPDATE,'mm/dd/yyyy') as WIPDATE
		q = " SELECT AGREEMENTCODE" & _
			"        , TO_CHAR(WIPDATE,'MM/DD/YYYY') AS WIPDATE" & _
			"        , AGREEMENTLINENO" & _
			"        , QTY" & _
			"        , REFERENCE" & _
			"        , REMARKS" & _
			"        , FCENTRY" & _
			"        , FCEDIT" & _
			"        , FCIP" & _
			"        , FCBA " & _
			" FROM WORKINPROGRESS" & _
			" WHERE FCBA = '" & session("fcBA") & "'" & _
			"       AND AGREEMENTCODE = '" & fcCode & "'" & _
			" Order By AGREEMENTCODE, AGREEMENTLINENO, WIPDATE"
			
		'response.Write q
		Set rsWC = Server.CreateObject("ADODB.Recordset")
		rsWC.CursorType = 3
		rsWC.CursorLocation = 3
		rsWC.Open q, connstr
%>
	<form action="WIP.asp " method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
		<%
		'Yang tidak bisa di edit adalah PK Table
		%>
		<%IF txtJenis="Edit" Then%>
			<tr>
				<td align="left" width="205">AGREEMENTLINENO</td>
				<td align="left">
					<input type="text" size="15" name="AGREEMENTLINENO" value="<%=AGREEMENTLINENO%>"  readonly="True" style="background:#999999">
				</td>
			</tr>
			<tr>
				<td align="left">WIPDATE</td>
				<td align="left">
					<input class="datepicker" type="text" size="15" name="WIPDATE" value="<%=WIPDATE%>" readonly="True" style="background:#999999">
				</td>
			</tr>
		<%else%>
 		    <tr>
				<td align="left">AGREEMENTLINENO</td>
				<td align="left">
					<input type="text" size="15" name="AGREEMENTLINENO" value="<%=AGREEMENTLINENO%>">
				</td>
			</tr>
			<tr>
				<td align="left">WIPDATE</td>
				<td align="left">
					<input class="datepicker" type="text" size="15" name="WIPDATE" value="<%=WIPDATE%>">
				</td>
			</tr>
		<%End if%>
			<tr>
				<td align="left">QUANTITY</td>
				<td align="left">
					<input type="text" size="15" name="QTY" value="<%=QTY%>">
				</td>
			</tr>
			<tr>
				<td align="left">REFERENCE</td>
				<td align="left">
					<input type="text" size="30" maxlength = "30" name="REFERENCE" value="<%=REFERENCE%>">
				</td>
			</tr>
			<tr>
				<td align="left">REMARKS</td>
				<td align="left">
					<textarea  name="REMARKS" maxlength = "100" cols="50" rows="3"><%=REMARKS%></textarea>
				</td>
			</tr>
			<tr>
				<td colspan="2" align="left">
					<input type="hidden" name="fcCode" value="<%=fcCode%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtJenis" value="Edit">
						<input Type="Submit"  name="Add" value="Update">
					<%Else%>
						<input type="hidden" name="txtJenis" value="New">
						<input Type="Submit"  name="Add" value="Add">
					<%End if%>
					<input Type="Reset"  name="Reset" value="Reset">
					<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				</td>
			</tr>
		</table>

	<table class="data">
		<br/>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
			<TD align="left">NO</TD>
			<TD colspan="2">&nbsp;</TD>
			<TD align="left">AGREEMENT CODE</TD>	
			<TD align="left">WIP DATE</TD>	
			<TD align="left">AGREEMENT LINE NO</TD>	
			<TD align="left">QUANTITY</TD>
			<TD align="left">REFERENCE</TD>	
			<TD align="left">REMARKS</TD>	
		</tr>   
   <%
   if Not rsWC.eof then
      nCountofField=rsWC.fields.count-1
	  no=1
	  While Not rsWC.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="WIP.asp?fcCode=<%=rsWC("AGREEMENTCODE")%>&AGREEMENTLINENO=<%=rsWC("AGREEMENTLINENO")%>&wipdate=<%=rsWC("WIPDATE")%>&QTY=<%=rsWC("QTY")%>&REFERENCE=<%=rsWC("REFERENCE")%>&REMARKS=<%=rsWC("REMARKS")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsWC(1)%>" border="0"></a></td><td>
			<a href="WIP.asp?fcCode=<%=rsWC("AGREEMENTCODE")%>&AGREEMENTLINENO=<%=rsWC("AGREEMENTLINENO")%>&wipdate=<%=rsWC("WIPDATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsWC(1)%>" border="0"   onClick="return confirm('Are you sure to delete');"></a>
			</td>
		    <%for i = 0 to nCountofField-4%>
			    <td align="left" title="<%=rsWC(i).Name%>" valign="top"><%=rsWC(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsWC.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsWC.Close
   set rsWC = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
