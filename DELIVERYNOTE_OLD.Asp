<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Delivery Note</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
<script language="JavaScript" src="date-picker.js"></script> 
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
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
 <%
	'-------------awal-----------------------


Dim FCENTRY, FCEDIT, FCIP, FCBA
dim txt_NOTECODE,Perioddate

Dim DNOTE_DATE, DELIVERYNOTECODE, CONTRACTCODE, GATEPASS, WBT_CHITNO, STORAGECODE, QUANTITY, FFA, DIRT, MOISTURE, TRANSPORTCO, DRIVERNAME, ICNO

DIM SEALNO_A,SEALNO_B,SEALNO_C,SEALNO_D,SEALNO_E,SEALNO_F,SEALNO_G,SEALNO_H,SEALNO_I,SEALNO_J
Dim SEALNO
Dim cmdView,rsDN
txt_NOTECODE=request("txt_NOTECODE")
Perioddate=request("Perioddate")
txttable=request("txttable")
DNOTE_DATE=request("DNOTE_DATE")
DELIVERYNOTECODE=request("DELIVERYNOTECODE")
CONTRACTCODE=request("CONTRACTCODE")
GATEPASS=request("GATEPASS")
WBT_CHITNO=request("WBT_CHITNO")
STORAGECODE=request("STORAGECODE")
QUANTITY=request("QUANTITY")
FFA=request("FFA")
DIRT=request("DIRT")
MOISTURE=request("MOISTURE")
TRANSPORTCO=request("TRANSPORTCO")
DRIVERNAME=request("DRIVERNAME")
ICNO=request("ICNO")
	SEALNO_A=request("SEALNO_A")
	SEALNO_B=request("SEALNO_B")
	SEALNO_C=request("SEALNO_C")
	SEALNO_D=request("SEALNO_D")
	SEALNO_E=request("SEALNO_E")
	SEALNO_F=request("SEALNO_F")
	SEALNO_G=request("SEALNO_G")
	SEALNO_H=request("SEALNO_H")
	SEALNO_I=request("SEALNO_I")
	SEALNO_J=request("SEALNO_J")

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
       q="Delete from DELIVERYNOTE where DELIVERYNOTECODE='" & DELIVERYNOTECODE & "' And DNOTE_DATE=to_Date('" & DNOTE_DATE & "','MM/DD/YYYY') And  fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
	  q="Insert into DELIVERYNOTE (DNOTE_DATE, DELIVERYNOTECODE, CONTRACTCODE, GATEPASS, WBT_CHITNO, STORAGECODE, QUANTITY, FFA, DIRT, MOISTURE, TRANSPORTCO, DRIVERNAME, ICNO, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values (to_Date('" & DNOTE_DATE & "','MM/DD/YYYY'),'" &  DELIVERYNOTECODE & "','" &  CONTRACTCODE & "','" & GATEPASS & "','" &  WBT_CHITNO & "'," & _
		"'" &  STORAGECODE & "','" &  QUANTITY & "','" &  FFA & "','" &  DIRT & "','" &  MOISTURE & "','" & TRANSPORTCO & "','" &  DRIVERNAME & "','" &  ICNO & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		'response.Flush()
		connStr.execute(q)
		cmdView	= "True"
		  IF REQUEST(SEALNO_A)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_A & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
			cmdView	= "True"
		  ELSEIF REQUEST(SEALNO_B)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_B & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
			cmdView	= "True"
		  ELSEIF REQUEST(SEALNO_C)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_C & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
			cmdView	= "True"
		  ELSEIF REQUEST(SEALNO_D)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_D & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
			cmdView	= "True"
		  ELSEIF REQUEST(SEALNO_E)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_E & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
			cmdView	= "True"
		  ELSEIF REQUEST(SEALNO_F)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_F & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
			cmdView	= "True"
		  ELSEIF REQUEST(SEALNO_G)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_G & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
			cmdView	= "True"
		  ELSEIF REQUEST(SEALNO_H)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_H & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
			cmdView	= "True"
		  ELSEIF REQUEST(SEALNO_I)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_I & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
			cmdView	= "True"
		  ELSEIF REQUEST(SEALNO_J)<> "" THEN
		  q="Insert into DELIVERYNOTESEAL (DELIVERYNOTECODE, SEALNO,FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('" &  DELIVERYNOTECODE & "','" &  SEALNO_J & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			'response.Write(q)
			'response.Flush()
			connStr.execute(q)
			cmdView	= "True"
		   END IF
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update DELIVERYNOTE set TRANSPORTCO ='" & TRANSPORTCO & "',DRIVERNAME='" & DRIVERNAME & "',ICNO ='" & ICNO & "',WBT_CHITNO='" & WBT_CHITNO & "',FFA ='" & FFA & "'," & _
		"DIRT ='" & DIRT & "',MOISTURE ='" & MOISTURE & "',STORAGECODE ='" & STORAGECODE & "',GATEPASS='" & GATEPASS & "',QUANTITY='" & QUANTITY & "'," & _
		"CONTRACTCODE ='" & CONTRACTCODE & "' where DELIVERYNOTECODE='" & DELIVERYNOTECODE & "' And DNOTE_DATE=to_Date('" & DNOTE_DATE & "','MM/DD/YYYY') And  fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		connStr.execute("commit")
		'RESPONSE.Write "SQLnsddsdya : " & Q
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if Request("Perioddate")<>"" then
   Perioddate   =Request("Perioddate")
else
   Perioddate   =textDate(date)
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if

 '-------------ISI Payment Form
 %>
 <form name="frmEntry" method="post" action="DELIVERYNOTE.asp">
	<table class="data" width="100%" cellpadding="0" cellspacing="0">
		<TR> 
		<td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE DELIVERY NOTE</strong></font></div></td>
		</TR>

		<tr><td align="left">Date</td>
		<td align="left"><input type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		<a href="javascript:show_calendar('frmEntry.Perioddate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>
		<TR>
		<td align="left">Delivery Note Code</td><td align="left"><input type="text" size="15" maxlength="15" name="txt_NOTECODE" value="<%=txt_NOTECODE%>"><font color="#FF0000">*</font></td>
		</TR>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintDeliveryNote.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_NOTECODE=<%=txt_NOTECODE%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select DNOTE_DATE, DELIVERYNOTECODE, CONTRACTCODE, GATEPASS, WBT_CHITNO, STORAGECODE, QUANTITY, FFA, " & _
     "DIRT, MOISTURE, TRANSPORTCO, DRIVERNAME, ICNO, FCENTRY, FCEDIT, FCIP, FCBA from DELIVERYNOTE where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(request("Perioddate"))) > 0 then
		q = q & " and to_char(DNOTE_DATE,'mm/yyyy')=to_char(to_date('" & request("Perioddate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("txt_NOTECODE"))) > 0 then
		q = q & " and DELIVERYNOTECODE='" & request("txt_NOTECODE") & "'"
	end if
	q = q
   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>

   <form action="DELIVERYNOTE.asp" method="post" name="eForm">
   <table class="data" width="100%" cellpadding="0" cellspacing="0">
		<tr>
		<td align="left">Date</td>
			<td align="left"><input type="text" maxlength="10" name="DNOTE_DATE" size="10" value="<%=DNOTE_DATE%>">
			<a href="javascript:show_calendar('eForm.DNOTE_DATE');" onMouseOver="window.status='Date Picker';return true;" 
			onmouseout="window.status='';return true;">
			<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
			</td>
		<td align="left">Seal No</td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_A" value="<%=SEALNO_A%>"></td>		
		</tr>
		<TR>
		<td align="left">Delivery Note Code</td><td align="left"><input type="text" size="15" maxlength="15" name="DELIVERYNOTECODE" value="<%=DELIVERYNOTECODE%>"><font color="#FF0000">*</font></td>
		<td align="left"></td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_B" value="<%=SEALNO_B%>"></td>		
		</TR>	
		<TR>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Contract Code</td>
			<td align="left"> <select name="CONTRACTCODE">         
			<%cmdCallComboSetup "DELIVERYCONTRACT","CONTRACTCODE","",""%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left">Contract Code</td>
			<td align="left"><select name="CONTRACTCODE">         
			<%cmdCallComboSetup "DELIVERYCONTRACT","CONTRACTCODE","",""%>
			</select> <font color="#FF0000">*</font></td>
		
		<%End IF%>
		<td align="left"></td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_C" value="<%=SEALNO_C%>"></td>					
		</tr>
		<tr>
		<td align="left">Gate Pass</td><td align="left"><input type="text" size="30" maxlength="50" name="GATEPASS" value="<%=GATEPASS%>"></td>
		<td align="left"></td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_D" value="<%=SEALNO_D%>"></td>					
		</tr>
		<TR>
		<td align="left">Weighbridge ChitNo</td><td align="left"><input type="text" size="15" maxlength="15" name="WBT_CHITNO" value="<%=WBT_CHITNO%>"><font color="#FF0000">*</font></td>
		<td align="left"></td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_E" value="<%=SEALNO_E%>"></td>							
		</tr>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>		
            <td align="left"><font face="Verdana" size="1">Storage Code</font></td>
            <td align="left"> <select name="STORAGECODE">
                <%cmdCallComboSetup "PRODUCTSTORAGE",request("STORAGECODE"),"",""%>
              </select> <font color="#FF0000">*</font> </td>
		<%else%>
				<td align="left"><font face="Verdana" size="1">Storage Code</font></td>
				<td align="left"> <select name="STORAGECODE">
				<%cmdCallComboSetup "PRODUCTSTORAGE","STORAGECODE","",""%>
				</select> <font color="#FF0000">*</font> </td>
		<%end if%>
		<td align="left"></td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_F" value="<%=SEALNO_F%>"></td>					
		</tr>
		<tr>
		<td align="left">Quantity (KG)</td><td align="left"><input type="text" size="30" maxlength="50" name="QUANTITY" value="<%=QUANTITY%>"></td>	
		<td align="left"></td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_H" value="<%=SEALNO_H%>"></td>					
		</tr>
		<tr>
		<td align="left">FFA</td><td align="left"><input type="text" size="15" name="FFA" value="<%=FFA%>"></td>
		<td align="left"></td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_G" value="<%=SEALNO_G%>"></td>						
		</tr>
		<tr>
		<td align="left">Dirt</td><td align="left"><input type="text" size="15" name="DIRT" value="<%=DIRT%>"></td>
		<td align="left"></td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_I" value="<%=SEALNO_I%>"></td>					
		</tr>
		<tr>
		<td align="left">Moisture</td><td align="left"><input type="text" size="15" name="MOISTURE" value="<%=MOISTURE%>"></td>
		<td align="left"></td><td align="left"><input type="text" size="15" maxlength="15" name="SEALNO_J" value="<%=SEALNO_J%>"></td>					
		</tr>
		<tr>
		<td align="left">Transport CO</td><td align="left"><input type="text" size="30" maxlength="50" name="TRANSPORTCO" value="<%=TRANSPORTCO%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Driver Name</td><td align="left"><input type="text" size="30"  maxlength="50" name="DRIVERNAME" value="<%=DRIVERNAME%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">IC No</td><td align="left"><input type="text" size="20" maxlength="20" name="ICNO" value="<%=ICNO%>"></td>
		</tr>
	    <TR><td colspan="2">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
		  <input type="hidden" name="txt_NOTECODE" value="<%=txt_NOTECODE%>">

		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset" style="background-color: #298CF7; color: #F7C610; font-weight:
		  bold; font-size: 8pt; font-family:Verdana" name="Reset" value="Reset">
		  <input Type="Button" style="background-color: #298CF7; color: #F7C610; font-weight: 
		  bold; font-size: 8pt; font-family:Verdana" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
   </form>
  <!--#include file="validation_show.Asp"--> 
<%
   JumRec=1
   intRecordsPerPage = 20
	
	Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 20
	
	rsGAD.PageSize = intRecordsPerPage
	intNumberOfPages = 0
	intNumberOfPages =  rsGAD.PageCount
	intRecords = rsGAD.RecordCount
	'Set our AbsolutePage property.
	if not rsGAD.EOF then
		If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
		  rsGAD.AbsolutePage = 1
		Else
		  rsGAD.AbsolutePage = Request("iPageNo")
		End If  
	end if
	if Request("iPageNo") = "" then
		nPage = 1
	Else
		nPage = Request("iPageNo")
	End if

   rsGAD.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount   
   %>
   <table width="737" class="data" >
	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE</TD>	
	<TD align="left">DELIVERY NOTE CODE</TD>	
	<TD align="left">CONTRACT CODE</TD>	
	<TD align="left">GATE PASS</TD>	
	<TD align="left">WBT CHIT NO</TD>	
	<TD align="left">STORAGE CODE</TD>	
	<TD align="left">QUATITY</TD>	
	<TD align="left">FFA</TD>	
	<TD align="left">DIRT</TD>	
	<TD align="left">MOISTURE</TD>	
	<TD align="left">TRASPORT CO</TD>	
	<TD align="left">DRIVER</TD>	
	<TD align="left">ICNO</TD>	
	</tr>
   
        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-5
	  no=1
	  While Not rsGAD.EOF and intRecordsPerPage > 0 
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">

			<a href="DELIVERYNOTE.asp?DELIVERYNOTECODE=<%=rsGAD("DELIVERYNOTECODE")%>&txttable=<%=request("txttable")%>&DNOTE_DATE=<%=rsGAD("DNOTE_DATE")%>&TRANSPORTCO=<%=rsGAD("TRANSPORTCO")%>&DRIVERNAME=<%=rsGAD("DRIVERNAME")%>&ICNO=<%=rsGAD("ICNO")%>&WBT_CHITNO=<%=rsGAD("WBT_CHITNO")%>&FFA=<%=rsGAD("FFA")%>&DIRT=<%=rsGAD("DIRT")%>&MOISTURE=<%=rsGAD("MOISTURE")%>&STORAGECODE=<%=rsGAD("STORAGECODE")%>&GATEPASS=<%=rsGAD("GATEPASS")%>&QUANTITY=<%=rsGAD("QUANTITY")%>&CONTRACTCODE=<%=rsGAD("CONTRACTCODE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="DELIVERYNOTE.asp?DELIVERYNOTECODE=<%=rsGAD("DELIVERYNOTECODE")%>&txttable=<%=request("txttable")%>&DNOTE_DATE=<%=rsGAD("DNOTE_DATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
	<tr>
	<td colspan="10">
	   <%if intNumberOfPages<>1 then
		  For i = 1 to intNumberOfPages
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""DELIVERYNOTE.asp?txtModule=" &_
			  txtModule & "&iPageNo=" &  i & "&txtTable=" & request("txtTable")& "&txtField=" & request("txtField")  & _
			  "&txtSearch=" & request("txtSearch") & _
			  """>" & i & "  "
		  Next
	   End if%>
	   </a><br>
	   <%if Request("iPageNo") = "" then%>
			<%nPage = 1%>
	   <%Else%><%'End if%>  
 
  <%'-------------akhir----------------------%>
 </td>
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
			<%nPage = Request("iPageNo")%>
	   <%End if%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Page : <%=nPage%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total Records : <%=intRecords%></td>
	</tr>
   </TABLE>
<%End if%>
</html>
