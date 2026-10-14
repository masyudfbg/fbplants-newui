<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Piezometer Transaction</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
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
Dim FCENTRY, FCEDIT, FCIP, FCBA
dim txt_PiezoCode,Perioddate,rscek

Dim TRANSACTIONTYPE, PIEZOID,TRANSDATE,WATERLEVEL, STATUS,WEEK, adarecord, tglrecord
Dim xCONDITION
Dim cmdView
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
txttable=request("txttable")

	txt_PiezoCode = request("txt_PiezoCode")
	Perioddate = request("Perioddate")
	
	TRANSDATE = request("TRANSDATE")
    WEEK = request("WEEK")
	WATERLEVEL = request("WATERLEVEL")
	xCONDITION = request("xCONDITION")
	
	'Add by Akbar


	STATUS = request("STATUS")

	
    ' Add by Akbar
	if WATERLEVEL <> "" and TRANSDATE <> "" then
	   dim rswaterlevel
	   Set rswaterlevel = Server.CreateObject("ADODB.Recordset")
	   rswaterlevel.Open " select waterlevelstatus_api.GET_Status(" & WATERLEVEL & ") status , weekmaster_api.GET_weekid(trunc(to_Date('" & TRANSDATE & "','MM/DD/YYYY'))) thisweek from dual" , connstr
	   IF Not rswaterlevel.eof then
		 ' request("STATUS") =  rswaterlevel("status")
		  STATUS = rswaterlevel("status")
		  WEEK = rswaterlevel("thisweek")
	   End if
	   rswaterlevel.Close				
	   set rswaterlevel = Nothing

	end if 
	
	
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
	
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), TRANSDATE)
		If Cek_IsDailyClosing = "FALSE" then
		  connstr.BeginTrans
			q = " delete from PiezoTransaction_tab" & _
			    " where fcba = '" & Session("fcBA") & "'" & _
				"       and weekid = '" & WEEK  & "'" & _
				"       and PIEZOID = '" & txt_PiezoCode & "'"
			   connStr.execute(q)
			   'response.Write(q)
			   'response.Flush()
			if err.number  = 0 then
						connStr.CommitTrans
			 end if   
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), TRANSDATE)
		If Cek_IsDailyClosing = "FALSE" then
		   connstr.BeginTrans
		    q = " select weekid,TRANSDATE from PiezoTransaction_tab" & _
			    " where fcba = '" & Session("fcBA") & "'" & _
				"       and weekid = '" & WEEK  & "'" & _
				"       and PIEZOID = '" & txt_PiezoCode & "'"
		    Set rscek = Server.CreateObject("ADODB.Recordset") 		
		    rscek.CursorType = 3
			rscek.CursorLocation = 3
			rscek.Open q, connstr
			if not rscek.eof then
				 adarecord = rscek("weekid")
				 tglrecord = rscek("TRANSDATE")
			else
				 adarecord = ""
			End if
			rscek.Close
			set rscek = Nothing
			
			if adarecord = "" then
				  
				  q=" Insert into PiezoTransaction_tab (PIEZOID,WEEKID, TRANSDATE, WATERLEVEL, STATUS, CONDITION, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
					" values ('" &  txt_PiezoCode & "',weekmaster_api.GET_weekid(trunc(to_Date('" & TRANSDATE & "','MM/DD/YYYY'))) ,trunc(to_Date('" & TRANSDATE & "','MM/DD/YYYY')) ," &  WATERLEVEL & ",waterlevelstatus_api.GET_Status(" & WATERLEVEL & "),'" & xCONDITION & "'," & _
					"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
					'response.Write(q)
					'response.Flush()
					'potsting date di ubah menjadi to_date dan TRANSDATE dihilangkan di update
					connStr.execute(q)
					if err.number  = 0 then
					   connStr.CommitTrans
					end if
			else
				err.number = 1001
				err.description = "WEEK  """ & WEEK & """, sudah di ada datanya pada tanggal """ & tglrecord & """, sehingga tidak bisa melakukan insert data, " &  chr(13) & " Silahkan konsultasikan dengan administrator anda"
			End If		
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), TRANSDATE)
		If Cek_IsDailyClosing = "FALSE" then
		    connstr.BeginTrans
				q = " update PiezoTransaction_tab" & _
                    " set TRANSDATE= trunc(to_Date('" & TRANSDATE & "','MM/DD/YYYY'))" & _
					"     , WATERLEVEL= " & WATERLEVEL & ""  & _ 
			        "     , STATUS= waterlevelstatus_api.GET_Status(" & WATERLEVEL & ")"  & _ 
					"     , CONDITION= '" & xCONDITION & "' " & _ 
					"     , fcedit = '" & Session("CodeName") & "'" & _
			        "     , LASTUPDATE = sysdate " & _
			        "     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			        " where fcba = '" & Session("fcBA") & "'" & _
					"       and weekid = '" & WEEK  & "'" & _
			        "       and PIEZOID = '" & txt_PiezoCode & "'"
				'response.Write(q)
				'response.Flush()
				connStr.execute(q)
				if err.number  = 0 then
						connStr.CommitTrans
			    end if
				
				'RESPONSE.Write "SQLnsddsdya : " & Q
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & 
else
    cmdView	= request("cmdView")
End if

 if err.number <> 0 then
   connStr.RollbackTrans   
   response.Write  "<div class = ""error"">" & err.description & "</div>"
 End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if len(trim(Perioddate)) <> 0 then
   Perioddate = Request("Perioddate")
else
   Perioddate = textDate(date)
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")


 '-------------ISI Payment Form
 %>

 <form name="frmEntry" method="post" action="PiezometerTransaction.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
		<TR> 
		<td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE PIZEOMETER TRANSACTION</strong></font></div></td>
		</TR>

		<tr><td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		
		</td></tr>
		<TR>
			<td align="left">Piezo Code</td>
			<% 	
			
			Dim rsPiezoid
			q = "select P.PIEZOID from  PIEZO_TAB P  where P.FCBA = '" & Session("fcBA") & "' order by  P.PIEZOID "
			Set rsPiezoid = Server.CreateObject("ADODB.Recordset")
			rsPiezoid.CursorType = 3
			rsPiezoid.CursorLocation = 3
			'response.write q
			rsPiezoid.Open q, connstr
			if not rsPiezoid.eof then%>
     	    <td align="left"> <select name="txt_PiezoCode" >         
			<option value=""></option> 
			<%			Do While Not rsPiezoid.eof%>
			<%if rsPiezoid("PIEZOID")=request("txt_PiezoCode") Then%>			
				<option value="<%=rsPiezoid("PIEZOID")%>" selected><%=rsPiezoid("PIEZOID")%></option> 
			<%Else%>
				<option value="<%=rsPiezoid("PIEZOID")%>"><%=rsPiezoid("PIEZOID")%></option> 
			<%End if%>
			<%rsPiezoid.movenext
			loop
			End if
			rsPiezoid.Close
			set rsPiezoid = Nothing
			q=""
			%>
			</select></td>
		</TR>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintPiezoTransaction_tabtest.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_PiezoCode=<%=REQUEST("txt_PiezoCode")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		<!-- HADIR -->
		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Weighbridge"
		dataSubmodule 	= "Weighbridge Ticket"
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td>
		</tr>
    </table>  
 </form>

<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select PIEZOID,WEEKID, TRANSDATE, WATERLEVEL, STATUS, CONDITION, FCBA, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME" & _ 
     " from PiezoTransaction_tab where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(Perioddate)) > 0 then
		q = q & " and to_char(TRANSDATE,'MMYYYY') = to_char(to_date('" & Perioddate & "','MM/DD/YYYY'),'MMYYYY')"
	end if
	q = q & " Order by TRANSDATE,PIEZOID"
   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>

   <form action="PiezometerTransaction.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">
		
		<tr>
			<td align="left">Piezo Code</td>
			<td align="left">
				<input type="text" size="15" maxlength="15" name="PIEZOID" value="<%=request("txt_PiezoCode")%>" readonly="True" style="background:#999999">
			</td>
		</tr>
		<tr>
		<td align="left">Transaction Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="TRANSDATE" size="10" value="<%=TRANSDATE%>">
			<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
		<td align="left">Week</td>
		   <td align="left"><input type="text" size="15" name="WEEK" value="<%=WEEK%>" readonly="True" style="background:#999999"> </td>
		</tr>
		<tr>
		<td align="left">Water Level</td><td align="left"><input type="text" size="5" maxlength="5" name="WATERLEVEL" value="<%=WATERLEVEL%>"><font color="#FF0000"> Cm*</font></td>
		</tr>
		<tr>
		<td align="left">Status</td>
		   <td align="left"><input type="text" size="25" name="STATUS" value="<%=STATUS%>" readonly="True" style="background:#999999"> </td>
		</tr>
		<tr>
		<td align="left">Condition Piezometer</td>
		<td align="left">      
		<select name="xCONDITION">
	    <%cmdCallComboNonFcba "CONDITIONPIEZO",xCONDITION,"",""%>
        </select><font color="#FF0000">&nbsp;*</font>
        </td>
		</td>
		</tr>

	    <tr><td colspan="2" align="left">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
		  <input type="hidden" name="txt_PiezoCode" value="<%=txt_PiezoCode%>">
	    
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
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
   </form>
   <!--#include file="validation_show.Asp"--> 

<%
  
	JumRec=1
	intRecordsPerPage = 20
	Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
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
        <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;"> 
          <TD align="left">NO</TD>	
		  <TD colspan="2">&nbsp;</TD>
          <TD align="left">PIEZOID</TD>
		  <TD align="left">WEEK</TD>
          <TD align="left">TRANS DATE</TD>
          <TD align="left">WATER LEVEL</TD>
          <TD align="left">STATUS</TD>
		  <TD align="left">CONDITION PIZEOMETER</TD>
        </tr>
        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-6
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
        <tr bgcolor="<%=cColor%>"> 
          <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">
			  <a href="PiezometerTransaction.asp?txttable=<%=request("txttable")%>&TRANSDATE=<%=rsGAD("TRANSDATE")%>&WEEK=<%=rsGAD("WEEKID")%>&STATUS=<%=rsGAD("STATUS")%>&txt_PiezoCode=<%=rsGAD("PIEZOID")%>&WATERLEVEL=<%=rsGAD("WATERLEVEL")%>&xCONDITION=<%=rsGAD("CONDITION")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a> 
			  </td>
			  <td valign="middle" width="20"> <a href="PiezometerTransaction.asp?txttable=<%=request("txttable")%>&txt_PiezoCode=<%=rsGAD("PIEZOID")%>&WEEK=<%=rsGAD("WEEKID")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0"></a> 
			  </td>
		
          <%for i = 0 to nCountofField -1%>
          <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
          <%next%>
        </tr>
        <%
		  rsGAD.MoveNext
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
   </div>
      </TABLE>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
