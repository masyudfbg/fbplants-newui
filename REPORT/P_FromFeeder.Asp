<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3%>
<!--#include file="p_top.Asp"-->
  <table width="100%" border="0" cellspacing="0"  cellpadding="0">
  </table>
  <font style="font-size:12px">
  </font>
<table border=0 Cellspacing =1 Cellpadding=1 width="100%">
    <tr><td align="left">
<!-- cmdTransaction to different of Transaction or table master 
     True = Transaction ; False = table Master
-->
<%cmdTransaction=False%>
<%
vJenisReport=trim(Request("vJenisReport"))
vyear=trim(Request("vyear"))
IF LEN(trim(Request("vMonth"))) =1 THEN
   v_Month="0" & trim(Request("vMonth"))
ELSE
   v_Month=trim(Request("vMonth"))
END IF
vBA=trim(Request("vBA"))
vOwnership=trim(Request("vOwnership"))
vStatus=trim(Request("vStatus"))
vCroptype=trim(Request("vCroptype"))
vPlantingDate=trim(Request("vPlantingDate"))
vCroptype=trim(Request("vCroptype"))
%>
<%If Jenis="tblUser" then%>
<span class="style2"></span></td>
<%End if%>
</tr>
  <center><td>
    </td>
  </tr>
  </table>
  <table border=0 Cellspacing =1 Cellpadding=1 width="100%">
  <!--Special to tblMapping for selected Modul -->
<br>  
<%
  Dim rs
  Set rs = Server.CreateObject("ADODB.Recordset")

  rs.PageSize = 5
  rs.CacheSize = 5
  rs.CursorLocation = adUseClient

	STRSQL1= "SELECT JOB.JOBDESCRIPTION,SOURCETYPE," & _
			 "SUM(TVALUE)  FROM COSTBOOK ,JOB " & _
			 " WHERE TO_CHAR(tdate,'mm-yyyy')='" & v_month &"-" & request("vYear") & "' AND "
	STRSQLA = "	TargetCode IN (SELECT f.FIELDCODE " & _
			"	   FROM FIELD F, FIELDMISC FM,FIELDCROP fc " & _
			"		   WHERE f.FIELDCODE = fm.FIELDCODE " & _
			"		   AND   f.FIELDCODE = fc.FIELDCODE	" & _	 
			"		   AND   f.fieldcode NOT IN (SELECT FieldCode FROM FIELDMISC WHERE active='N') " & _
			"		   AND   FM.OWNERSHIP = '" & request("vOwnership") & "'" & _
			"		   AND   FM.STATUS = '" & request("vStatus") & "'" & _
			"		   AND   fc.CROPTYPE = '" & request("vCropType") & "'" 
	STRSQLD	=	"	) AND"
	STRSQL3="	COSTBOOK.JOBCODE = JOB.JOBCODE AND"			   
	STRSQL4="	JOB.Jobcode  in (select Jobcode from job where jobgroupcode= '" & REQUEST("v_JobGroup") & "')" 
	if  request("vPlantingDate") <>"ALL" then
	  strsqlB = " and TO_CHAR(fc.PLANTINGDATE,'YYYY')='" &request("vPlantingDate") & "'"
	else
	  strsqlB = " "
	end if

	if  request("vOrganization") <>"ALL" then  ' Devision di set ke variable Orgnization
	  strsqlC = " AND   f.DIVISION ='" & request("vOrganization") & "'" 
	else
	  strsqlC = " "
	end if

    strsql5= " group by JOB.JOBDESCRIPTION,SOURCETYPE" 

    strsql2 = strsqlA & strsqlB & strsqlC & strsqlD
	strsql = strsql1 + strsql2 + strsql3 + strsql4 + strsql5

	'response.Write strsql
    Dim rsJob
    Set rsJob = Server.CreateObject("ADODB.Recordset")	
	
	strJob= "Select DESCRIPTION from jobgroup where GROUPCODE='" & REQUEST("v_JobGroup") & "'"
	rsJob.Open strJob,connStr,adOpenForwardOnly, adLockReadOnly
	if not rsJob.eof then
	   v_DescJob = rsJob("DESCRIPTION")
	else
	   v_DescJob = ""
	end if
	rsJob.close
	set rsJob= Nothing

	Set connFeeder=Server.Createobject("ADODB.Connection")
	connFeeder.open  "Provider=MSDAORA.1;Data Source=" & "BUJF"& request("vBA")& ";User ID=TDSUITE;" & "Password=oracle9i"

rs.Open strsql, connFeeder, adOpenForwardOnly, adLockReadOnly
	  
If rs.EOF Then%>
  <tr><td colspan="3">Data is empty</tr></td>
<%Else
  rs.movefirst
  
  howmanyfields=rs.fields.count-1
  
  JumRec=1
  intRecordsPerPage = 10000
  Rs.PageSize = intRecordsPerPage
  intNumberOfPages = 0
  intNumberOfPages =  Rs.PageCount
  intRecords = Rs.RecordCount
  ' Set our AbsolutePage property.
  If isEmpty(Request("iPageNo")) Then
	Rs.AbsolutePage = 1
  Else
	Rs.AbsolutePage = Request("iPageNo")
  End If  
  %>
  
    <tr>
      <td width="150" height="130" colspan="3"><img src="Images/Logoawal.gif" width="150" height="130"></td>
      <td align="LEFT" width="580"><font style="font-size:12px"><B>
	  ASIAN AGRI GROUP<BR>
	  ANALISA DETAIL<br>
      Kebun : <%=request("vBA")%>, <%=request("vOrganization")%> <br>
      Period Year :<%=request("vYear")%> Month : <%=textMonth(request("vMOnth"))%></font></div><Br>
	  Activity Group : <%=v_DescJob%></td>
    </tr>
  <table bgcolor="#F7F7F7">
  <tr bgcolor="#F7F7F7" valign="middle" align="center">
  <td><font face="Verdana" size="1">No</td>
  <td><font face="Verdana" size="1">Activity</td>
  <td><font face="Verdana" size="1">Cost Element</td>
  <td><font face="Verdana" size="1">Cost This Month</td>
  </tr>

  <%
  While not Rs.EOF and intRecordsPerPage > 0%>
          </font>
          <tr>
		  <td Valign="Top" bgcolor="#C0FBBD"><font face="Verdana" size="1"><%=JumRec%>
          </font>
          <font size="1" face="Arial Narrow" color="#FFFFFF">
          <% for i = 0 to howmanyfields%>
            </font>
			<%if i = 2 then%>
                 <td valign="top" bgcolor="#C0FBBD" align="right"> <font face="Verdana" size="1">
				 <%=RubahAngka(rs(i),2)%></font></td>
			<%Else%>
                 <%if trim(rs("JOBDESCRIPTION")) = trim(HeadJob) and i = 0 then%>
				      <td valign="top" bgcolor="#C0FBBD" align="Left"> <font face="Verdana" size="1">&nbsp;
					  </font></td>
				 <%Else%>
				      <td valign="top" bgcolor="#C0FBBD" align="Left"> <font face="Verdana" size="1">
					  <%=rs(i)%></font></td>
				 <%End if%>
			<%End if%>
          <% Next %>
          </tr>
	 <%	  
	 HeadJob = rs.fields("JOBDESCRIPTION")
	 HeadTTanam   = rs.fields(2)
	 'Headactivity     = rs.fields(3)
     rs.movenext
	 
     JumRec=JumRec+1
	 intRecordsPerPage = intRecordsPerPage - 1
  Wend
End If
rs.close
set rs = nothing%>
<tr>
<td colspan="6">
   <%if intNumberOfPages<>1 then
      For i = 1 to intNumberOfPages
          Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""AccountReport.asp?&iPageNo=" &  i & "&vJenisReport=" & request("vJenisReport")& "&fcValue_Organization=" & request("fcValue_Organization") & """>" & i & "  "
      Next
   End if%>
   </a><br>
   <%if Request("iPageNo") = "" then%>
        <%nPage = 1%>
   <%Else%>
        <%nPage = Request("iPageNo")%>
   <%End if%>
   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Page : <%=nPage%>
   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total Records : <%=intRecords%></td>
</tr>
</Table></center>
   <font size="-4">
   <%'response.Write(strsql)%></font><%
%>
</font>

<!--#include file="p_bottom.Asp"-->