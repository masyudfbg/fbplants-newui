<!--#include file="../connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
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
vMonth=trim(Request("vMonth"))
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
  'Dim rs
  Set rs = Server.CreateObject("ADODB.Recordset")

  rs.PageSize = 5
  rs.CacheSize = 5
  rs.CursorLocation = adUseClient
  if request("vJenisReport")="Group" then
     strsql1 = " SELECT  A.ACCOUNTGROUPCODE,O.GROUP_BA,A.ACCOUNTGROUP,sum(A.TM_AMT_ACTUAL)," & _
		    " sum(A.TM_HA_KG_ACTUAL), " & _
			" sum(A.TM_AMT_BUDGET),sum(A.TM_HA_KG_BUDGET),sum(A.TY_AMT_ACTUAL),sum(A.TY_HA_KG_ACTUAL), " & _
			" sum(A.TY_AMT_BUDGET),sum(A.TY_HA_KG_BUDGET) from  sac_ACCOUNTREPORT_tm_tbm a,sac_Organization o WHERE" & _
			" a.ba = o.ba and" & _
			" A.CROPTYPE='" & request("vCROPTYPE") & "'" & _
			" and A.OWNERSHIP='" & request("vOWNERSHIP") & "'" & _
			" and A.Status='" & request("vStatus") & "'" & _
			" and A.PERIOD_YEAR=" & request("vYear")
  elseif request("vJenisReport")="BA" then
     strsql1 = " SELECT  A.ACCOUNTGROUPCODE,O.BA,A.ACCOUNTGROUP,sum(A.TM_AMT_ACTUAL)," & _
		    " sum(A.TM_HA_KG_ACTUAL), " & _
			" sum(A.TM_AMT_BUDGET),sum(A.TM_HA_KG_BUDGET),sum(A.TY_AMT_ACTUAL),sum(A.TY_HA_KG_ACTUAL), " & _
			" sum(A.TY_AMT_BUDGET),sum(A.TY_HA_KG_BUDGET) from  sac_ACCOUNTREPORT_tm_tbm a,sac_Organization o WHERE" & _
			" a.ba = o.ba and" & _
			" A.CROPTYPE='" & request("vCROPTYPE") & "'" & _
			" and A.OWNERSHIP='" & request("vOWNERSHIP") & "'" & _
			" and A.Status='" & request("vStatus") & "'" & _
			" and A.PERIOD_YEAR=" & request("vYear")
  elseif request("vJenisReport")="Job" then
     strsql1 = " SELECT  A.ACCOUNTGROUPCODE,A.DIVISION,A.ACCOUNTGROUP,sum(A.TM_AMT_ACTUAL)," & _
		    " sum(A.TM_HA_KG_ACTUAL), " & _
			" sum(A.TM_AMT_BUDGET),sum(A.TM_HA_KG_BUDGET),sum(A.TY_AMT_ACTUAL),sum(A.TY_HA_KG_ACTUAL), " & _
			" sum(A.TY_AMT_BUDGET),sum(A.TY_HA_KG_BUDGET) from  sac_ACCOUNTREPORT_tm_tbm a,sac_Organization o WHERE" & _
			" a.ba = o.ba and" & _
			" A.CROPTYPE='" & request("vCROPTYPE") & "'" & _
			" and A.OWNERSHIP='" & request("vOWNERSHIP") & "'" & _
			" and A.Status='" & request("vStatus") & "'" & _
			" and A.PERIOD_YEAR=" & request("vYear")

  End if			

  if  request("vMonth") <>"ALL" then
	  strsql2 = " and A.PERIOD_MONTH=" & request("vMonth") 
  else
	  strsql2 = " "
  end if
  if  request("vBA") <>"ALL" then
	  strsql3 = "  and A.BA='" &request("vBA") & "'" 
  else
	  strsql3 = " "
  end if

  if  request("vOrganization") <>"ALL"  and request("vJenisReport")="Group"  then
	  strsql3A = "  and A.BA in (Select BA from sac_Organization where GROUP_BA='" & trim(request("vOrganization")) & "')" 
  elseif  request("vOrganization") <>"ALL"  and request("vJenisReport")="BA"  then
  	  strsql3A = "  and A.BA in (Select BA from sac_Organization where GROUP_BA='" & trim(request("vOrganization")) & "')" 
  elseif  request("vOrganization") <>"ALL"  and request("vJenisReport")="Job"  then
  	  strsql3A = "  and A.BA in (Select BA from sac_Organization where BA='" & trim(request("vOrganization")) & "')" 
  end if

  
  if  request("vPlantingDate") <>"ALL" then
	  strsql4 = " and A.PlantingDate='" &request("vPlantingDate") & "'"
  else
	  strsql4 = " "
  end if
  if  request("vDivision") <>"ALL" then
	  strsql5 = " and A.Division='" & request("vDivision") & "'" 
  else
	  strsql5 = " "
  end if
  'response.Write "Jenis : " & request("vJenisReport")
  
  if request("vJenisReport")="Group" Then
	 strsql6  =" and AccountCode in (select fccode from job where " & _
			   " JOBGROUPCODE ='" & request("v_Jobgroup") & "')" 
	 strsql9  =" group by A.ACCOUNTGROUPCODE,O.GROUP_BA,A.ACCOUNTGROUP" 
  elseif request("vJenisReport")="BA" Then
	 strsql6  =" and AccountCode in (select fccode from job where " & _
			   " JOBGROUPCODE ='" & request("v_Jobgroup") & "')" 
	 strsql9  =" group by A.ACCOUNTGROUPCODE,O.BA,A.ACCOUNTGROUP" 
  elseif request("vJenisReport")="Job" Then
	 strsql6  =" and AccountCode in (select fccode from job where " & _
			   " JOBGROUPCODE ='" & request("v_Jobgroup") & "')" 
	 strsql9  =" group  BY A.ACCOUNTGROUPCODE,A.DIVISION,A.ACCOUNTGROUP" 
  end if			
  
   
  
  strsql7  =" "
  
  strsql8  =" "
	  			
	  
  strsql = strsql1 + strsql2 + strsql3 + strsql3A + strsql4 + strsql5 + strsql6 +  strsql7 + strsql8 + strsql9
  
  'response.Write(strsql)

  'response.flush()	
	  
  rs.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly

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

  '---- Khusus untuk Judul
  if trim(request("vOwnership")) ="1" then
     vTextOwnership = "Inti"
  elseif trim(request("vOwnership")) ="2" then 
     vTextOwnership = "PMBN"
  elseif trim(request("vOwnership")) ="3" then
     vTextOwnership = "PMU"
  elseif trim(request("vOwnership")) ="4" then	 
     vTextOwnership = "Plasma"
  elseif trim(request("vOwnership")) ="5" then	 
     vTextOwnership = "Plasma Pre"
  elseif trim(request("vOwnership")) ="6" then	 
     vTextOwnership = "KKPA"
  elseif trim(request("vOwnership")) ="7" then	 
     vTextOwnership = "KKPA Pre"
  elseif trim(request("vOwnership")) ="8" then	 
     vTextOwnership = "TKD"
  end if	 

  if Trim(Request("vCROPTYPE"))="1" then
     vTextCroptype="Oil Palm"
  elseif Trim(Request("vCROPTYPE"))="2" then
     vTextCroptype="Rubber"
  elseif Trim(Request("vCROPTYPE"))="3" then
     vTextCroptype="CoCoa"
  End if
  '---- end Judul
  %>
   
 	<table border="1">
<TR>
<TD COLSPAN="11">

       <%If request.Form("vJenisReport")="Group" Then%>
	ANALISA BIAYA <%=request("vStatus")%> Per-Group Pekerjaan<br>
	<%else%>
	ANALISA BIAYA <%=request("vStatus")%> Per-Pekerjaan<br>
	<%End if%>
	Group : <%=Request("vOrganization")%>, Kebun : <%=Request("vBA")%> <br>
	Period Year :<%=Request("vYear")%> Month : <%=Request("vMOnth")%><br>
	Crop :<%=vTextCroptype%> Ownership : <%=vTextOwnership%>
	<br>
	Afdeling <%=Request("vDivision")%>, Planting Date: <%=Request("vPlantingDate")%>

</TD>
</TR>

	<tr>
	<td colspan="3">&nbsp;</td>
	<td colspan="4" valign="middle" align="center">Bulan ini (TM)</td>
	<td colspan="4" valign="middle" align="center">s.d Bulan ini (TY)</td>
	</tr>
  <tr bgcolor="#F7F7F7" valign="middle" align="center">
  <td><font face="Verdana" size="1">No</td>
  <%if request("vjenisreport")="Source" then%>
     <td><font face="Verdana" size="1">Group</td>
     <td><font face="Verdana" size="1">Pekerjaan</td>
  <%else%>
     <td><font face="Verdana" size="1">Group</td>
     <td><font face="Verdana" size="1">Pekerjaan Group</td>
  <%End if%>
  <td><font face="Verdana" size="1">Actual</td>
  <td><font face="Verdana" size="1">Rp/Ha</td>
  <td><font face="Verdana" size="1">Budget</td>
  <td><font face="Verdana" size="1">Rp/Ha</td>
  <td><font face="Verdana" size="1">Actual</td>
  <td><font face="Verdana" size="1">Rp/Ha</td>
  <td><font face="Verdana" size="1">Budget</td>
  <td><font face="Verdana" size="1">Rp/Ha</td>
  </tr>

  <%
  If request("vJenisReport")="Group" Then
     Turunan = "BA"
  elseif request("vJenisReport") = "BA" Then
     Turunan ="Job"
  elseif request("vJenisReport") = "Job" Then
     Turunan ="CostElement"
  end if
  'response.Write "hasil " &  trim(request("vJenisReport")) = "BA"
  While not Rs.EOF and intRecordsPerPage > 0%>
          </font>
          <tr>
		  <td Valign="Top" bgcolor="#C0FBBD"><font face="Verdana" size="1"><%=JumRec%>
          </font>
          <font size="1" face="Arial Narrow" color="#FFFFFF">
          <% for i = 1 to howmanyfields%>
            </font>
			<%if i >=3 then%>
            <td valign="top" bgcolor="#C0FBBD" align="right"> <font face="Verdana" size="1"><%=FormatNumber(rs(i),2)%></font></td>
			<%else
				if HeadAfdeling = rs(i) and i = 1 then%>
				  <td valign="top" bgcolor="#C0FBBD"> <font face="Verdana" size="1">&nbsp;</font></td>
				<%elseif i = 2  then%>
		<td valign="top" bgcolor="#C0FBBD" align="left"><font face="Verdana" size="1">
		<%if request("vJenisReport") = "Group" then%>
<A href="P_byJobcode.Asp?vJenisReport=<%=Turunan%>&vYear=<%=request("vYear")%>&vMOnth=<%=request("vMonth")%>&vBA=<%=vBA%>&vStatus=<%=request("vStatus")%>&vOwnership=<%=request("vOwnership")%>&vCropType=<%=request("vCropType")%>&vDivision=<%=request("vDivision")%>&vPlantingDate=<%=request("vPlantingDate")%>&v_JobGroup=<%=rs(0)%>&vOrganization=<%=rs(1)%>" target="_blank"><%=rs(i)%>
		<%elseif trim(request("vJenisReport")) = "BA" then%>
<A href="P_byJobcode.Asp?vJenisReport=<%=Turunan%>&vOrganization=<%=rs(1)%>&vYear=<%=request("vYear")%>&vMOnth=<%=request("vMonth")%>&vBA=<%=vBA%>&vStatus=<%=request("vStatus")%>&vOwnership=<%=request("vOwnership")%>&vCropType=<%=request("vCropType")%>&vDivision=<%=request("vDivision")%>&vPlantingDate=<%=request("vPlantingDate")%>&v_JobGroup=<%=rs(0)%>" target="_blank"><%=rs(i)%>
		<%elseif trim(request("vJenisReport")) = "Job" then%>
<%=rs(i)%>
		<%End if%>
				  
				<%else%>
				  <td valign="top" bgcolor="#C0FBBD"> <font face="Verdana" size="1"><%=rs(i)%></font></td>
				<%End if%>			
			<%end if%>
             <font size="1" face="Arial Narrow" color="#FFFFFF">
          <% Next %>
          </tr>
     <%		  
	 HeadAfdeling = rs.fields(1)
	 HeadTTanam   = rs.fields(2)
	 Headactivity     = rs.fields(3)
     rs.movenext
	 
     JumRec=JumRec+1
	 intRecordsPerPage = intRecordsPerPage - 1
  Wend
End If
rs.close
set rs = nothing%>
<tr>
<td colspan="11">
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