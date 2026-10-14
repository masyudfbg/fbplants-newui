<% ValidateSession %>
<!--#include file="../ConndB1.inc"-->
<!--#include file="../txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<!--#include file="top1.Asp"-->
<%vJenisReport=trim(Request.Form("vJenisReport"))%>
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
<%If Jenis="tblUser" then%>
<span class="style2"></span></td>
<%End if%>
</tr>
  <center>
  <tr align="Left" valign="middle"><td valign="middle" align="center"><form name="form1" method="post">
      <label><font size="2"  color="#0000FF" ></font>
        <select name="vYear">
          <option value="2006" selected="selected">2006</option>
          <option value="2007">2007</option>
        </select>
        </label>
      <label><font size="2" color="#0000FF" ></font>
        <select name="vMonth">
		  <%if request.Form("vMonth")="1" then%>
             <option value="1" selected="selected">January</option>
		  <%else%>
             <option value="1">January</option>
		  <%End if%>
		  <%if request.Form("vMonth")="2" then%>
             <option value="2" selected="selected">February</option>
		  <%else%>
             <option value="2">February</option>
		  <%End if%>
		  <%if request.Form("vMonth")="3" then%>
             <option value="3" selected="selected">March</option>
		  <%else%>
             <option value="3">March</option>
		  <%End if%>
		  <%if request.Form("vMonth")="4" then%>
             <option value="4" selected="selected">April</option>
		  <%else%>
             <option value="4">April</option>
		  <%End if%>
		  <%if request.Form("vMonth")="5" then%>
             <option value="5" selected="selected">May</option>
		  <%else%>
             <option value="5">May</option>
		  <%End if%>
		  <%if request.Form("vMonth")="6" then%>
             <option value="6" selected="selected">June</option>
		  <%else%>
             <option value="6">June</option>
		  <%End if%>
		  <%if request.Form("vMonth")="7" then%>
             <option value="7" selected="selected">July</option>
		  <%else%>
             <option value="7">July</option>
		  <%End if%>
		  <%if request.Form("vMonth")="8" then%>
             <option value="8" selected="selected">August</option>
		  <%else%>
             <option value="8">August</option>
		  <%End if%>
		  <%if request.Form("vMonth")="9" then%>
             <option value="9" selected="selected">September</option>
		  <%else%>
             <option value="9">September</option>
		  <%End if%>
		  <%if request.Form("vMonth")="10" then%>
             <option value="10" selected="selected">October</option>
		  <%else%>
             <option value="10">October</option>
		  <%End if%>
		  <%if request.Form("vMonth")="11" then%>
             <option value="11" selected="selected">November</option>
		  <%else%>
             <option value="11">November</option>
		  <%End if%>
		  <%if request.Form("vMonth")="12" then%>
             <option value="12" selected="selected">December</option>
		  <%else%>
             <option value="12">December</option>
		  <%End if%>
		     <option value="ALL">::ALL::</option>
        </select>
        </label>		
      <label><font size="2" color="#0000FF" ></font>
        <select name="vBA">
	     <%IF request.Form("vBA")="ALL" THEN%>
			<option value="ALL" selected="selected">::ALL::</option>
		 <%ELSE%>
			<option value="ALL">::ALL::</option>
		 <%END IF%>

	     <%IF request.Form("vBA")="PGS" THEN%>
			<option value="PGS" selected="selected">PGS</option>
		 <%ELSE%>
			<option value="PGS">PGS</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PGD" THEN%>
			<option value="PGD" selected="selected">PGD</option>
		 <%ELSE%>
			<option value="PGD">PGD</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PTD" THEN%>
			<option value="PTD" selected="selected">PTD</option>
		 <%ELSE%>
			<option value="PTD">PTD</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PTS" THEN%>
			<option value="PTS" selected="selected">PTS</option>
		 <%ELSE%>
			<option value="PTS">PTS</option>
		 <%END IF%>

	     <%IF request.Form("vBA")="PNS" THEN%>
			<option value="PNS" selected="selected">PNS</option>
		 <%ELSE%>
			<option value="PNS">PNS</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PND" THEN%>
			<option value="PND" selected="selected">PND</option>
		 <%ELSE%>
			<option value="PND">PND</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PAN" THEN%>
			<option value="PAN" selected="selected">PAN</option>
		 <%ELSE%>
			<option value="PAN">PAN</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PTP" THEN%>
			<option value="PTP" selected="selected">PTP</option>
		 <%ELSE%>
			<option value="PTP">PTP</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PBD" THEN%>
			<option value="PBD" selected="selected">PBD</option>
		 <%ELSE%>
			<option value="PBD">PBD</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PBS" THEN%>
			<option value="PBS" selected="selected">PBS</option>
		 <%ELSE%>
			<option value="PBS">PBS</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PUS" THEN%>
			<option value="PUS" selected="selected">PUS</option>
		 <%ELSE%>
			<option value="PUS">PUS</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PUD" THEN%>
			<option value="PUD" selected="selected">PUD</option>
		 <%ELSE%>
			<option value="PUD">PUD</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PSG" THEN%>
			<option value="PSG" selected="selected">PSG</option>
		 <%ELSE%>
			<option value="PSG">PSG</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PPN" THEN%>
			<option value="PPN" selected="selected">PPN</option>
		 <%ELSE%>
			<option value="PPN">PPN</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PTZ" THEN%>
			<option value="PTZ" selected="selected">PTZ</option>
		 <%ELSE%>
			<option value="PTZ">PTZ</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PMB" THEN%>
			<option value="PMB" selected="selected">PMB</option>
		 <%ELSE%>
			<option value="PMB">PMB</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PTR" THEN%>
			<option value="PTR" selected="selected">PTR</option>
		 <%ELSE%>
			<option value="PTR">PTR</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="PTU" THEN%>
			<option value="PTU" selected="selected">PTU</option>
		 <%ELSE%>
			<option value="PTU">PTU</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="CDS" THEN%>
			<option value="CDS" selected="selected">CDS</option>
		 <%ELSE%>
			<option value="CDS">CDS</option>
		 <%END IF%>
	     <%IF request.Form("vBA")="CMP" THEN%>
			<option value="CMP" selected="selected">CMP</option>
		 <%ELSE%>
			<option value="CMP">CMP</option>
		 <%END IF%>
        </select>
        </label>
      <label><font size="2"  color="#0000FF" ></font>
        <select name="vCROPTYPE">
		<%if Request.Form("vCROPTYPE")="1" or Request.Form("vCROPTYPE")="" Then%>
          <option value="1" selected>Oil Palm</option>
          <option value="2">Rubber</option>
          <option value="3">CoCoa</option>
		 <%elseif Request.Form("vCROPTYPE")="2" Then%>
          <option value="1">Oil Palm</option>
          <option value="2" selected="selected">Rubber</option>
          <option value="3">CoCoa</option>
		 <%Elseif Request.Form("vCROPTYPE")="3" then%> 
          <option value="1">Oil Palm</option>
          <option value="2">Rubber</option>
          <option value="3" selected="selected">CoCoa</option>
		 <%End if%> 
        </select></label>
      <label><font size="2"  color="#0000FF" ></font>
        <select name="vJenisReport">
		 <%if request.Form("vJenisReport")="Station" Then%>
          <option value="Station"  Selected>Station Group</option>
          <option value="Machine">Station per Activity</option>
          <option value="StationMachine">Station Per Machine</option>
          <option value="MachineJob">Machine/Activity/Cost Element</option>
          <option value="JobMachine">Activity/Machine</option>
		 <%elseif request.Form("vJenisReport")="Machine" Then%>
          <option value="Station">Station Group</option>
          <option value="StationMachine">Station Per Machine</option>
          <option value="Machine" Selected>Station per Activity</option>
          <option value="MachineJob">Machine/Activity/Cost Element</option>
          <option value="JobMachine">Activity/Machine</option>
		 <%elseif request.Form("vJenisReport")="StationMachine" Then%>
          <option value="Station">Station Group</option>
          <option value="StationMachine"  Selected>Station Per Machine</option>
          <option value="Machine">Station per Activity</option>
          <option value="MachineJob">Machine/Activity/Cost Element</option>
          <option value="JobMachine">Activity/Machine</option>
		 <%elseif request.Form("vJenisReport")="MachineJob" Then%>
          <option value="Station">Station Group</option>
          <option value="StationMachine">Station Per Machine</option>
          <option value="Machine">Station per Activity</option>
          <option value="MachineJob" selected>Machine/Activity/Cost Element</option>
          <option value="JobMachine">Activity/Machine</option>
		 <%elseif request.Form("vJenisReport")="JobMachine" Then%>
          <option value="Station">Station Group</option>
          <option value="StationMachine">Station Per Machine</option>
          <option value="Machine">Station per Activity</option>
          <option value="MachineJob">Machine/Activity/Cost Element</option>
          <option value="JobMachine"  selected>Activity/Machine</option>
		 <%else %>
          <option value="Station" Selected>Station Group</option>
          <option value="StationMachine">Station Per Machine</option>
          <option value="Machine">Station per Activity</option>
          <option value="MachineJob">Machine/Activity/Cost Element</option>
          <option value="JobMachine">Activity/Machine</option>
		  <%End if%>
        </select>
        </label>
		<input type="submit" value="View" >
    </form>
    </td>
	  </tr>
</TABLE>	  
<%


  'Dim rs
  Set rs = Server.CreateObject("ADODB.Recordset")

  Dim rsBudget
  Set rsBudget = Server.CreateObject("ADODB.Recordset")

  rs.PageSize = 5
  rs.CacheSize = 5
  rs.CursorLocation = adUseClient
  IF request.Form("vYear")="" THEN
	strsql = "SELECT STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION," & _
		     "ACCOUNTCODE,SOURCETYPE,sum(TM_AMT_ACTUAL),max(TM_FFB_ACTUAL), " & _
			" SUM(TM_AMT_BUDGET),max(TM_FFB_BUDGET),sum(TY_AMT_ACTUAL),sum(TY_FFB_ACTUAL), " & _
			" max(TY_AMT_BUDGET),max(TY_FFB_BUDGET) from SAC_ACCOUNTREPORT_MILL  WHERE " & _
			" CROPTYPE='" & "1" & "' AND " & _
			" BA='" & "xxx" & "' AND " & _
			" PERIOD_YEAR=" & "2006" & _
			" group by STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION,ACCOUNTCODE,ACCOUNTCODE_DESCRIPTION,SOURCETYPE"
  ELSE			
      if request.Form("vJenisReport")="Station" Then
	          strsql1 = "SELECT STATIONCODE,STATIONCODE_DESCRIPTION," & _
		"sum(TM_AMT_ACTUAL) as TM_AMT_ACTUAL,sum(TM_AMT_ACTUAL)/(max(TM_CPO_ACTUAL)+max(TM_KER_ACTUAL)) as TM_RP_PALM_ACTUAL," & _
		"sum(TM_AMT_ACTUAL)/max(TM_FFB_ACTUAL) as TM_RP_FFB_ACTUAL," & _
		"sum(TM_AMT_BUDGET) as TM_AMT_BUDGET,sum(TM_AMT_BUDGET)/(max(TM_CPO_BUDGET)+max(TM_KER_BUDGET)) as TM_RP_PALM_BUDGET," & _
		"sum(TM_AMT_BUDGET)/max(TM_FFB_BUDGET) as TM_RP_FFB_BUDGET," & _
		"sum(TY_AMT_ACTUAL) as TY_AMT_ACTUAL,sum(TY_AMT_ACTUAL)/(max(TY_CPO_ACTUAL)+max(Ty_KER_ACTUAL)) as Ty_RP_PALM_ACTUAL," & _
		"sum(TY_AMT_ACTUAL)/max(TY_FFB_ACTUAL) as TY_RP_FFB_ACTUAL," & _
		"sum(TY_AMT_BUDGET) as TY_AMT_BUDGET,sum(TY_AMT_BUDGET)/(max(TY_CPO_BUDGET)+max(Ty_KER_BUDGET)) as Ty_RP_PALM_BUDGET," & _
		"sum(TY_AMT_BUDGET)/max(TY_FFB_BUDGET) as TY_RP_FFB_BUDGET," & _
		"sum(AMT_ANNUAL) as AMT_ANNUAL,sum(AMT_ANNUAL)/(max(CPO_ANNUAL)+max(KER_ANNUAL)) as RP_PALM_ANNUAL," & _
		"sum(AMT_ANNUAL)/max(FFB_ANNUAL) as RP_FFB_ANNUAL " & _
		" from SAC_ACCOUNTREPORT_MILL  WHERE " & _
		" CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
		" and PERIOD_YEAR=" & request.Form("vYear")
		strBudget1 = "SELECT " & _
		" max(TM_CPO_ACTUAL) AS TM_CPO_ACTUAL,max(TM_KER_ACTUAL) AS TM_KER_ACTUAL,max(TM_FFB_ACTUAL) AS TM_FFB_ACTUAL," & _
		" max(TM_CPO_BUDGET) AS TM_CPO_BUDGET,max(TM_KER_BUDGET) AS TM_KER_BUDGET,max(TM_FFB_BUDGET) AS TM_FFB_BUDGET," & _
		" max(Ty_CPO_ACTUAL) AS Ty_CPO_ACTUAL,max(Ty_KER_ACTUAL) AS Ty_KER_ACTUAL,max(Ty_FFB_ACTUAL) AS Ty_FFB_ACTUAL," & _
		" max(Ty_CPO_BUDGET) AS Ty_CPO_BUDGET,max(Ty_KER_BUDGET) AS Ty_KER_BUDGET,max(Ty_FFB_BUDGET) AS Ty_FFB_BUDGET," & _
		" max(CPO_ANNUAL) AS CPO_ANNUAL ,max(KER_ANNUAL) AS KER_ANNUAL,max(FFB_ANNUAL) AS FFB_ANNUAL" & _
		" from SAC_ACCOUNTREPORT_MILL  WHERE " & _
		" CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
		" and PERIOD_YEAR=" & request.Form("vYear")
	  elseif request.Form("vJenisReport")="Machine" Then			
	        strsql1 = "SELECT STATIONCODE,STATIONCODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION," & _
" sum(TM_AMT_ACTUAL) as TM_AMT_ACTUAL," & _
" DECODE((max(TM_CPO_ACTUAL)+max(TM_KER_ACTUAL)),0,0,sum(TM_AMT_ACTUAL)/(max(TM_CPO_ACTUAL)+max(TM_KER_ACTUAL))) as TM_RP_PALM_ACTUAL," & _
" DECODE(max(TM_FFB_ACTUAL),0,0,sum(TM_AMT_ACTUAL)/max(TM_FFB_ACTUAL)) as TM_RP_FFB_ACTUAL," & _
" sum(TM_AMT_BUDGET) as TM_AMT_BUDGET," & _
" DECODE((max(TM_CPO_BUDGET)+max(TM_KER_BUDGET)),0,0,sum(TM_AMT_BUDGET)/(max(TM_CPO_BUDGET)+max(TM_KER_BUDGET))) as TM_RP_PALM_BUDGET," & _
" DECODE(max(TM_FFB_BUDGET),0,0,sum(TM_AMT_BUDGET)/max(TM_FFB_BUDGET)) as TM_RP_FFB_BUDGET," & _
" sum(TY_AMT_ACTUAL) as TY_AMT_ACTUAL," & _
" DECODE((max(TY_CPO_ACTUAL)+max(Ty_KER_ACTUAL)),0,0,sum(TY_AMT_ACTUAL)/(max(TY_CPO_ACTUAL)+max(Ty_KER_ACTUAL))) as Ty_RP_PALM_ACTUAL," & _
" DECODE(max(TY_FFB_ACTUAL),0,0,sum(TY_AMT_ACTUAL)/max(TY_FFB_ACTUAL)) as TY_RP_FFB_ACTUAL," & _
" sum(TY_AMT_BUDGET) as TY_AMT_BUDGET," & _
" DECODE((max(TY_CPO_BUDGET)+max(Ty_KER_BUDGET)),0,0,sum(TY_AMT_BUDGET)/(max(TY_CPO_BUDGET)+max(Ty_KER_BUDGET))) as Ty_RP_PALM_BUDGET," & _
" DECODE(max(TY_FFB_BUDGET),0,0,sum(TY_AMT_BUDGET)/max(TY_FFB_BUDGET)) as TY_RP_FFB_BUDGET," & _
" sum(AMT_ANNUAL) as AMT_ANNUAL," & _
" DECODE((max(CPO_ANNUAL)+max(KER_ANNUAL)),0,0,sum(AMT_ANNUAL)/(max(CPO_ANNUAL)+max(KER_ANNUAL))) as RP_PALM_ANNUAL," & _
" DECODE(max(FFB_ANNUAL),0,0,sum(AMT_ANNUAL)/max(FFB_ANNUAL)) as RP_FFB_ANNUAL " & _
" from SAC_ACCOUNTREPORT_MILL  WHERE " & _
				" CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
				" and PERIOD_YEAR=" & request.Form("vYear")
	          strBudget1 = "SELECT " & _
		" max(TM_CPO_ACTUAL) AS TM_CPO_ACTUAL,max(TM_KER_ACTUAL) AS TM_KER_ACTUAL,max(TM_FFB_ACTUAL) AS TM_FFB_ACTUAL," & _
		" max(TM_CPO_BUDGET) AS TM_CPO_BUDGET,max(TM_KER_BUDGET) AS TM_KER_BUDGET,max(TM_FFB_BUDGET) AS TM_FFB_BUDGET," & _
		" max(Ty_CPO_ACTUAL) AS Ty_CPO_ACTUAL,max(Ty_KER_ACTUAL) AS Ty_KER_ACTUAL,max(Ty_FFB_ACTUAL) AS Ty_FFB_ACTUAL," & _
		" max(Ty_CPO_BUDGET) AS Ty_CPO_BUDGET,max(Ty_KER_BUDGET) AS Ty_KER_BUDGET,max(Ty_FFB_BUDGET) AS Ty_FFB_BUDGET," & _
		" max(CPO_ANNUAL) AS CPO_ANNUAL ,max(KER_ANNUAL) AS KER_ANNUAL,max(FFB_ANNUAL) AS FFB_ANNUAL" & _
			  " from SAC_ACCOUNTREPORT_MILL  WHERE " & _
 			  " CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
			  " and PERIOD_YEAR=" & request.Form("vYear")
	  elseif request.Form("vJenisReport")="MachineJob" Then			
	        strsql1 = "SELECT STATIONCODE,STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION,sourcetype," & _
" sum(TM_AMT_ACTUAL) as TM_AMT_ACTUAL," & _
" DECODE((max(TM_CPO_ACTUAL)+max(TM_KER_ACTUAL)),0,0,sum(TM_AMT_ACTUAL)/(max(TM_CPO_ACTUAL)+max(TM_KER_ACTUAL))) as TM_RP_PALM_ACTUAL," & _
" DECODE(max(TM_FFB_ACTUAL),0,0,sum(TM_AMT_ACTUAL)/max(TM_FFB_ACTUAL)) as TM_RP_FFB_ACTUAL," & _
" sum(TM_AMT_BUDGET) as TM_AMT_BUDGET," & _
" DECODE((max(TM_CPO_BUDGET)+max(TM_KER_BUDGET)),0,0,sum(TM_AMT_BUDGET)/(max(TM_CPO_BUDGET)+max(TM_KER_BUDGET))) as TM_RP_PALM_BUDGET," & _
" DECODE(max(TM_FFB_BUDGET),0,0,sum(TM_AMT_BUDGET)/max(TM_FFB_BUDGET)) as TM_RP_FFB_BUDGET," & _
" sum(TY_AMT_ACTUAL) as TY_AMT_ACTUAL," & _
" DECODE((max(TY_CPO_ACTUAL)+max(Ty_KER_ACTUAL)),0,0,sum(TY_AMT_ACTUAL)/(max(TY_CPO_ACTUAL)+max(Ty_KER_ACTUAL))) as Ty_RP_PALM_ACTUAL," & _
" DECODE(max(TY_FFB_ACTUAL),0,0,sum(TY_AMT_ACTUAL)/max(TY_FFB_ACTUAL)) as TY_RP_FFB_ACTUAL," & _
" sum(TY_AMT_BUDGET) as TY_AMT_BUDGET," & _
" DECODE((max(TY_CPO_BUDGET)+max(Ty_KER_BUDGET)),0,0,sum(TY_AMT_BUDGET)/(max(TY_CPO_BUDGET)+max(Ty_KER_BUDGET))) as Ty_RP_PALM_BUDGET," & _
" DECODE(max(TY_FFB_BUDGET),0,0,sum(TY_AMT_BUDGET)/max(TY_FFB_BUDGET)) as TY_RP_FFB_BUDGET," & _
" sum(AMT_ANNUAL) as AMT_ANNUAL," & _
" DECODE((max(CPO_ANNUAL)+max(KER_ANNUAL)),0,0,sum(AMT_ANNUAL)/(max(CPO_ANNUAL)+max(KER_ANNUAL))) as RP_PALM_ANNUAL," & _
" DECODE(max(FFB_ANNUAL),0,0,sum(AMT_ANNUAL)/max(FFB_ANNUAL)) as RP_FFB_ANNUAL " & _
" from SAC_ACCOUNTREPORT_MILL  WHERE " & _
				" CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
				" and PERIOD_YEAR=" & request.Form("vYear")
	          strBudget1 = "select STATIONCODE,STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION,sourcetype," & _
		" max(TM_CPO_ACTUAL) AS TM_CPO_ACTUAL,max(TM_KER_ACTUAL) AS TM_KER_ACTUAL,max(TM_FFB_ACTUAL) AS TM_FFB_ACTUAL," & _
		" max(TM_CPO_BUDGET) AS TM_CPO_BUDGET,max(TM_KER_BUDGET) AS TM_KER_BUDGET,max(TM_FFB_BUDGET) AS TM_FFB_BUDGET," & _
		" max(Ty_CPO_ACTUAL) AS Ty_CPO_ACTUAL,max(Ty_KER_ACTUAL) AS Ty_KER_ACTUAL,max(Ty_FFB_ACTUAL) AS Ty_FFB_ACTUAL," & _
		" max(Ty_CPO_BUDGET) AS Ty_CPO_BUDGET,max(Ty_KER_BUDGET) AS Ty_KER_BUDGET,max(Ty_FFB_BUDGET) AS Ty_FFB_BUDGET," & _
		" max(CPO_ANNUAL) AS CPO_ANNUAL ,max(KER_ANNUAL) AS KER_ANNUAL,max(FFB_ANNUAL) AS FFB_ANNUAL" & _
			  " from SAC_ACCOUNTREPORT_MILL  WHERE " & _
 			  " CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
			  " and PERIOD_YEAR=" & request.Form("vYear")
	  elseif request.Form("vJenisReport")="StationMachine" Then			
	        strsql1 = "SELECT STATIONCODE,STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION," & _
" sum(TM_AMT_ACTUAL) as TM_AMT_ACTUAL," & _
" DECODE((max(TM_CPO_ACTUAL)+max(TM_KER_ACTUAL)),0,0,sum(TM_AMT_ACTUAL)/(max(TM_CPO_ACTUAL)+max(TM_KER_ACTUAL))) as TM_RP_PALM_ACTUAL," & _
" DECODE(max(TM_FFB_ACTUAL),0,0,sum(TM_AMT_ACTUAL)/max(TM_FFB_ACTUAL)) as TM_RP_FFB_ACTUAL," & _
" sum(TM_AMT_BUDGET) as TM_AMT_BUDGET," & _
" DECODE((max(TM_CPO_BUDGET)+max(TM_KER_BUDGET)),0,0,sum(TM_AMT_BUDGET)/(max(TM_CPO_BUDGET)+max(TM_KER_BUDGET))) as TM_RP_PALM_BUDGET," & _
" DECODE(max(TM_FFB_BUDGET),0,0,sum(TM_AMT_BUDGET)/max(TM_FFB_BUDGET)) as TM_RP_FFB_BUDGET," & _
" sum(TY_AMT_ACTUAL) as TY_AMT_ACTUAL," & _
" DECODE((max(TY_CPO_ACTUAL)+max(Ty_KER_ACTUAL)),0,0,sum(TY_AMT_ACTUAL)/(max(TY_CPO_ACTUAL)+max(Ty_KER_ACTUAL))) as Ty_RP_PALM_ACTUAL," & _
" DECODE(max(TY_FFB_ACTUAL),0,0,sum(TY_AMT_ACTUAL)/max(TY_FFB_ACTUAL)) as TY_RP_FFB_ACTUAL," & _
" sum(TY_AMT_BUDGET) as TY_AMT_BUDGET," & _
" DECODE((max(TY_CPO_BUDGET)+max(Ty_KER_BUDGET)),0,0,sum(TY_AMT_BUDGET)/(max(TY_CPO_BUDGET)+max(Ty_KER_BUDGET))) as Ty_RP_PALM_BUDGET," & _
" DECODE(max(TY_FFB_BUDGET),0,0,sum(TY_AMT_BUDGET)/max(TY_FFB_BUDGET)) as TY_RP_FFB_BUDGET," & _
" sum(AMT_ANNUAL) as AMT_ANNUAL," & _
" DECODE((max(CPO_ANNUAL)+max(KER_ANNUAL)),0,0,sum(AMT_ANNUAL)/(max(CPO_ANNUAL)+max(KER_ANNUAL))) as RP_PALM_ANNUAL," & _
" DECODE(max(FFB_ANNUAL),0,0,sum(AMT_ANNUAL)/max(FFB_ANNUAL)) as RP_FFB_ANNUAL " & _
" from SAC_ACCOUNTREPORT_MILL  WHERE " & _
				" CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
				" and PERIOD_YEAR=" & request.Form("vYear")
	          strBudget1 = "select " & _
		" max(TM_CPO_ACTUAL) AS TM_CPO_ACTUAL,max(TM_KER_ACTUAL) AS TM_KER_ACTUAL,max(TM_FFB_ACTUAL) AS TM_FFB_ACTUAL," & _
		" max(TM_CPO_BUDGET) AS TM_CPO_BUDGET,max(TM_KER_BUDGET) AS TM_KER_BUDGET,max(TM_FFB_BUDGET) AS TM_FFB_BUDGET," & _
		" max(Ty_CPO_ACTUAL) AS Ty_CPO_ACTUAL,max(Ty_KER_ACTUAL) AS Ty_KER_ACTUAL,max(Ty_FFB_ACTUAL) AS Ty_FFB_ACTUAL," & _
		" max(Ty_CPO_BUDGET) AS Ty_CPO_BUDGET,max(Ty_KER_BUDGET) AS Ty_KER_BUDGET,max(Ty_FFB_BUDGET) AS Ty_FFB_BUDGET," & _
		" max(CPO_ANNUAL) AS CPO_ANNUAL ,max(KER_ANNUAL) AS KER_ANNUAL,max(FFB_ANNUAL) AS FFB_ANNUAL" & _
			  " from SAC_ACCOUNTREPORT_MILL  WHERE " & _
 			  " CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
			  " and PERIOD_YEAR=" & request.Form("vYear")
	  elseif request.Form("vJenisReport")="JobMachine" Then			
	        strsql1 = "SELECT STATIONCODE,STATIONCODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION,MACHINECODE_DESCRIPTION," & _
" sum(TM_AMT_ACTUAL) as TM_AMT_ACTUAL," & _
" DECODE((max(TM_CPO_ACTUAL)+max(TM_KER_ACTUAL)),0,0,sum(TM_AMT_ACTUAL)/(max(TM_CPO_ACTUAL)+max(TM_KER_ACTUAL))) as TM_RP_PALM_ACTUAL," & _
" DECODE(max(TM_FFB_ACTUAL),0,0,sum(TM_AMT_ACTUAL)/max(TM_FFB_ACTUAL)) as TM_RP_FFB_ACTUAL," & _
" sum(TM_AMT_BUDGET) as TM_AMT_BUDGET," & _
" DECODE((max(TM_CPO_BUDGET)+max(TM_KER_BUDGET)),0,0,sum(TM_AMT_BUDGET)/(max(TM_CPO_BUDGET)+max(TM_KER_BUDGET))) as TM_RP_PALM_BUDGET," & _
" DECODE(max(TM_FFB_BUDGET),0,0,sum(TM_AMT_BUDGET)/max(TM_FFB_BUDGET)) as TM_RP_FFB_BUDGET," & _
" sum(TY_AMT_ACTUAL) as TY_AMT_ACTUAL," & _
" DECODE((max(TY_CPO_ACTUAL)+max(Ty_KER_ACTUAL)),0,0,sum(TY_AMT_ACTUAL)/(max(TY_CPO_ACTUAL)+max(Ty_KER_ACTUAL))) as Ty_RP_PALM_ACTUAL," & _
" DECODE(max(TY_FFB_ACTUAL),0,0,sum(TY_AMT_ACTUAL)/max(TY_FFB_ACTUAL)) as TY_RP_FFB_ACTUAL," & _
" sum(TY_AMT_BUDGET) as TY_AMT_BUDGET," & _
" DECODE((max(TY_CPO_BUDGET)+max(Ty_KER_BUDGET)),0,0,sum(TY_AMT_BUDGET)/(max(TY_CPO_BUDGET)+max(Ty_KER_BUDGET))) as Ty_RP_PALM_BUDGET," & _
" DECODE(max(TY_FFB_BUDGET),0,0,sum(TY_AMT_BUDGET)/max(TY_FFB_BUDGET)) as TY_RP_FFB_BUDGET," & _
" sum(AMT_ANNUAL) as AMT_ANNUAL," & _
" DECODE((max(CPO_ANNUAL)+max(KER_ANNUAL)),0,0,sum(AMT_ANNUAL)/(max(CPO_ANNUAL)+max(KER_ANNUAL))) as RP_PALM_ANNUAL," & _
" DECODE(max(FFB_ANNUAL),0,0,sum(AMT_ANNUAL)/max(FFB_ANNUAL)) as RP_FFB_ANNUAL " & _
" from SAC_ACCOUNTREPORT_MILL  WHERE " & _
				" CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
				" and PERIOD_YEAR=" & request.Form("vYear")
	          strBudget1 = "select " & _
		" max(TM_CPO_ACTUAL) AS TM_CPO_ACTUAL,max(TM_KER_ACTUAL) AS TM_KER_ACTUAL,max(TM_FFB_ACTUAL) AS TM_FFB_ACTUAL," & _
		" max(TM_CPO_BUDGET) AS TM_CPO_BUDGET,max(TM_KER_BUDGET) AS TM_KER_BUDGET,max(TM_FFB_BUDGET) AS TM_FFB_BUDGET," & _
		" max(Ty_CPO_ACTUAL) AS Ty_CPO_ACTUAL,max(Ty_KER_ACTUAL) AS Ty_KER_ACTUAL,max(Ty_FFB_ACTUAL) AS Ty_FFB_ACTUAL," & _
		" max(Ty_CPO_BUDGET) AS Ty_CPO_BUDGET,max(Ty_KER_BUDGET) AS Ty_KER_BUDGET,max(Ty_FFB_BUDGET) AS Ty_FFB_BUDGET," & _
		" max(CPO_ANNUAL) AS CPO_ANNUAL ,max(KER_ANNUAL) AS KER_ANNUAL,max(FFB_ANNUAL) AS FFB_ANNUAL" & _
			  " from SAC_ACCOUNTREPORT_MILL  WHERE " & _
 			  " CROPTYPE='" & request.Form("vCROPTYPE") & "'" & _
			  " and PERIOD_YEAR=" & request.Form("vYear")
	  end if			
	  
	  
	  
	  if  request.Form("vMonth") <>"ALL" then
	      strsql2 = " and PERIOD_MONTH=" & request.Form("vMonth") 
	  else
	      strsql2 = " "
	  end if
	  if  request.Form("vBA") <>"ALL" then
	      strsql3 = "  and BA='" & request.Form("vBA") & "'" 
	  else
	      strsql3 = " "
	  end if

	  if request.Form("vJenisReport")="Station" Then
	     IF (request.Form("vPlantingDate")="ALL" and request.Form("vDivision")="ALL") then
	         strsql6= " group by STATIONCODE,STATIONCODE_DESCRIPTION" 
		 else
	         strsql6= " group by STATIONCODE,STATIONCODE_DESCRIPTION" 
		 end if	 
	  elseif request.Form("vJenisReport")="Machine" Then
	     IF (request.Form("vPlantingDate")="ALL" and request.Form("vDivision")="ALL") then
	         strsql6= " group by  STATIONCODE,STATIONCODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION"
		 else
	         strsql6= " group by  STATIONCODE,STATIONCODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION"
		 end if
	  elseif request.Form("vJenisReport")="MachineJob" Then
	     IF (request.Form("vPlantingDate")="ALL" and request.Form("vDivision")="ALL") then
	         strsql6= " group by  STATIONCODE,STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION,sourcetype"
		 else
	         strsql6= " group by  STATIONCODE,STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION,sourcetype"
		 end if
	  elseif request.Form("vJenisReport")="JobMachine" Then
	     IF (request.Form("vPlantingDate")="ALL" and request.Form("vDivision")="ALL") then
	         strsql6= " group by  STATIONCODE,STATIONCODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION,MACHINECODE_DESCRIPTION"
		 else
	         strsql6= " group by  STATIONCODE,STATIONCODE_DESCRIPTION,ACCOUNTCODE_DESCRIPTION,MACHINECODE_DESCRIPTION"
		 end if
	  elseif request.Form("vJenisReport")="StationMachine" Then
	     IF (request.Form("vPlantingDate")="ALL" and request.Form("vDivision")="ALL") then
	         strsql6= " group by  STATIONCODE,STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION"
		 else
	         strsql6= " group by  STATIONCODE,STATIONCODE_DESCRIPTION,MACHINECODE_DESCRIPTION"
		 end if
	  end if	 
	  strsql = strsql1 + strsql2 + strsql3 + strsql4 + strsql5 + strsql6
	  if request.Form("vJenisReport")="MachineJob" then
	     strBudget = strBudget1 + strsql2 + strsql3 + strsql4 + strsql5+strsql6
	  else
	     strBudget = strBudget1 + strsql2 + strsql3 + strsql4 + strsql5
	  end if
	  
  END IF

if strBudget <> "" then
rsBudget.Open strBudget, connStr, adOpenForwardOnly, adLockReadOnly
if not rsBudget.eof then
	v_TM_CPO_ACTUAL=rsBudget("TM_CPO_ACTUAL")
	v_TM_KER_ACTUAL=rsBudget("TM_KER_ACTUAL")
	v_TM_PALM_ACTUAL=CDBL(v_TM_CPO_ACTUAL)+CDBL(v_TM_KER_ACTUAL)
	v_TM_FFB_ACTUAL=rsBudget("TM_FFB_ACTUAL")
	
	v_TM_CPO_BUDGET=rsBudget("TM_CPO_BUDGET")
	V_TM_KER_BUDGET=rsBudget("TM_KER_BUDGET")
	v_TM_FFB_BUDGET=rsBudget("TM_FFB_BUDGET")
	V_TM_PALM_BUDGET=CDBL(v_TM_CPO_BUDGET)+CDBL(V_TM_KER_BUDGET)
	
	v_Ty_CPO_ACTUAL=rsbudget("Ty_CPO_ACTUAL")
	v_Ty_KER_ACTUAL=rsBudget("Ty_KER_ACTUAL")
	v_Ty_FFB_ACTUAL=rsBudget("Ty_FFB_ACTUAL")
	v_Ty_PALM_ACTUAL=CDBL(v_Ty_CPO_ACTUAL)+CDBL(v_Ty_KER_ACTUAL)
	
	v_Ty_CPO_BUDGET=rsBudget("Ty_CPO_BUDGET")
	v_Ty_KER_BUDGET=rsBudget("Ty_KER_BUDGET")
	v_Ty_FFB_BUDGET=rsBudget("Ty_FFB_BUDGET")
	v_TY_PALM_BUDGET=CDBL(v_Ty_CPO_BUDGET)+CDBL(v_Ty_KER_BUDGET)
	
	v_CPO_ANNUAL =rsBudget("CPO_ANNUAL")
	v_KER_ANNUAL=rsBudget("KER_ANNUAL")
	v_FFB_ANNUAL=rsBudget("FFB_ANNUAL")
	v_PALM_ANNUAL = CDBL(v_CPO_ANNUAL)+CDBL(v_KER_ANNUAL)
else
	v_TM_CPO_ACTUAL=0
	v_TM_KER_ACTUAL=0
	v_TM_FFB_ACTUAL=0
	v_TM_CPO_BUDGET=0
	V_TM_KER_BUDGET=0
	v_TM_FFB_BUDGET=0
	v_Ty_CPO_ACTUAL=0
	v_Ty_KER_ACTUAL=0
	v_Ty_FFB_ACTUAL=0
	v_Ty_CPO_BUDGET=0
	v_Ty_KER_BUDGET=0
	v_Ty_FFB_BUDGET=0
	v_CPO_ANNUAL =0
	v_KER_ANNUAL=0
	v_FFB_ANNUAL=0
end if
RSBUDGET.CLOSE
SET RSBUGDTE = NOTHING
End if
rs.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly
If rs.EOF and len(trim(request.Form("vJenisReport"))) > 0 Then%>
  <tr><td colspan="3">Data is empty</tr></td>
<%elseIf not rs.EOF and len(trim(request.Form("vJenisReport"))) > 0 Then%>
  <table Cellspacing =1 Cellpadding=1 width="100%" style="font:Verdana, Arial, Helvetica, sans-serif;SIZE:1">
		<tr>
		  <td width="150" colspan="4"><img src="Images/Logoawal.gif" width="150" height="130"></td>
		  <td width="580" colspan="8" align="center"><font style="font-size:12px"><B>
		  MULTI INDO OIL<BR>
		  ANALISA BIAYA PABRIK<br>
		  Pabrik : <%=Request.Form("vBA")%> <br>
		  Period Year :<%=Request.Form("vYear")%> Month : <%=textMonth(trim(Request.Form("vMOnth")))%></font></div></B></td>
		  <td><a href="p_AccountReportMill.Asp?strsql=<%=strsql%>" target="_blank"><img src="Icon/Print.gif" border="0">
		  </a></td></tr>
  </table>
  <table border=0 Cellspacing =0 Cellpadding=0 width="100%" style="font:Verdana, Arial, Helvetica, sans-serif;SIZE:1">

  <%rs.movefirst
  
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
  <%if request.Form("vJenisReport")="Station" Then%>
  <!--#include file="HEADERMILL.HTML"-->  

  <%end if%>
  <%HeadAfdeling=""
  if request.Form("vJenisReport")="Station" then
     fieldnumeric = 2
  elseif request.Form("vJenisReport")="Machine" then
     fieldnumeric = 3
  elseif request.Form("vJenisReport")="MachineJob" then
     fieldnumeric =  5
  elseif request.Form("vJenisReport")="JobMachine" then
     fieldnumeric =  4
  elseif request.Form("vJenisReport")="StationMachine" then
     fieldnumeric =  3
  end if
  While not Rs.EOF and intRecordsPerPage > 0%>
          </font>
          <tr>
		     <%if trim(HeadCode)<>trim(rs(0)) then%>
			      <td style="<%=cmdstyle%>;border-left: 1 solid #0000FF" Valign="Top" bgcolor="#C0FBBD"><font face="Verdana" size="1"><%=rs(0)%></td>
             <%Else%>
			      <td style="<%=cmdstyle%>;border-left: 1 solid #0000FF" Valign="Top" bgcolor="#C0FBBD"><font face="Verdana" size="1">&nbsp;</td>
			 <%end if%>
          </font>
          <font size="1" face="Arial Narrow" color="#FFFFFF">
          <% for i = 1 to howmanyfields%>
            </font>
			<%if trim(HeadAfdeling)=trim(rs(i)) and i = 1 then%>
				  <td style="<%=cmdstyle%>" Valign="top" bgcolor="#C0FBBD"> <font face="Verdana" size="1">&nbsp;</font></td>
			<%elseif  trim(Headactivity)=trim(rs(i)) and i = 2 then%>
				  <td style="<%=cmdstyle%>" Valign="top" bgcolor="#C0FBBD"> <font face="Verdana" size="1">&nbsp;</font></td>
			<%elseif  trim(headMachine)=trim(rs(i)) and i = 3 then%>
				  <td style="<%=cmdstyle%>" Valign="top" bgcolor="#C0FBBD"> <font face="Verdana" size="1">&nbsp;</font></td>
			<%elseif i >=fieldnumeric then%>
            <td style="<%=cmdstyle%>" Valign="top" bgcolor="#C0FBBD" align="right"> <font face="Verdana" size="1"><%=FormatNumber(rs(i),2)%></font></td>
			<%else
				 if i=1 and request.Form("vJenisReport")="Station" then%>
				  <td style="<%=cmdstyle%>" Valign="top" bgcolor="#C0FBBD"> <font face="Verdana" size="1"><A href="P_byJobcode.Asp?vJenisReport=<%=vJenisReport%>&vYear=<%=request.Form("vYear")%>&vMOnth=<%=request.Form("vMonth")%>&vBA=<%=request.Form("vBA")%>&vStatus=<%=request.Form("vStatus")%>&vOwnership=<%=request.Form("vOwnership")%>&vCropType=<%=request.Form("vCropType")%>&vDivision=<%=request.Form("vDivision")%>&vPlantingDate=<%=request.Form("vPlantingDate")%>&v_Description=<%=rs(0)%>&v_Division=<%=rs(1)%>&v_PlantingDate=<%=rs(2)%>" target="_blank"><%=rs(i)%></A></font></td>
				<%elseif i=1 and request.Form("vJenisReport")="Machine" then%>
				  <td style="<%=cmdstyle%>" Valign="top" bgcolor="#C0FBBD"> <font face="Verdana" size="1"><A href="P_byJobcode.Asp?vJenisReport=<%=vJenisReport%>&vYear=<%=request.Form("vYear")%>&vMOnth=<%=request.Form("vMonth")%>&vBA=<%=request.Form("vBA")%>&vStatus=<%=request.Form("vStatus")%>&vOwnership=<%=request.Form("vOwnership")%>&vCropType=<%=request.Form("vCropType")%>&vDivision=<%=request.Form("vDivision")%>&vPlantingDate=<%=request.Form("vPlantingDate")%>&v_JobGroup=<%=rs(0)%>&v_Division=<%=rs(1)%>&v_PlantingDate=<%=rs(2)%>" target="_blank"><%=rs(i)%></A></font></td>
				<%else%>
				  <td style="<%=cmdstyle%>" Valign="top" bgcolor="#C0FBBD"> <font face="Verdana" size="1"><%=rs(i)%></font>
				  </td>
				<%End if%>			
			<%end if%>
             <font size="1" face="Arial Narrow" color="#FFFFFF">
          <% Next %>
          </tr>
     <%		  
	 HeadAfdeling = rs.fields(1)
	 headCode = rs.fields(0)
	 headActivity = rs.fields(2)
	 headMachine = rs.fields(3)

     nTotal_TM_Actual = nTotal_TM_Actual + cdbl(rs("TM_AMT_ACTUAL"))
     nTotal_TY_Actual = nTotal_TY_Actual + cdbl(rs("Ty_AMT_ACTUAL"))
     nTotal_TM_BUDGET = nTotal_TM_BUDGET + cdbl(rs("TM_AMT_BUDGET"))
     nTotal_TY_BUDGET = nTotal_TY_BUDGET + cdbl(rs("Ty_AMT_BUDGET"))
	 nTotal_AMT_Annual= nTotal_AMT_Annual + cdbl(rs("AMT_ANNUAL"))

     xTM_RP_PALM_ACTUAL = rs("TM_RP_PALM_ACTUAL")
	 xTM_FFB_ACTUAL = rs("TM_RP_FFB_ACTUAL")
	 xTM_RP_PALM_Budget = rs("TM_RP_PALM_Budget") 
	 xTM_FFB_Budget = rs("TM_RP_FFB_Budget")

     xTy_RP_PALM_ACTUAL = rs("Ty_RP_PALM_ACTUAL")
	 xTy_FFB_ACTUAL = rs("Ty_RP_FFB_ACTUAL")
	 xTy_RP_PALM_Budget = rs("Ty_RP_PALM_Budget")
	 xTy_FFB_Budget = rs("Ty_RP_FFB_Budget")

	 xRP_PALM_ANNUAL = rs("RP_PALM_ANNUAL")
	 xRP_FFB_ANNUAL = rs("RP_FFB_ANNUAL")
	 
	 
	 
     rs.movenext
	 
     JumRec=JumRec+1
	 intRecordsPerPage = intRecordsPerPage - 1
  Wend%>

	<tr bgcolor="#0099FF">
        <%IF request.Form("vJenisReport")="Station" then%>
	     <td style="<%=cmdstyle%>"colspan="2" align="center"><b><font color="#FFFFFF">T O T A L</b></td>
        <%elseIF request.Form("vJenisReport")="MachineJob" then%>
		 <td style="<%=cmdstyle%>"colspan="5" align="center"><b><font color="#FFFFFF">T O T A L</b></td>
        <%elseIF request.Form("vJenisReport")="JobMachine" then%>
		 <td style="<%=cmdstyle%>"colspan="4" align="center"><b><font color="#FFFFFF">T O T A L</b></td>
        <%else%>
	     <td style="<%=cmdstyle%>"colspan="3" align="center"><b><font color="#FFFFFF">T O T A L</b></td>
        <%End if%>
	<td style="<%=cmdstyle%>"colspan="1"><b><font color="#FFFFFF"><%=FormatNumber(nTotal_TM_Actual,2)%></b></td>
	<%if cdbl(v_TM_PALM_ACTUAL) >0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TM_Actual/CDBl(v_TM_PALM_ACTUAL),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	<%if cdbl(v_TM_FFB_ACTUAL) >0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TM_Actual/CDBl(v_TM_FFB_ACTUAL),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TM_BUDGET,2)%></b></td>
	<%if cdbl(v_TM_PALM_BUDGET) > 0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TM_BUDGET/CDBl(v_TM_PALM_BUDGET),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	<%if cdbl(v_TM_FFB_BUDGET) > 0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TM_BUDGET/CDBl(v_TM_FFB_BUDGET),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_Actual,2)%></b></td>
	<%if cdbl(v_TY_PALM_ACTUAL) > 0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_Actual/CDBl(v_TY_PALM_ACTUAL),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	<%if cdbl(v_TY_FFB_ACTUAL) > 0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_Actual/CDBl(v_TY_FFB_ACTUAL),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_BUDGET,2)%></b></td>
	<%if cdbl(v_TY_PALM_BUDGET) > 0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_BUDGET/CDBl(v_TY_PALM_BUDGET),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	<%if cdbl(v_TY_FFB_BUDGET) > 0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_TY_BUDGET/CDBl(v_TY_FFB_BUDGET),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	<td style="<%=cmdstyle%>"colspan="1"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_AMT_Annual,2)%></b></td>
	<%if cdbl(v_PALM_ANNUAL) > 0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_AMT_Annual/CDBl(v_PALM_ANNUAL),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	<%if cdbl(v_FFB_ANNUAL) > 0 then%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b><%=FormatNumber(nTotal_AMT_Annual/CDBl(v_FFB_ANNUAL),2)%></td>
	<%else%>
	<td style="<%=cmdstyle%>"colspan="1" align="Right"><font color="#FFFFFF"><b>0</td>
	<%end if%>
	</tr>

<tr>
<%
Dim colsbottom 
if request.Form("vJenisReport")="Machine" Then
   colsbottom=18
elseif request.Form("vJenisReport")="StationMachine" Then
   colsbottom=19
elseif request.Form("vJenisReport")="JobMachine" Then
   colsbottom=20
elseif request.Form("vJenisReport")="MachineJob" Then
   colsbottom=21
else
   colsbottom=17   
end if   
%>

<td style="<%=cmdstyle%>"colspan="<%=colsbottom%>">
   <%if intNumberOfPages<>1 then
      For i = 1 to intNumberOfPages
          Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""AccountReport.asp?&iPageNo=" &  i & "&vJenisReport=" & request.Form("vJenisReport")& "&fcValue_Organization=" & request("fcValue_Organization") & """>" & i & "  "
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
<%
End If
rs.close
set rs = nothing
    %>

<!--#include file="../Bottom_Report.asp"-->