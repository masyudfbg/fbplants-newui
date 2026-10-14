<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Stock Ledger</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="../plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="../plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="../plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="../plugins/responsivegridsystem/12cols.css" media="all">
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

<%
' Const adOpenForwardOnly = 0
' Const adLockReadOnly = 1
' Const adCmdTableDirect = &H0200
' Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

  v_txtGroupPlantation  = Request("txtGroupPlantation")
  v_txtPlantation       = Request("txtPlantation")
  v_txtGroupRegion      = Request("txtGroupRegion")
  v_txtRegion           = Request("txtRegion")
  v_txtcompany          = Request("txtCompany")
  v_txtfcba             = Request("txtfcba")
  v_txtDivision         = Request("txtDivision")
  v_txtBlok             = Request("txtBlok")
  v_txtPlantingDate     = Request("txtPlantingDate")
  v_txttodaydate        = Request("txttodaydate")
  v_txtYear             = Request("txtYear")
  v_txtMonth            = Request("txtmonth")
  v_txtcustomercode       = Request("txtcustomercode")
  v_txtstockgroup=trim(request("txtstockgroup"))
  v_txtitemcode=trim(request("txtitemcode"))
  v_txtstorecode=trim(request("txtstorecode"))
  
  txtyear       = right(Request("txttodaydate"),4)
  txtmonth      = left(Request("txttodaydate"),2)
  txttodaydate  = Request("txttodaydate")
  
  
  txtComp         = Request("txtCompany")
  v_txtdate_from = trim(Request("txtdate_from"))
  v_txtdate_to   = trim(Request("txtdate_to "))

  if "" & v_txtdate_from = "" then v_txtdate_from = textDate(date)
  if "" & v_txtdate_to = "" then v_txtdate_to = textDate(date)

  nSpanTotal = 8
  v_txtComp           =   txtComp
  
  if request("txtYear")="" then
     v_txtYear  = Year(Date)
  end if
  if request("txtMonth")="" then
     v_txtMonth = Month(date)
  end if
     
    'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
    
    Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
    strFieldCons =  "   USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "', fcba) GroupPlantation" & _
                    " , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "', fcba) Plantation" & _
                    " , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "', fcba) groupregion" & _
                    " , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "', fcba) region" & _
                    " , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', fcba) company" & _
                    " , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"

             
    wherecons = " where company is not null " 

    q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"  
      set rsCat = connStr.Execute(q)
      if not rsCat.eof then
         CompanyName=rsCat("FCCOMPANYNAME")        
        end if 

    q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
      set rsCat = connStr.Execute(q)
      if not rsCat.eof then
         fcbaname = rsCat("fcName")
         CompanyName=rsCat("FCCOMPANYNAME")
         MgrName = rsCat("FCMANAGER")
         KTUName = rsCat("FCKTU")
         CropName=rsCat("FCCROP")
         FcbaType=rsCat("fcType")
        end if 

    If Len(Trim(v_txtGroupPlantation)) > 0 Then
      If v_txtGroupPlantation = "--SUMMARY--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " "
      Elseif v_txtGroupPlantation = "--ALL--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " " 
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " And GroupPlantation = '" & v_txtGroupPlantation & "'"
        orderCons = orderCons & " " 
        orderCons2 = orderCons2 & " "
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtPlantation)) > 0 Then
      If v_txtPlantation = "--SUMMARY--" then
        selectCons  = selectCons  & " " 
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " "
      Elseif v_txtPlantation = "--ALL--" then
        selectCons  = selectCons  & " , Plantation" 
        selectCons2 = selectCons2  & " , Plantation" 
        whereCons = whereCons & " "
        orderCons = orderCons & " , Plantation" 
        orderCons2 = orderCons2 & " , Plantation" 
        GroupCons = GroupCons & " , Plantation" 
        GroupCons2  =  GroupCons2 & " , Plantation"
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , Plantation"
        selectCons2 = selectCons2  & " , Plantation" 
        whereCons = whereCons & " And Plantation = '" & v_txtPlantation & "'"
        orderCons = orderCons & " , Plantation" 
        orderCons2 = orderCons2 & " , Plantation" 
        GroupCons   =  GroupCons & " , Plantation"
        GroupCons2  =  GroupCons2 & " , Plantation"
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtRegion)) > 0 Then
      If v_txtRegion = "--SUMMARY--" then
        selectCons  = selectCons  & " " 
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
      Elseif v_txtRegion = "--ALL--" then
        selectCons  = selectCons  & " , Region"
        selectCons2 = selectCons2  & " , Region"  
        whereCons = whereCons & " "
        orderCons = orderCons & " , Region" 
        orderCons2 = orderCons2 & " , Region" 
        GroupCons   =  GroupCons & " , Region"
        GroupCons2  =  GroupCons2 & " , Region"
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , Region" 
        selectCons2 = selectCons2  & " , Region" 
        whereCons = whereCons & " And Region = '" & v_txtRegion & "'"
        orderCons = orderCons & " , Region" 
        orderCons2 = orderCons2 & " , Region" 
        GroupCons   =  GroupCons & " , Region"
        GroupCons2  =  GroupCons2 & " , Region"
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtGroupRegion)) > 0 Then
      If v_txtGroupRegion = "--SUMMARY--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " "
      Elseif v_txtGroupRegion = "--ALL--" then
        selectCons  = selectCons  & " , GroupRegion" 
        whereCons = whereCons & " "
        orderCons = orderCons & " , GroupRegion" 
        orderCons2 = orderCons2 & " , GroupRegion" 
        GroupCons   =  GroupCons & " ,  GroupRegion"
        GroupCons2  =  GroupCons2 & " ,  GroupRegion"
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , GroupRegion"
        selectCons2  = selectCons2 & " , GroupRegion" 
        whereCons = whereCons & " And GroupRegion = '" & v_txtGroupRegion & "'"
        orderCons = orderCons & " , GroupRegion"
        orderCons2 = orderCons2 & " , GroupRegion" 
        GroupCons   =  GroupCons & " , GroupRegion"
        GroupCons2  =  GroupCons2 & " , GroupRegion"
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtCompany)) > 0 Then
      If v_txtCompany = "--SUMMARY--" then
        selectCons  = selectCons  & " " 
        selectCons2  = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " "
      Elseif v_txtCompany = "--ALL--" then
        selectCons  = selectCons  & " , company"
        selectCons2  = selectCons2  & " , company" 
        whereCons = whereCons & " "
        orderCons = orderCons & " , COMPANY_API.get_OrderData(company)"
        orderCons2 = orderCons2 & " , COMPANY_API.get_OrderData(company)" 
        GroupCons   =  GroupCons & " , company"
        GroupCons2  =  GroupCons2 & " , company"
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , company"
        selectCons2  = selectCons2  & " , company" 
        whereCons = whereCons & " And company = '" & v_txtCompany & "'"
        orderCons = orderCons & " , COMPANY_API.get_OrderData(company)"
        orderCons2 = orderCons2 & " , COMPANY_API.get_OrderData(company)"
        GroupCons   =  GroupCons & " , company" 
        GroupCons2  =  GroupCons2 & " , company" 
        var_colspan = var_colspan + 1
      End if
    End if

    If Len(Trim(v_txtfcba)) > 0 Then
      If v_txtfcba = "--SUMMARY--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        whereCons2 = whereCons2 & " "
        orderCons = orderCons & " "
        orderCons2 = orderCons2 & " "
      Elseif v_txtfcba = "--ALL--" then
        selectCons  = selectCons  & " , fcba" 
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        whereCons2 = whereCons2 & " "
        orderCons = orderCons & " , fcba"
        orderCons2 = orderCons2 & " "
        GroupCons   =  GroupCons & " , fcba"
        GroupCons2  =  GroupCons2 & " "
        var_colspan = var_colspan + 1
      Else
        selectCons  = selectCons  & " , fcba" 
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " And fcba = '" & v_txtfcba & "'"
        whereCons2 = whereCons2 & " And fcba = '" & v_txtfcba & "'"
        orderCons = orderCons & " , fcba"
        orderCons2 = orderCons2 & " "
        GroupCons  =  GroupCons & " , fcba"
        GroupCons2 =  GroupCons2 & " "
        var_colspan = var_colspan + 1
      End if
    End if 
    
    If Len(Trim(v_txtstockgroup)) > 0 Then
      If v_txtstockgroup = "--SUMMARY--" then
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " " 
        whereCons = whereCons & " "
        orderCons = orderCons & " "
      Elseif v_txtstockgroup = "--ALL--" then
        selectCons  = selectCons  & " " 
        selectCons2 = selectCons2  & " "
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        GroupCons = GroupCons & " "
      Else
        selectCons  = selectCons  & " "
        selectCons  = selectCons  & " " 
        whereCons = whereCons & " And stock_api.get_stockgroup (stockcode, fcba) = '" & v_txtstockgroup & "'"
        orderCons = orderCons & " "
        GroupCons = GroupCons & " "
      End if
    End If

    If Len(Trim(v_txtitemcode)) > 0 Then
      If v_txtitemcode = "--SUMMARY--" then
        selectCons  = selectCons  & " " 
        selectCons2 = selectCons2  & " "
        whereCons = whereCons & " "
        orderCons = orderCons & " "
      Elseif v_txtitemcode = "--ALL--" then
        selectCons  = selectCons  & " ,stockcode , stockname, uom"
        selectCons2 = selectCons2  & " "  
        whereCons = whereCons & " "
        orderCons = orderCons & " "
        GroupCons   =  GroupCons & " , stockcode , stockname, uom"
        var_colspan = var_colspan + 3
      Else
        selectCons  = selectCons  & " ,stockcode, stockname, uom"
        selectCons2 = selectCons2  & " "
        whereCons = whereCons & " And stockcode = '" & v_txtitemcode & "'"
        orderCons = orderCons & " "
        GroupCons   =  GroupCons & " , stockcode, stockname, uom"
        var_colspan = var_colspan + 3
      End if
    End If 

    If Len(Trim(v_txtstorecode)) > 0 Then
      If v_txtstorecode = "--SUMMARY--" then
        selectCons  = selectCons  & " " 
        selectCons2 = selectCons2  & " "
        whereCons = whereCons & " "
        orderCons = orderCons & " "
      Elseif v_txtstorecode = "--ALL--" then
        selectCons  = selectCons  & "  "
        selectCons2 = selectCons2  & " "  
        whereCons = whereCons & " "
        orderCons = orderCons & " "
      Else
        selectCons  = selectCons  & " "
        selectCons2 = selectCons2  & " "
        whereCons = whereCons & " And storecode = '" & v_txtstorecode & "'"
        orderCons = orderCons & " "
      End if
    End If 
' Dim wherestatus       
' if "" & txtstatus <> "" and "" & txtstatus <> "--ALL--" then
'   wherestatus = " and field_api.GetValid_status(fieldcode,fcba,TO_DATE('" & v_txttodaydate & "', 'mm/dd/yyyy') ) = '" & txtstatus & "'"
' end if
  ScriptName  = Request.ServerVariables("SCRIPT_NAME")
  

txtSQL =	" select "& strFieldCons &"" & _
			"        , itemcode stockcode" & _
			"        , stock_api.get_stockname(itemcode, fcba) stockname" & _
			"        , stock_api.get_uom(itemcode, fcba) uom" & _
			"        , stockgroup_api.get_controljob(stock_api.get_stockgroup(itemcode, fcba), fcba) controljobcode" & _
			"        , job_api.get_jobcode_name(stockgroup_api.get_controljob(stock_api.get_stockgroup(itemcode, fcba), fcba), fcba) controljobname" & _
			"        , stock_api.get_stockgroup(itemcode, fcba)||' - '||stockgroup_api.Get_StockGrpName(stock_api.get_stockgroup(itemcode, fcba), fcba) stockgroup" & _
			"        , nvl(sum(nvl(openingbalanceqty, 0)), 0) openingbalanceqty" & _
			"        , nvl(sum(nvl(qtyreceived, 0)), 0) qtyreceived" & _
			"        , nvl(sum(nvl(qtyreceivedtask, 0)), 0) qtyreceivedtask" & _
			"        , nvl(sum(nvl(qtyreceivetaskdiv, 0)), 0) qtyreceivetaskdiv" & _
			"        , nvl(sum(nvl(qtyreturned, 0)), 0) qtyreturned" & _
			"        , nvl(sum(nvl(qtywriteoff, 0)), 0) qtywriteoff" & _
			"        , nvl(sum(nvl(qtyissued, 0)), 0) qtyissued" & _
			"        , nvl(sum(nvl(qtytransporttask, 0)), 0) qtytransporttask" & _
			"        , nvl(sum(nvl(qtytransporttaskdiv, 0)), 0) qtytransporttaskdiv" & _
			"        , nvl(sum(nvl(endingbalanceqty, 0)), 0) endingbalanceqty " & _
			"        , nvl(sum(nvl(openingbalanceamt, 0)), 0) openingbalanceamt" & _
			"        , nvl(sum(nvl(amtreceived, 0)), 0) amtreceived" & _
			"        , nvl(sum(nvl(amtreceivedtask, 0)), 0) amtreceivedtask" & _
			"        , nvl(sum(nvl(amtreceivetaskdiv, 0)), 0) amtreceivetaskdiv" & _
			"        , nvl(sum(nvl(amtreturned, 0)), 0) amtreturned" & _
			"        , nvl(sum(nvl(amtwriteoff, 0)), 0) amtwriteoff" & _
			"        , nvl(sum(nvl(amtissued, 0)), 0) amtissued" & _
			"        , nvl(sum(nvl(amttransporttask, 0)), 0) amttransporttask" & _
			"        , nvl(sum(nvl(amttransporttaskdiv, 0)), 0) amttransporttaskdiv" & _
			"        , nvl(sum(nvl(endingbalanceamt, 0)), 0) endingbalanceamt " & _
			"        , round(nvl(sum(nvl(amtperunit, 0)), 0), 2) amtperunit" & _
			" from stockledger" & _
			" where " & _
			"	    year = '"& v_txtYear &"'" & _
			"	    and month = '"& v_txtMonth &"'" & _
			" group by fcba" & _
			"          , itemcode"


' if  len(trim(v_txtyear)) > 0 then
'     txtSQL = txtSQL & " and year = " & v_txtyear
' 	txtSelect = txtSelect & " Year " & v_txtyear
' end if
' if  len(trim(v_txtmonth)) > 0 then
'     txtSQL =txtSQL &  " and month = " & v_txtmonth
' 	txtSelect = txtSelect & " Month " & v_txtmonth
' end if
' if  len(trim(v_txtstorecode)) > 0 then
'     txtSQL = txtSQL & " and storecode = '" & v_txtstorecode & "'"
' 	txtSelect = txtSelect & " Store Code " & v_txtstorecode
' end if
' if  len(trim(v_txtstockgroup)) > 0 then
'     txtSQL = txtSQL & " and itemcode in (select fccode from stock where stockgroupcode = " & v_txtstockgroup & ")"
' 	txtSelect = txtSelect & " Stock Group " & v_txtstockgroup
' end if
' if  len(trim(v_txtstock)) > 0 then
'     txtSQL =txtSQL &  " and itemcode = '" & v_txtstock & "'"
' 	txtSelect = txtSelect & " Stock Code " & v_txtstock
' end if 

  strEPS2 = "        " & _
            "        , stockgroup" & _
			"        , nvl(sum(nvl(openingbalanceqty, 0)), 0) openingbalanceqty" & _
			"        , nvl(sum(nvl(qtyreceived, 0)), 0) qtyreceived" & _
			"        , nvl(sum(nvl(qtyreceivedtask, 0)), 0) qtyreceivedtask" & _
			"        , nvl(sum(nvl(qtyreceivetaskdiv, 0)), 0) qtyreceivetaskdiv" & _
			"        , nvl(sum(nvl(qtyreturned, 0)), 0) qtyreturned" & _
			"        , nvl(sum(nvl(qtywriteoff, 0)), 0) qtywriteoff" & _
			"        , nvl(sum(nvl(qtyissued, 0)), 0) qtyissued" & _
			"        , nvl(sum(nvl(qtytransporttask, 0)), 0) qtytransporttask" & _
			"        , nvl(sum(nvl(qtytransporttaskdiv, 0)), 0) qtytransporttaskdiv" & _
			"        , nvl(sum(nvl(endingbalanceqty, 0)), 0) endingbalanceqty " & _
			"        , nvl(sum(nvl(openingbalanceamt, 0)), 0) openingbalanceamt" & _
			"        , nvl(sum(nvl(amtreceived, 0)), 0) amtreceived" & _
			"        , nvl(sum(nvl(amtreceivedtask, 0)), 0) amtreceivedtask" & _
			"        , nvl(sum(nvl(amtreceivetaskdiv, 0)), 0) amtreceivetaskdiv" & _
			"        , nvl(sum(nvl(amtreturned, 0)), 0) amtreturned" & _
			"        , nvl(sum(nvl(amtwriteoff, 0)), 0) amtwriteoff" & _
			"        , nvl(sum(nvl(amtissued, 0)), 0) amtissued" & _
			"        , nvl(sum(nvl(amttransporttask, 0)), 0) amttransporttask" & _
			"        , nvl(sum(nvl(amttransporttaskdiv, 0)), 0) amttransporttaskdiv" & _
			"        , nvl(sum(nvl(endingbalanceamt, 0)), 0) endingbalanceamt " & _
			"        , round(nvl(sum(nvl(amtperunit, 0)), 0), 2) amtperunit" 

  strEPS =    " SELECT GroupPlantation " & selectCons & strEPS2 &"  From (" & txtSQL & ")" & whereCons & ""  & _
              " Group by GroupPlantation , stockgroup " & GroupCons & "" & _
              " ORDER BY GroupPlantation " & orderCons & ""

' txtsql = txtsql &	" group by fcba" & _
' 					"          , storecode" & _
' 					"          , itemcode"
' txtsql = txtsql &	" order by fcba" & _
' 					"          , storecode" & _
' 					"          , stockgroup_api.get_controljob(stock_api.get_stockgroup(itemcode, fcba), fcba)" & _
' 					"          , stock_api.get_stockgroup(itemcode, fcba)" & _
' 					"          , itemcode"
' response.Write(strEPS)
response.Flush()
%>
<body>
	  <div id="wrapper">
  <!--#include file="../print_excel.asp"-->
<form name="frmEntry" method="post" action="<%=ScriptName%>">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	<!--#include file="FormSelectNewSum.html"-->
	<table border="0" id="form-table" class="centered-margin">
		<tr>
			<td align="center"><font face="verdana" size="1" style ="font-weight: bold;">YEAR</font></td>
			<td align="center"><font face="verdana" size="1" style ="font-weight: bold;">MONTH</font></td>  
			
		</tr>
		<tr> 
		<% 
		  Dim mon, Yr, MinYr, MaxYr, i, MM
		  mon = Month(Session("v_Costbook"))
		  If len(v_txtMonth)=0 then v_txtMonth=mon

		  Yr = Year(Session("v_Costbook"))
		  MinYr = 2015
		  MaxYr = Yr  + 2
		  If len(v_txtYear)= 0 then v_txtYear = Yr
		%>
		    <td width="4">
		      <select name="txtYear">
		<%
		  For i = MinYr to MaxYr
		%>
		    <Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
		<%
		  Next
		%>
		      </select>
		    </td> 
		    <td width="2">
		      <select name="txtMonth">
		<%
		    For i = 1 to 15
		%>
		    <Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=BulanAP(i)%></option>
		<%
		  Next
		%>
		      </select>
		    </td>

			<td align="center">
				<input type="submit" value="View" />
			</td> 
		</tr>
	</table> 
        <table border="0" width="100%" id="product-table" class="xthf-site">
        <thead>
        <tr>
        <th colspan="<%=21+ var_colspan%>" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>STOCK LEDGER - PER STOCK GROUP</h2>
				</div>
			</div>
			  </th>
			</tr>
	<!--<div class="printhide">-->
						<tr>
							<th colspan="<%=21+ var_colspan%>" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
	<!--</div>-->
	<tr>
	<%	
'	response.write strEPS
	Set rsCat = connStr.Execute(strEPS)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
	<TR>
		<% if v_txtGroupPlantation <> "--SUMMARY--" and v_txtGroupPlantation <> "" then %>
            <th style="background-color:#00A300;color:#fff;" rowspan="2">Group Plantation</th>
      <%      
        end if %>

      <% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
            <th style="background-color:#00A300;color:#fff;" rowspan="2">Plantation</th>
      <%      
        end if %>
      
      <% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
            <th style="background-color:#00A300;color:#fff;" rowspan="2">Group Region</th>
      <%      
        end if %>
      
      <% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
            <th style="background-color:#00A300;color:#fff;" rowspan="2">Region</th>
      <%
        end if %>
      
      <% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
             <th style="background-color:#00A300;color:#fff;" rowspan="2">Company</th>
      <%    
        end if %>
      
      <% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
             <th style="background-color:#00A300;color:#fff;" rowspan="2">Site</th>
      <%    
        end if %>  
   
	  <th rowspan="2">Stock Group</th>
    
	  <% if v_txtitemcode <>  "--SUMMARY--" and v_txtitemcode <>  "" then %>
             <th rowspan="2">Stock Code</th>
		     <th rowspan="2">Description</th>
			 <th rowspan="2">UOM</th>
      <%    
        end if %>  
		
		<Th colspan="10">Quantity</Th>
		<Th colspan="10">Amount (Rp)</Th>
	
	</TR>
	<tr>
		<th>Opening</th>
		<th>Received</th>
		<th>Received.Task</th>
		<th>Received.Div</th>
		<th>Return</th>
		<th>WriteOff</th>
		<th>Issued</th>
		<th>Transport.Task</th>
		<th>Transport.Div</th>
		<th>Saldo</th>
		
		<th>Opening</th>
		<th>Received</th>
		<th>Received.Task</th>
		<th>Received.Div</th>
		<th>Return</th>
		<th>WriteOff</th>
		<th>Issued</th>
		<th>Transport.Task</th>
		<th>Transport.Div</th>
		<th>Saldo</th>
	</tr>
  </thead>
  <tbody>
	<%
	While Not rsCat.EOF
		' Currcontroljobcode = rsCat("controljobcode")
        Currstockgroup = rsCat("stockgroup")
		If Currstockgroup <> Prevstockgroup Then
%>


		<tr>
			<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
          <% tempPlantation = rsCat("Plantation")
             tempGroupRegion = "--ALL--"
             tempRegion = "--SUMMARY--"
             tempCompany ="--SUMMARY--"
             tempfcba = "--SUMMARY--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"
          %>
          <% if SprevGroupPlantation <> rsCat("GroupPlantation") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("GroupPlantation")%>
              </a>
            </td>
            <% 
              SprevGroupPlantation = rsCat("GroupPlantation")
            %>  
            
          <% else %>
            <td align="center"></td>          
          <% end if %>  
        <% end if %>

        <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
          <% tempPlantation = rsCat("Plantation")
             tempGroupRegion = "--ALL--"
             tempRegion = "--SUMMARY--"
             tempCompany ="--SUMMARY--"
             tempfcba = "--SUMMARY--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"
          %>
          <% if SprevPlantation <> rsCat("Plantation") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("Plantation")%>
              </a>
            </td>
            <% 
              SprevPlantation = rsCat("plantation")
            %>  
            
          <% else %>
            <td align="center"></td>          
          <% end if %>  
        <% end if %>
        
        <% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
            <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
              tempPlantation = rsCat("Plantation")
            else
              tempPlantation = ""
            end if
             tempGroupRegion = rsCat("GroupRegion")
             tempRegion = "--ALL--"
             tempCompany ="--SUMMARY--"
             tempfcba = "--SUMMARY--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"
          %>
          
          <% if SprevGroupRegion <> rsCat("GroupRegion") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("groupregion")%>
              </a>
            </td>
            <% SprevGroupRegion = rsCat("GroupRegion") %> 
          <% else %>
            <td align="center"></td>          
          <% end if %>  
        <% end if %>
        
        <% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
            <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
                tempPlantation = rsCat("Plantation")
             else
                tempPlantation = ""
             end if
             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
                 tempGroupRegion = rsCat("GroupRegion")
             else
                 tempGroupRegion = ""
             end if
             
             tempRegion = rsCat("Region") 
             tempCompany ="--ALL--"
             tempfcba = "--SUMMARY--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"
          %>
          
          <% if SprevRegion <> rsCat("Region") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("region")%>
              </a>
            </td>
            <% SprevRegion = rsCat("region") %> 
          <% else %>
            <td align="center"></td>          
          <% end if %>  
        <% end if %>
        
        <% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
            <% 
             if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
                tempPlantation = rsCat("Plantation")
             else
                tempPlantation = ""
             end if
             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
                 tempGroupRegion = rsCat("GroupRegion")
             else
                 tempGroupRegion = ""
             end if
             if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
                 tempRegion = rsCat("Region") 
             else
                 tempRegion = ""
             end if
             tempCompany =rsCat("Company")
             tempfcba = "--ALL--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"            
           %>
          
          <% if SprevCompany <> rsCat("Company") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("company")%>
              </a>
            </td>
            <% SprevCompany = rsCat("Company") %> 
          <% else %>
            <td align="center"></td>          
          <% end if %>  

        <% end if %>
        
        <% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
            <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
                tempPlantation = rsCat("Plantation")
             else
                tempPlantation = ""
             end if
             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
                 tempGroupRegion = rsCat("GroupRegion")
             else
                 tempGroupRegion = ""
             end if
             if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
                 tempRegion = rsCat("Region") 
             else
                 tempRegion = ""
             end if
             if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
                 tempCompany =rsCat("Company")
             else
                 tempCompany =""
             end if
             tempfcba = rsCat("fcba")
             tempDivision = "--ALL--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"            
           %>
          
          <% if sprevFcba <> rsCat("fcba") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
                <%=rsCat("fcba")%>
              </a>
            </td>
            <% sprevFcba = rsCat("fcba") %> 
          <% else %>
            <td align="center"><%=rsCat("fcba")%></td>          
          <% end if %>   
        <% end if %>

          <td align="left"><font face="Verdana" size="1"><%=rsCat("stockgroup")%></font></td>
         
        <% if v_txtitemcode <>  "--SUMMARY--" and v_txtitemcode <>  "" then %>
            <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
                tempPlantation = rsCat("Plantation")
             else
                tempPlantation = ""
             end if
             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
                 tempGroupRegion = rsCat("GroupRegion")
             else
                 tempGroupRegion = ""
             end if
             if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
                 tempRegion = rsCat("Region") 
             else
                 tempRegion = ""
             end if
             if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
                 tempCompany =rsCat("Company")
             else
                 tempCompany =""
             end if
             tempstockcode = rsCat("stockcode")
             tempDivision = "--ALL--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"            
           %>
          
          <% if sprevstockcode <> rsCat("stockcode") then %>
              <td align="left"><font face="Verdana" size="1"><%=rsCat("stockcode")%></font></td>
		      <td align="left"><font face="Verdana" size="1"><%=rsCat("stockname")%></font></td>
		      <td align="left"><font face="Verdana" size="1"><%=rsCat("uom")%></font></td>
            <% sprevstockcode = rsCat("stockcode") %> 
          <% else %>
              <td align="left"><font face="Verdana" size="1"><%=rsCat("stockcode")%></font></td>
		      <td align="left"><font face="Verdana" size="1"><%=rsCat("stockname")%></font></td>          
          	  <td align="left"><font face="Verdana" size="1"><%=rsCat("uom")%></font></td>
          <% end if %>   
        <% end if %>
	
	
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("openingbalanceqty"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreceived"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreceivedtask"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreceivetaskdiv"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreturned"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtywriteoff"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyissued"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtytransporttask"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtytransporttaskdiv"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("endingbalanceqty"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("openingbalanceamt"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreceived"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreceivedtask"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreceivetaskdiv"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreturned"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtwriteoff"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtissued"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amttransporttask"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amttransporttaskdiv"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("endingbalanceamt"), 2)%></font></td>
		     
		</tr>
<%
		  subopeningbalanceamt = subopeningbalanceamt + cdbl(rsCat("openingbalanceamt"))
		  subamtreceived = subamtreceived + cdbl(rsCat("amtreceived"))
		  subamtwriteoff = subamtwriteoff + cdbl(rsCat("amtwriteoff"))
		  subamtissued = subamtissued + cdbl(rsCat("amtissued"))
          subamtsaldo = subamtsaldo + cdbl(rsCat("endingbalanceamt"))
		  subamttransporttask= subamttransporttask + cdbl(rsCat("AMTTRANSPORTTASK"))
		  subamtreceivedtask= subamtreceivedtask + cdbl(rsCat("AMTRECEIVEDTASK"))
%>
<%Else%>
		<tr>
		<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
          <% tempPlantation = rsCat("Plantation")
             tempGroupRegion = "--ALL--"
             tempRegion = "--SUMMARY--"
             tempCompany ="--SUMMARY--"
             tempfcba = "--SUMMARY--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"
          %>
          <% if SprevGroupPlantation <> rsCat("GroupPlantation") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("GroupPlantation")%>
              </a>
            </td>
            <% 
              SprevGroupPlantation = rsCat("GroupPlantation")
            %>  
            
          <% else %>
            <td align="center"></td>          
          <% end if %>  
        <% end if %>

        <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
          <% tempPlantation = rsCat("Plantation")
             tempGroupRegion = "--ALL--"
             tempRegion = "--SUMMARY--"
             tempCompany ="--SUMMARY--"
             tempfcba = "--SUMMARY--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"
          %>
          <% if SprevPlantation <> rsCat("Plantation") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("Plantation")%>
              </a>
            </td>
            <% 
              SprevPlantation = rsCat("plantation")
            %>  
            
          <% else %>
            <td align="center"></td>          
          <% end if %>  
        <% end if %>
        
        <% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
            <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
              tempPlantation = rsCat("Plantation")
            else
              tempPlantation = ""
            end if
             tempGroupRegion = rsCat("GroupRegion")
             tempRegion = "--ALL--"
             tempCompany ="--SUMMARY--"
             tempfcba = "--SUMMARY--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"
          %>
          
          <% if SprevGroupRegion <> rsCat("GroupRegion") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("groupregion")%>
              </a>
            </td>
            <% SprevGroupRegion = rsCat("GroupRegion") %> 
          <% else %>
            <td align="center"></td>          
          <% end if %>  
        <% end if %>
        
        <% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
            <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
                tempPlantation = rsCat("Plantation")
             else
                tempPlantation = ""
             end if
             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
                 tempGroupRegion = rsCat("GroupRegion")
             else
                 tempGroupRegion = ""
             end if
             
             tempRegion = rsCat("Region") 
             tempCompany ="--ALL--"
             tempfcba = "--SUMMARY--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"
          %>
          
          <% if SprevRegion <> rsCat("Region") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("region")%>
              </a>
            </td>
            <% SprevRegion = rsCat("region") %> 
          <% else %>
            <td align="center"></td>          
          <% end if %>  
        <% end if %>
        
        <% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
            <% 
             if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
                tempPlantation = rsCat("Plantation")
             else
                tempPlantation = ""
             end if
             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
                 tempGroupRegion = rsCat("GroupRegion")
             else
                 tempGroupRegion = ""
             end if
             if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
                 tempRegion = rsCat("Region") 
             else
                 tempRegion = ""
             end if
             tempCompany =rsCat("Company")
             tempfcba = "--ALL--"
             tempDivision = "--SUMMARY--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"            
           %>
          
          <% if SprevCompany <> rsCat("Company") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
              <%=rsCat("company")%>
              </a>
            </td>
            <% SprevCompany = rsCat("Company") %> 
          <% else %>
            <td align="center"></td>          
          <% end if %>  

        <% end if %>
        
        <% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
            <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
                tempPlantation = rsCat("Plantation")
             else
                tempPlantation = ""
             end if
             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
                 tempGroupRegion = rsCat("GroupRegion")
             else
                 tempGroupRegion = ""
             end if
             if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
                 tempRegion = rsCat("Region") 
             else
                 tempRegion = ""
             end if
             if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
                 tempCompany =rsCat("Company")
             else
                 tempCompany =""
             end if
             tempfcba = rsCat("fcba")
             tempDivision = "--ALL--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"            
           %>
          
          <% if sprevFcba <> rsCat("fcba") then %>
            <td align="center">
              <a href="<%=ScriptName%>?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
                <%=rsCat("fcba")%>
              </a>
            </td>
            <% sprevFcba = rsCat("fcba") %> 
          <% else %>
            <td align="center"><%=rsCat("fcba")%></td>          
          <% end if %>   
        <% end if %>

        	 <td align="left"><font face="Verdana" size="1"><%=rsCat("storecode")%></font></td>
           <td align="left"><font face="Verdana" size="1"><%=rsCat("stockgroup")%></font></td>

        <% if v_txtitemcode <>  "--SUMMARY--" and v_txtitemcode <>  "" then %>
            <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
                tempPlantation = rsCat("Plantation")
             else
                tempPlantation = ""
             end if
             if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
                 tempGroupRegion = rsCat("GroupRegion")
             else
                 tempGroupRegion = ""
             end if
             if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
                 tempRegion = rsCat("Region") 
             else
                 tempRegion = ""
             end if
             if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
                 tempCompany =rsCat("Company")
             else
                 tempCompany =""
             end if
             tempstockcode = rsCat("stockcode")
             tempDivision = "--ALL--"
             tempPlantingDate = "--SUMMARY--"
             tempBlok = "--SUMMARY--"            
           %>
          
          <% if sprevstockcode <> rsCat("stockcode") then %>
             <td align="left"><font face="Verdana" size="1"><%=rsCat("stockcode")%></font></td>
		     <td align="left"><font face="Verdana" size="1"><%=rsCat("stockname")%></font></td>
		     <td align="left"><font face="Verdana" size="1"><%=rsCat("uom")%></font></td>
            <% sprevstockcode = rsCat("stockcode") %> 
          <% else %>
             <td align="left"><font face="Verdana" size="1"><%=rsCat("stockcode")%></font></td>
		     <td align="left"><font face="Verdana" size="1"><%=rsCat("stockname")%></font></td>          
          	 <td align="left"><font face="Verdana" size="1"><%=rsCat("uom")%></font></td>
          <% end if %>   
        <% end if %>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("openingbalanceqty"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreceived"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreceivedtask"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreceivetaskdiv"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyreturned"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtywriteoff"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyissued"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtytransporttask"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtytransporttaskdiv"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("endingbalanceqty"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("openingbalanceamt"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreceived"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreceivedtask"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreceivetaskdiv"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtreturned"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtwriteoff"), 2)%></font></td>
		      <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amtissued"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amttransporttask"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amttransporttaskdiv"), 2)%></font></td>
			  <td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("endingbalanceamt"), 2)%></font></td>
			
		</tr>
		
<%
		
%>

<%End If%>
	  <%
      totalopeningbalanceqty = totalopeningbalanceqty + cdbl(rsCat("openingbalanceqty"))
      totalendingbalanceqty = totalendingbalanceqty + cdbl(rsCat("endingbalanceqty"))
		  
		  totalopeningbalanceamt = totalopeningbalanceamt + cdbl(rsCat("openingbalanceamt"))
		  totalamtreceived = totalamtreceived + cdbl(rsCat("amtreceived"))
		  totalamtwriteoff = totalamtwriteoff + cdbl(rsCat("amtwriteoff"))
		  totalamtissued = totalamtissued + cdbl(rsCat("amtissued"))
		  totalamtsaldo = totalamtsaldo + cdbl(rsCat("endingbalanceamt"))
		  totalamttransporttask = totalamttransporttask + cdbl(rsCat("AMTTRANSPORTTASK"))
		  totalamtreceivedtask  = totalamtreceivedtask + cdbl(rsCat("AMTRECEIVEDTASK"))
		  totalamtreturn        = totalamtreturn + cdbl(rsCat("amtreturned"))

	  rsCat.MoveNext

	  response.Flush()
	  no = no + 1
	  Prevstockgroup= Currstockgroup
	Wend
	rsCat.close()
	rsCat=null%>
		
		<tr style="background-color:#9CB770;"> 
			  <td align="Left" colspan="<%=11+var_colspan%>" style="padding-left:10px;"><font face="Verdana" size="1"><strong> Total</strong></font></td>
		     
		      <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalopeningbalanceamt, 2)%></strong></font></td>
		      <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtreceived, 2)%></strong></font></td>
			  <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtreceivedtask, 2)%></strong></font></td>
			  <td align="right"><font face="Verdana" size="1">&nbsp;</font></td>
			  <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtreturn, 2)%></strong></font></td>
		      <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtwriteoff, 2)%></strong></font></td>

		      <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtissued, 2)%></strong></font></td>	
              <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamttransporttask, 2)%></strong></font></td>	
			  <td align="right"><font face="Verdana" size="1">&nbsp;</font></td>
			  <td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamtsaldo, 2)%></strong></font></td>
		      <td align="right"><font face="Verdana" size="1"><strong>&nbsp;</strong></font></td>
		</tr>


	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
</table>
<!--#include file="../Bottom_Report.asp"-->
<script src="../js/jquery.inputmask.bundle.min.js"></script>
  <script src="../js/inputmask.binding.js"></script>
  <script src="../plugins/chosen/chosen.jquery.min.js"></script>
  <script type="text/javascript"> 
    // $(".chosen-select").empty();
    
    var config = {
      '.chosen-select'           : {},
      '.chosen-select-deselect'  : {allow_single_deselect:true},
      '.chosen-select-no-single' : {disable_search_threshold:10},
      '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
      '.chosen-select-width'     : {width:"45%"}
    }
    for (var selector in config) {
      $(selector).chosen(config[selector]);
    }
    // $(".chosen-select").trigger("chosen:updated");

  </script>
		</DIV>
	</BODY>
</HTML>
	