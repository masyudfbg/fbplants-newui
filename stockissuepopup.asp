<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Print Stock Issue</title>
<link rel="stylesheet" href="css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
   
</head>
<style type="text/css">


#btnselect {
  background-color: #4CAF50; /* Green */
  border: none;
  color: white;
  padding: 3.5px 8px;
  text-align: center;
  text-decoration: none;
  display: inline-block;
  font-size: 12px;
}


</style> 
<body>
<%


 ' v_txtFcba = Request("fcba") 
  varDefaultBa = Session("fcBA")   

  If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
    v_txtFcba = varDefaultBa
  End If  
    

    ' selectCons = selectCons & ",fcba"
    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
  q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
  set rsCat = connStr.Execute(q)
  if not rsCat.eof then
     fcbaname = rsCat("fcName")
     CompanyName=rsCat("FCCOMPANYNAME")
     MgrName = rsCat("FCMANAGER")
     KTUName = rsCat("FCKTU")
     CropName=rsCat("FCCROP")
     FcbaType=rsCat("fcType")
 ' else
  '    nSpanTotal = nSpanTotal - 1 
  end if

  Dim strqry,filterby,txtcari,btncari,filterdate,filtercode,rsdata
  

%>
<div id="wrapper">
<div class="navi_to_print">

</div>
<div class="letter">
<div id="note">
<form action="stockissuepopup.asp">
<div class="page" >  
  <table border="0" width="100%" id="product-table">
    <thead>
      <tr>
        <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
          <div id="header-wrapper">
             <div id="head_judul" style="float:left; text-align:left;">
              <h4>PT. <%=CompanyName%></h4>
              <h5><%=FcbaName%></h5>
            </div>
            <div id="subheader" class="centered" style="width:50%; margin:auto;"> 
              <h1 class="judul">Stock Issue</h1> 
            </div>  
            <div style="clear:both;"></div>
          </div>
        </th>
      </tr> 
    
      <tr>
        <th colspan="5">
          Search by : &nbsp;&nbsp;&nbsp;
          <select name="filterby" id="filterby">
              <option value="all">Semua</option>
              <option value="sivdate">SIV Date</option>
              <option value="sivcode">SIV Code</option>
          </select>&nbsp;&nbsp;&nbsp;<input type="text" name="txtcari" id="txtcari" placeholder="cari data" style=width:300px;height:20px></input>&nbsp;&nbsp;&nbsp;<input type="SUBMIT" name="btncari" value="VIEW DATA"></input>
        </th>
      </tr>
    
      <!--</div>--> 
        <th style="background-color: #0c5b24; color: #fff;">Action</th>
        <th style="background-color: #0c5b24; color: #fff;">Issue Date</th>
        <th style="background-color: #0c5b24; color: #fff;">Issue Code</th>
        <th style="background-color: #0c5b24; color: #fff;">Store</th> 
        <th style="background-color: #0c5b24; color: #fff;">Doc Ref</th>
        
      </tr>
    </thead>
  <tbody>
<%
  filterby = request("filterby")
  txtcari = request("txtcari")
  btncari = request("btncari")

  if btncari = "VIEW DATA" then
    if trim(filterby) = "sivdate" then
		  filterdate = " AND sivdate=to_Date('" & txtcari & "','dd/mm/yyyy')"
	  end if
    if trim(filterby) = "sivcode" then
		  filtercode = " AND sivcode like '%" & txtcari & "%'"
	  end if
    strqry		=	"Select SIVCODE, STORECODE, to_char(SIVDATE,'dd/mm/yyyy') SIVDATE, FCENTRY, FCEDIT, " & _
    " FCIP, FCBA, LASTUPDATE, LASTTIME, DOC_REF " & _
    " From  ISSUE " & _
    " WHERE FCBA='" & varDefaultBa & "' AND STORECODE='ST1' " & _
    filterdate & _ 
    filtercode & _
    " ORDER BY SIVDATE,SIVCODE DESC " 
  



    'response.write strqry  
    Set rsdata = Server.CreateObject("ADODB.Recordset")
    rsdata.Open strqry, connstr, adOpenStatic 

  dim no 
  no = 1
  while not rsdata.eof 


  %> 
    <tr class="">
      <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><button id="btnselect" onclick="getSelected(<%=no%>)" >Select</button></td>
      <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsdata("sivdate")%></td>
      <td align="Center" style="<%=cmdstyle%>"><input type="hidden" name="noref" id="noref<%=no%>" value="<%=rsdata("sivcode")%>"></input><font face="Verdana" size="1"><%=rsdata("sivcode")%>
      </td>
      <td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsdata("storecode")%>
      </td> 
      <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsdata("DOC_REF")%></td>
    
    </tr>
<%    
      no = no + 1  
      rsdata.MoveNext
    Wend
  rsdata.Close
  set rsdata = Nothing
  end if
%> 
</tbody>  
</table> 
</body>
</html>
<script src="/js/jquery-1.9.1.js"></script>
<script>
 
  function getSelected(a) {

      var noref = $('#noref' + a).val();
      window.opener.parent.document.getElementById("noref").value = noref;
      window.opener.parent.getDetailreturn();
      window.close();
     
	}
</script>






