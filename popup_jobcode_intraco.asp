<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Print Stock On Hand</title>
<link rel="stylesheet" href="css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />

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
</head>

<body>
<%

Dim sqljob,rsjob,v_txtFcba

  v_txtFcba = session("fcba") 
  
  q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
  set rsCat = connStr.Execute(q)
  if not rsCat.eof then
     fcbaname = rsCat("fcName")
     CompanyName=rsCat("FCCOMPANYNAME")
     MgrName = rsCat("FCMANAGER")
     KTUName = rsCat("FCKTU")
     CropName=rsCat("FCCROP")
     FcbaType=rsCat("fcType")
  else
      nSpanTotal = nSpanTotal - 1 
  end if


	
	sqljob		=	"SELECT FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, UNITOFMEASUREMENT, ACCOUNTCODE " & _
								      " FROM JOB " & _
                      " WHERE ACTIVATION = 'Y' AND FCBA = '" & session("FCBA") & "' " & _ 
                      " ORDER BY FCCODE ASC"

  Set rsjob = Server.CreateObject("ADODB.Recordset")
  rsjob.Open sqljob, connstr, adOpenStatic 

%>
<div id="wrapper">
<div class="navi_to_print">

</div>
<div class="letter">
<div id="note">

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
              <h1 class="judul">Data Master Job</h1> 
            </div>  
            <div style="clear:both;"></div>
          </div>
        </th>
      </tr> 
     
      <tr>
        <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
      </tr>
     
        <th style="background-color: #32CD32; color: #fff;">Action</th>
        <th style="background-color: #32CD32; color: #fff;">Job Code</th>
        <th style="background-color: #32CD32; color: #fff;">Activity Name</th>
        <th style="background-color: #32CD32; color: #fff;">Job Group Code</th>
        <th style="background-color: #32CD32; color: #fff;">Job For Crop Type</th>
        <th style="background-color: #32CD32; color: #fff;">Job For Field Status</th>
        <th style="background-color: #32CD32; color: #fff;">Job Category</th>
        <th style="background-color: #32CD32; color: #fff;">UOM</th>
        <th style="background-color: #32CD32; color: #fff;">Acount Code</th>
      </tr>
    </thead>
  <tbody>
<%
   dim no 
   no =1
  while not rsjob.eof 

  %> 
    <tr class="<%=cClass%>">
      <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><button id="btnselect" onclick="getSelected(<%=no%>)" >Select</button></td>
      <td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsjob("FCCODE")%>
        <input type=hidden id='jobcode<%=no%>' value="<%=rsjob("FCCODE")%>"></input>
      </td>
      <td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsjob("FCNAME")%>
        <input type=hidden id='jobname<%=no%>' value="<%=rsjob("FCNAME")%>"></input>
      </td> 
      <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsjob("JOBGROUPCODE")%></td>
      <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsjob("JOB_FORCROPTYPE")%></td>
      <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsjob("JOB_FORFIELDSTATUS")%></td>
      <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsjob("JOB_CATEGORY")%></td>
      <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsjob("UNITOFMEASUREMENT")%>
        <input type=hidden id='jobcategori<%=no%>' value="<%=rsjob("JOB_CATEGORY")%>"></input>
      </td>
      <td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsjob("ACCOUNTCODE")%></td>
      
    </tr>
<%      
      rsjob.MoveNext
      no = no +1
    Wend
  rsjob.Close
  set rsjob = Nothing
%> 
</tbody>  
</table> 
</body>
</html>
<script src="/js/jquery-1.9.1.js"></script>
<script>
 
  function getSelected(a) {

      var jobcode     = $('#jobcode' + a).val();
      var jobname     = $('#jobname' + a).val();
      var jobcategori = $('#jobcategori' + a).val();
      window.opener.parent.document.getElementById("chargejob_").value = jobcode;
      window.opener.parent.document.getElementById("jobname").innerHTML = jobname;
      window.opener.parent.document.getElementById("chargetype_").value = jobcategori;
      window.close();
	}
</script>





