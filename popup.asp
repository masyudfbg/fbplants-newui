<html>
<head>
<body>
<% 
dim Id
ID = Request.QueryString("ID")
%>
<script language="javascript">
function winOpen(strURL)
{     
newwind = window.open(strURL, null, "height=200,width=400,status=yes,toolbar=no,menubar=no,location=no"); 
}

</script>


</script>
</body>
 </head>
              
<h4 align=center>Please Click On Following Links</h4>
        
<div align="center">
  <center>
             
<TABLE border=1 cellPadding=1 cellSpacing=1 width="75%">
  
  <TR>
    <TD valign="middle" align="center">No</TD>
    <TD valign="middle" align="center">Model</TD>
    <TD valign="middle" align="center">type</TD>      </center>
</div>
                           
                  
 <%
strDB= "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="
        strDB = strDB & Server.MapPath("images.mdb")
    
                  set data = server.createobject("ADODB.RecordSet")
                  data.activeconnection = strDB
                  data.source = "SELECT * FROM try1" 
                  data.PageSize = 2
                  'data.CacheSize  = 5
                  data.CursorLocation = 3
                  data.open
                   
                  If Len(Request("pagenum")) = 0  Then
                data.AbsolutePage = 1
            Else
                If CInt(Request("pagenum")) <= data.PageCount Then
                        data.AbsolutePage = Request("pagenum")
                    Else
                        data.AbsolutePage = 1
                End If
        End If
        Dim abspage, pagecnt
                Dim fldF, intRec
            abspage = data.AbsolutePage
            pagecnt = data.PageCount
                


                  For intRec=1 To data.PageSize
                   
                   if data.EOF  then
                   
                   else
                                                     
            Response.Write "<tr><td valign='middle' align='center'>"
            Response.Write data("Id") & "</td><td valign='middle' align='center'>"
           'Response.Write "<a href=""file.asp?ID=" & data("Id") & """>"
           
              '(here is the code which i had change as you suggest)

                Response.Write "<a href=""javascript:winOpen('file.asp?ID=" & data(0) & "')" & """>"   
            
            
            
            
            Response.Write data(1)
            Response.Write "</td><td valign='middle' align='center'>" & data("Model Type") & "</td></tr>"
            'Response.Write "</td><td valign='middle' align='center'>" & data("Content Type") & "</td></tr>"
                      
            data.MoveNext
                                   
                       end if
                    next
                    
                   
                    data.close
                    set data = nothing %></tr></table>
           

        %> 
                  </html>  