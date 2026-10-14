<style>
#siteInfo{ background-image:url(IMAGES/login_02.gif); background-repeat: repeat-x; height:107px; border:none; }
.colR{ float:right; margin-top:30px;  }
.logo { margin-top:10px; width:30%; float:left; }
</style>
 <link rel="stylesheet" href="2col_leftNav.css" type="text/css" style="masthead" />
 
<div id="siteInfo">
	 
	  <div id="globalNav" style="width:980px; margin:0 auto; float:none; position:static;"> 
      <div class="logo">
      <!--
	  <img src="iplas.png" width="170">
      -->
	  </div>
          <div class="colR">
          <%if Session("USERNAME")="" then%>
               <a href="logout.asp"><font size="6" color="#FFFFFF">LOGIN</font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Please login first...</B>
          <%else%>
              </br> 
              <a href="logout.asp"><font size="6" color="#FFFFFF"></font><font color="#FFFFFF">LOGOUT</font></a><b><%=session("FULLNAME")%></b> <font size="4" color="#FFCC00">|</font> <b>BA : <%=Session("BA")%> ( <%=session("fcBA")%> )</b> <font size="4"  color="#FFCC00">|</font> Current Period : <%=FormatDateTime(session("v_Costbook"),1)%>
           <%end if%>
           </div>
      </div>
</div>
