<link rel="stylesheet" href="2col_leftNav.css" type="text/css" style="masthead" />
<script type="text/javascript">
		function cancelFullScreen(el) {
            var requestMethod = el.cancelFullScreen||el.webkitCancelFullScreen||el.mozCancelFullScreen||el.exitFullscreen;
            if (requestMethod) { // cancel full screen.
                requestMethod.call(el);
            } else if (typeof window.ActiveXObject !== "undefined") { // Older IE.
                var wscript = new ActiveXObject("WScript.Shell");
                if (wscript !== null) {
                    wscript.SendKeys("{F11}");
                }
            }
        }

        function requestFullScreen(el) {
            // Supports most browsers and their versions.
            var requestMethod = el.requestFullScreen || el.webkitRequestFullScreen || el.mozRequestFullScreen || el.msRequestFullScreen;

            if (requestMethod) { // Native full screen.
                requestMethod.call(el);
            } else if (typeof window.ActiveXObject !== "undefined") { // Older IE.
                var wscript = new ActiveXObject("WScript.Shell");
                if (wscript !== null) {
                    wscript.SendKeys("{F11}");
                }
            }
            return false
        }

        function toggleFull() {
            var elem = document.body; // Make the body go full screen.
            var isInFullScreen = (document.fullScreenElement && document.fullScreenElement !== null) ||  (document.mozFullScreen || document.webkitIsFullScreen);

            if (isInFullScreen) {
                cancelFullScreen(document);
            } else {
                requestFullScreen(elem);
				 alert("Tekan 'F11' pada Keyboard Anda, untuk selalu Full Screen.");
            }
            return false;
        }

		
</script>
<style>
#siteInfo{ background-image:url(IMAGES/login_02.gif); background-repeat: repeat-x; height:107px; border:none; left: 0;
    position: fixed;
    right: 0;
    top: 0; }
.colR{ float:right; margin-top:30px;  }
.logo { margin-top:10px; width:30%; float:left; }
</style>
<div id="siteInfo" style="height:60px; border-bottom:4px #FFCC00 solid;">
    <img src="/css/login/images/fbg-logo.png" alt="FBG" class="fbp-logo" height="44" style="margin-left:10px; float:left;">
	<!--<font color="yellow" pointsize="7px">iPlaS</font>-->
	&nbsp;&nbsp;&nbsp;
	<h1 style="float:left; width:300px; margin-top:21px; margin-left:10px;"><font color="#FFCC00" style="margin-top:5px;">First Borneo Plantation System</font></h1>
	<div id="globalNav">
       <%if Session("USERNAME")="" then
		Response.Redirect("logout.asp")
		%>
		<a class="mn_top" href="logout.asp"><span>log-in</span></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Please login first...</B><a class="mn_top"  href="login.Asp"><i class="home_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Home</span></a></div>
		<%else%>
            <a class="mn_top" href="logout.asp"><i class="off_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>log-out</span></a>
            <a id="button" class="mn_top_full" href="javascript:void(0);" onclick="toggleFull();">
				<i class="full">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Full Screen</span>
			</a>
            <a id="button" class="mn_top" target="_blank" style="width: 25px;text-align:center" href="http://36.37.102.166:90/fbghelpdesk/public/"><span>Help</span></a>
            <a id="button" class="mn_top" target="_blank" style="width: 45px;text-align:center" href="http://36.37.102.166:90/FBPlants-Guidance/index.html"><span>Guidance</span></a>
            <a id="button" class="mn_top" target="_blank" style="width: auto;text-align:center" href="https://www.janjangemas.id/id/auth/login"><span>Janjang Emas</span></a>
            <a id="button" class="mn_top" style="width: 25px;text-align:center" href="http://firstborneo.com/dx/"><span>DX</span></a>
            <a class="mn_top"  href="login.Asp"><i class="home_icon">&nbsp;&nbsp;&nbsp;&nbsp;</i><span>Home</span></a>
		</div>
      
        <br/>
        <div class="identity" ><b><%=session("FULLNAME")%></b> <font size="4" color="#FFCC00">|</font> <b>BA : <%=Session("BA")%> ( <%=session("fcBA")%> )</b> <font size="4"  color="#FFCC00">|</font> Current Period : <%=FormatDateTime(Cdate(Now),1)%>
		<%end if%> </div>
	
</div>