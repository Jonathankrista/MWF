<%@ LANGUAGE="JSCRIPT"%>
<!--#INCLUDE FILE="TXT_TheLibrary.asp"-->
<%var XType=""+Request.Form("XType");
if (XType=="S")
{
	var EmailFrom="bk@mwfgroupbahamas.com";
	var EmailTo="bk@mwfgroupbahamas.com";
	var Subject="MWF Group - Contact Form";

	var TheName=""+Request.Form("TheName");
	var TheEmail=""+Request.Form("TheEmail");
	var TheSubject=""+Request.Form("TheSubject");
	var TheMessage=""+Request.Form("TheMessage");
	var HTML="<html>";
			HTML="<body>";
			HTML+="The Name:" + TheName + "<br>";
			HTML+="The Email:" + TheEmail + "<br>";
			HTML+="The Subject:" + TheSubject + "<br>";
			HTML+="The Messsgae:" + TheMessage + "<br>";
		HTML+="</body>"
	HTML+="</HTML>"

	SendEmailByCDOReal("HTML",EmailFrom,EmailTo,Subject,HTML);%>
	
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">




<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
    <head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js" type="text/javascript" charset="utf-8"></script>

	<link rel="stylesheet" href="_css/stylesheet.css" type="text/css" charset="utf-8" />


	<title>MWF Group Bahamas</title>
	
	
</head>
<style>
    
#cycler{position:relative;}
#cycler img{position:absolute;z-index:1}
#cycler img.active{z-index:3}

.bigtitle
    {
	font-family:'proxima_nova_ththin';font-size:80px;color:#DDDDDD;padding-bottom:30px;text-align: center;
	}

.mainmenutext
    {
	font-family:'proxima_nova_ltsemibold';font-size:15px;color:#ffffff;
	}
		
.submenutext
    {
	font-family:'proxima_nova_ltsemibold';font-size:15px;color:#84cfff;
	}
		
.catagery
    {
	font-family:'proxima_nova_ltsemibold';font-size:15px;color:#555555;padding-bottom:30px;
	}
.heading
    {
	font-family:'proxima_nova_rgregular';font-size:40px;color:#0274be;padding-bottom:35px;
	}
.bodytext
    {
	font-family:'proxima_nova_rgregular';font-size:15px;color:#333333;line-height:26px;
	}	
.sidetext
    {
	font-family:'proxima_nova_rgregular';font-size:14px;color:#333333;line-height:28px;
	}
.sidetextheader
    {
	font-family:'proxima_nova_rgregular';font-size:14px;color:#585858;line-height:28px;font-weight:bold;
	}	
.copyright
    {
	font-family:'proxima_nova_rgregular';font-size:12px;color:#888888;line-height:26px;
	}		

</style>
<style>
body
 { 
 background-image:url('images/bkg.jpg');
 
 background-attachment:fixed;
 background-position:center; 
 background-position:top; 
 overflow-y:scroll;
 }
</style>
<script language="JavaScript">
function cycleImages(){
      var $active = $('#cycler .active');
      var $next = ($active.next().length > 0) ? $active.next() : $('#cycler img:first');
      $next.css('z-index',2);//move the next image up the pile
      $active.fadeOut(1500,function(){//fade out the top image
	  $active.css('z-index',1).show().removeClass('active');//reset the z-index and unhide the image
          $next.css('z-index',3).addClass('active');//make the next image the top one
      });
    }

$(document).ready(function(){
// run every 7s
setInterval('cycleImages()', 7000);
})
function GoSend()
{
	this.document.Form1.XType.value="S";	
	if (this.document.Form1.TheName.value=="")
	{
		alert("Please enter your name.");
		return;
	}
	if (this.document.Form1.TheEmail.value=="")
	{
		alert("Please enter your email.");
		return;
	}
	if (this.document.Form1.TheSubject.value=="")
	{
		alert("Please enter the subject.");
		return;
	}
	if (this.document.Form1.TheMessage.value=="")
	{
		alert("Please enter the message.");
		return;
	}
	this.document.Form1.submit();
}
</script>



<body topmargin="0" leftmargin="0" bottommargin="0" rightmargin="0">
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="900">
    <tr>
        <td>
        
        <table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="white">
    <tr>
        <td><img src="images/logo.gif" /></td>
        <td><a href="default.htm"><img src="images/btn_home-off.gif"  onmouseup="this.src='images/btn_home-off.gif'" onmousedown="this.src='images/btn_home-on.gif'" onmouseover="this.src='images/btn_home-on.gif'" onmouseout="this.src='images/btn_home-off.gif'" border="0"></a></td>
        
        <td><a href="service.htm"><img src="images/btn_service-off.gif" onmouseup="this.src='images/btn_service-off.gif'" onmousedown="this.src='images/btn_service-on.gif'" onmouseover="this.src='images/btn_service-on.gif'" onmouseout="this.src='images/btn_service-off.gif'" border="0"/></a></td>
        
        <td><a href="aboutus.htm"><img src="images/btn_aboutus-off.gif" onmouseup="this.src='images/btn_aboutus-off.gif'" onmousedown="this.src='images/btn_aboutus-on.gif'" onmouseover="this.src='images/btn_aboutus-on.gif'" onmouseout="this.src='images/btn_aboutus-off.gif'" border="0"/></a></td>
        <td><a href="ourteam.htm"><img src="images/btn_ourteam-off.gif" onmouseup="this.src='images/btn_ourteam-off.gif'" onmousedown="this.src='images/btn_ourteam-on.gif'" onmouseover="this.src='images/btn_ourteam-on.gif'" onmouseout="this.src='images/btn_ourteam-off.gif'" border="0"/></a></td>
        <td><a href="openposition.htm"><img src="images/btn_openposition-off.gif" onmouseup="this.src='images/btn_openposition-off.gif'" onmousedown="this.src='images/btn_openposition-on.gif'" onmouseover="this.src='images/btn_openposition-on.gif'" onmouseout="this.src='images/btn_openposition-off.gif'" border="0"/></a></td>
        
        <td><img src="images/btn_contact-on.gif" border="0"/></a></td>
        
    </tr>
    </table></td>
    </tr>
    <tr>
        <td   valign="top"><img class="active" src="images/contactus.jpg" width="900" height="208"></td>
    </tr>
    <tr>
        <td><img src="images/main2_shadow.gif" /></td>
    </tr>
    <tr>
        <td bgcolor="#FFFFFF" align="center" style="padding-bottom:30px;">
        
        <table border="0" width="100%" cellpadding="0" cellspacing="0">
            <tr>
                
                <td align="left" style="border-right:0px solid rgb(200,200,200);width:450px;padding-left:50px;padding-right:20px;padding-top:20px;padding-bottom:20px;"  valign="top">
                <p></p>
              <div class="sidetextheader" style="text-align:center">We look forward to meeting you!</div>
              <p>
                <div class="sidetextheader"  style="text-align:center">
                	Thank you!
               </div>
                <p>&nbsp;<p>
                <p>&nbsp;<p>
               </td>
             
            </tr>
            <tr>
            <td bgcolor="#dee6e8" height="70" align="center"><div class="copyright">Copyright © 2004 - 2014 by MWF Group. All rights reserved.</div></td>
            </tr>
        </table>
        </td>
    </tr>
    
</table>
</div>
</body>
</html>

	<%Response.end;
}%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">




<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
    <head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js" type="text/javascript" charset="utf-8"></script>

	<link rel="stylesheet" href="_css/stylesheet.css" type="text/css" charset="utf-8" />


	<title>MWF Group Bahamas</title>
	
	
</head>
<style>
    
#cycler{position:relative;}
#cycler img{position:absolute;z-index:1}
#cycler img.active{z-index:3}

.bigtitle
    {
	font-family:'proxima_nova_ththin';font-size:80px;color:#DDDDDD;padding-bottom:30px;text-align: center;
	}

.mainmenutext
    {
	font-family:'proxima_nova_ltsemibold';font-size:15px;color:#ffffff;
	}
		
.submenutext
    {
	font-family:'proxima_nova_ltsemibold';font-size:15px;color:#84cfff;
	}
		
.catagery
    {
	font-family:'proxima_nova_ltsemibold';font-size:15px;color:#555555;padding-bottom:30px;
	}
.heading
    {
	font-family:'proxima_nova_rgregular';font-size:40px;color:#0274be;padding-bottom:35px;
	}
.bodytext
    {
	font-family:'proxima_nova_rgregular';font-size:15px;color:#333333;line-height:26px;
	}	
.sidetext
    {
	font-family:'proxima_nova_rgregular';font-size:14px;color:#333333;line-height:28px;
	}
.sidetextheader
    {
	font-family:'proxima_nova_rgregular';font-size:14px;color:#585858;line-height:28px;font-weight:bold;
	}	
.copyright
    {
	font-family:'proxima_nova_rgregular';font-size:12px;color:#888888;line-height:26px;
	}		

</style>
<style>
body
 { 
 background-image:url('images/bkg.jpg');
 
 background-attachment:fixed;
 background-position:center; 
 background-position:top; 
 overflow-y:scroll;
 }
</style>
<script language="JavaScript">
function cycleImages(){
      var $active = $('#cycler .active');
      var $next = ($active.next().length > 0) ? $active.next() : $('#cycler img:first');
      $next.css('z-index',2);//move the next image up the pile
      $active.fadeOut(1500,function(){//fade out the top image
	  $active.css('z-index',1).show().removeClass('active');//reset the z-index and unhide the image
          $next.css('z-index',3).addClass('active');//make the next image the top one
      });
    }

$(document).ready(function(){
// run every 7s
setInterval('cycleImages()', 7000);
})
function GoSend()
{
	this.document.Form1.XType.value="S";	
	if (this.document.Form1.TheName.value=="")
	{
		alert("Please enter your name.");
		return;
	}
	if (this.document.Form1.TheEmail.value=="")
	{
		alert("Please enter your email.");
		return;
	}
	if (this.document.Form1.TheSubject.value=="")
	{
		alert("Please enter the subject.");
		return;
	}
	if (this.document.Form1.TheMessage.value=="")
	{
		alert("Please enter the message.");
		return;
	}
	this.document.Form1.submit();
}
</script>



<body topmargin="0" leftmargin="0" bottommargin="0" rightmargin="0">
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="900">
    <tr>
        <td>
        
        <table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="white">
    <tr>
        <td><img src="images/logo.gif" /></td>
        <td><a href="default.htm"><img src="images/btn_home-off.gif"  onmouseup="this.src='images/btn_home-off.gif'" onmousedown="this.src='images/btn_home-on.gif'" onmouseover="this.src='images/btn_home-on.gif'" onmouseout="this.src='images/btn_home-off.gif'" border="0"></a></td>
        
        <td><a href="service.htm"><img src="images/btn_service-off.gif" onmouseup="this.src='images/btn_service-off.gif'" onmousedown="this.src='images/btn_service-on.gif'" onmouseover="this.src='images/btn_service-on.gif'" onmouseout="this.src='images/btn_service-off.gif'" border="0"/></a></td>
        
        <td><a href="aboutus.htm"><img src="images/btn_aboutus-off.gif" onmouseup="this.src='images/btn_aboutus-off.gif'" onmousedown="this.src='images/btn_aboutus-on.gif'" onmouseover="this.src='images/btn_aboutus-on.gif'" onmouseout="this.src='images/btn_aboutus-off.gif'" border="0"/></a></td>
        <td><a href="ourteam.htm"><img src="images/btn_ourteam-off.gif" onmouseup="this.src='images/btn_ourteam-off.gif'" onmousedown="this.src='images/btn_ourteam-on.gif'" onmouseover="this.src='images/btn_ourteam-on.gif'" onmouseout="this.src='images/btn_ourteam-off.gif'" border="0"/></a></td>
        <td><a href="openposition.htm"><img src="images/btn_openposition-off.gif" onmouseup="this.src='images/btn_openposition-off.gif'" onmousedown="this.src='images/btn_openposition-on.gif'" onmouseover="this.src='images/btn_openposition-on.gif'" onmouseout="this.src='images/btn_openposition-off.gif'" border="0"/></a></td>
        
        <td><img src="images/btn_contact-on.gif" border="0"/></a></td>
        
    </tr>
    </table></td>
    </tr>
    <tr>
        <td   valign="top"><img class="active" src="images/contactus.jpg" width="900" height="208"></td>
    </tr>
    <tr>
        <td><img src="images/main2_shadow.gif" /></td>
    </tr>
    <tr>
        <td bgcolor="#FFFFFF" align="center" style="padding-bottom:30px;">
        
        <table border="0" width="100%" cellpadding="0" cellspacing="0">
            <tr>
                
                <td align="left" style="border-right:0px solid rgb(200,200,200);width:450px;padding-left:50px;padding-right:20px;padding-top:20px;padding-bottom:20px;"  valign="top">
                <p></p>
              <div class="sidetextheader">We look forward to meeting you!</div>
              <p>
                <div class="sidetext">
                	<form method="POST" action="contact.asp" name="Form1"> 
					<input type="hidden" name="XType" value="">
                <table cellpadding="5">
                <tr>
                <td>Your Name:</td>
                <td>
					<input name="TheName" type="text" value="" style="font-family:Arial;font-size:12px;border:1px solid rgb(200,200,200);padding:3px;width:300px;"></td>
                </tr>
                
                <tr>
                <td>Your Email:</td>
                <td><input name="TheEmail" type="text" value="" style="font-family:Arial;font-size:12px;border:1px solid rgb(200,200,200);padding:3px;width:300px;"></td>
                </tr>
                
                
                <tr>
                <td>Subject:</td>
                <td><input name="TheSubject" type="text" value="" style="font-family:Arial;font-size:12px;border:1px solid rgb(200,200,200);padding:3px;width:300px;"></td>
                </tr>
                
                
                <tr>
                <td valign="top">Message:</td>
                <td><textarea name="TheMessage" type="text" value="" style="border:1px solid rgb(200,200,200);padding:3px;width:600px;height:100px;font-family:Arial;font-size:12px;"></textarea></td>
                </tr>
                
                <tr>
                <td valign="top"></td>
                <td><input type="button" onclick="GoSend()" value="Send" style="font-family:Arial;font-size:12px;color:white;background-color:#f97419;padding:10px;border:1px;"></td>
                </tr>
                
                </table>
                
                </form>
                
               
               </div>
                <p>&nbsp;<p>
                <p>&nbsp;<p>
               </td>
             
            </tr>
            <tr>
            <td bgcolor="#dee6e8" height="70" align="center"><div class="copyright">Copyright © 2004 - 2014 by MWF Group. All rights reserved.</div></td>
            </tr>
        </table>
        </td>
    </tr>
    
</table>
</div>
</body>
</html>
