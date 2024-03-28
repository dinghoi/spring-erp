<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
<HEAD>
<TITLE>popup</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
-->
</style>
<Script>

function setCookie(name, value, expiredays) {
var today = new Date();
    today.setDate(today.getDate() + expiredays);

    document.cookie = name + '=' + escape(value) + '; path=/; expires=' + today.toGMTString() + ';'
}

function closePop() {        
if(document.forms[0].todayPop.checked)                
	setCookie('popup', 'rangs', 1);
self.close();
}

 function closewin(){
   var expire = new Date();
   expire.setDate(expire.getDate() - 1);
   document.cookie = "ww2=1; expires=" + expire.toGMTString()+ "; path=/";
   self.close();
 }
</Script>

</HEAD>
<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<!-- ImageReady Slices (popup.psd) -->
<form name="frm" method="post" action="">
  <TABLE WIDTH=400 BORDER=0 CELLPADDING=0 CELLSPACING=0>
    <TR>
      <TD COLSPAN=2>
			<IMG SRC="images/popup_01.gif" WIDTH=400 HEIGHT=455 ALT=""></TD>
    </TR>
    <TR>
      <TD width="300" height="25" valign="middle"><div align="right"><span class="style1"> 오늘 하루 그만 보기</span></div></TD>
      <TD width="100" height="25" valign="middle"><input name="todayPop" type="checkbox" id="todayPop" onClick="closePop()" value="checkbox"> </TD>
    </TR>
  </TABLE>
</form>
<!-- End ImageReady Slices -->
</BODY>
</HTML>