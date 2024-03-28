<!--#include virtual="/common/inc_top.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
<HEAD>
<TITLE>popup</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
-->
</style>

<script type="text/javascript">

	function setCookie(name, value, expiredays){
		var today = new Date();
		today.setDate(today.getDate() + expiredays);

		document.cookie = name + '=' + escape(value) + '; path=/; expires=' + today.toGMTString() + ';'
	}

	function closePop(){
		if(document.forms[0].todayPop.checked)
			setCookie('nkp_notice', 'rangs', 1);

		self.close();
	}

	 function closewin(){
	   var expire = new Date();

	   expire.setDate(expire.getDate() - 1);
	   document.cookie = "ww2=1; expires=" + expire.toGMTString()+ "; path=/";
	   self.close();
	 }
</script>

</HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<!-- ImageReady Slices (popup.psd) -->
<form name="frm" method="post" action="">
  <TABLE WIDTH="400" BORDER="0" CELLPADDING="0" CELLSPACING="0">
    <TR>
      <TD COLSPAN=2>
			<IMG SRC="/image/nkp_popup1.png" WIDTH="645" HEIGHT="629" ALT=""></TD>
    </TR>
    <TR>
      <TD width="585" height="25" valign="middle"><div align="right"><span class="style1"><strong>닫 기</strong></span></div></TD>
      <TD width="50" height="25" align="center" valign="middle"><input name="todayPop" type="checkbox" id="todayPop" onClick="closePop()" value="checkbox"> </TD>
    </TR>
  </TABLE>
</form>
<!-- End ImageReady Slices -->
</BODY>
</HTML>