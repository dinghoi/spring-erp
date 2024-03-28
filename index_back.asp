<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<%
save_id = request.cookies("kwon_id_check")
%>
<HTML>
<HEAD>
<TITLE>케이원정보통신 서비스 전산관리</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {
	font-size: 12px;
	font-family: "굴림체", "돋움체", Seoul;
}
-->
</style>
</HEAD>
<BODY BGCOLOR=#FFFFFF LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<!-- ImageReady Slices (index_01.psd) -->
<form name="form1" method="post" action="login.asp">
  <table width="100%"  height="100%" border="0">
    <tr>
      <td align="center" valign="middle"><TABLE WIDTH=729 BORDER=0 CELLPADDING=0 CELLSPACING=0>
	<TR>
		<TD COLSPAN=8>
			<IMG SRC="images/index_01_01.gif" WIDTH=729 HEIGHT=117 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=117 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=3 ROWSPAN=4>
			<IMG SRC="images/index_01_02.gif" WIDTH=51 HEIGHT=43 ALT=""></TD>
		<TD COLSPAN=2 ROWSPAN=2 align="left" valign="middle" bgcolor="#E4E4E4" class="style1">
		  <input name="id" type="text" id="id3"  style="width:85px;height:20px" tabindex="1" value="<%=Request.Cookies("kwon_id")%>">
		</TD>
		<TD ROWSPAN=4>
			<IMG SRC="images/index_01_04.gif" WIDTH=6 HEIGHT=43 ALT=""></TD>
		<TD>
			<IMG SRC="images/index_01_05.gif" WIDTH=31 HEIGHT=15 ALT=""></TD>
		<TD ROWSPAN=9>
			<IMG SRC="images/index_01_06.gif" WIDTH=549 HEIGHT=185 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=15 ALT=""></TD>
	</TR>
	<TR>
		<TD ROWSPAN=4>
			<input name="login" type="image" SRC="images/index_01_07.gif" WIDTH=31 HEIGHT=36 ALT="" tabindex="3"></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=11 ALT=""></TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="images/index_01_08.gif" WIDTH=8 HEIGHT=8 ALT=""></TD>
		<TD>
			<IMG SRC="images/index_01_09.gif" WIDTH=84 HEIGHT=8 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=8 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=2 ROWSPAN=3 align="left" valign="middle" bgcolor="#E4E4E4" class="style1">
		  <input name="pass" type="password" id="pass" tabindex="2" style="width:85px;height:20px">
		</TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=9 ALT=""></TD>
	</TR>
	<TR>
		<TD ROWSPAN=5>
			<IMG SRC="images/index_01_11.gif" WIDTH=21 HEIGHT=142 ALT=""></TD>
		<TD COLSPAN=2 ROWSPAN=2>
			<IMG SRC="images/index_01_12.gif" WIDTH=30 HEIGHT=17 ALT=""></TD>
		<TD ROWSPAN=2>
			<IMG SRC="images/index_01_13.gif" WIDTH=6 HEIGHT=17 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=8 ALT=""></TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="images/index_01_14.gif" WIDTH=31 HEIGHT=9 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=9 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=6>
			<IMG SRC="images/index_01_15.gif" WIDTH=159 HEIGHT=10 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=10 ALT=""></TD>
	</TR>
	<TR>
		<TD ROWSPAN=2>
			<IMG SRC="images/index_01_16.gif" WIDTH=8 HEIGHT=115 ALT=""></TD>
		<TD COLSPAN=5 align="left" valign="middle" bgcolor="#E4E4E4" class="style1"><input name="save_id" type="checkbox" id="save_id" value="1" <%If save_id = "1" then %>checked<% end if %>> 
		  아이디저장 </TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=23 ALT=""></TD>
	</TR>
	<TR>
		<TD COLSPAN=5>
			<IMG SRC="images/index_01_18.gif" WIDTH=151 HEIGHT=92 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=92 ALT=""></TD>
	</TR>
	<TR>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=21 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=8 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=22 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=8 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=84 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=6 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=31 HEIGHT=1 ALT=""></TD>
		<TD>
			<IMG SRC="images/spacer.gif" WIDTH=549 HEIGHT=1 ALT=""></TD>
		<TD></TD>
	</TR>
      </TABLE></td>
    </tr>
  </table>
</form>
<!-- End ImageReady Slices -->
</BODY>
</HTML>