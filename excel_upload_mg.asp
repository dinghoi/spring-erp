<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/srvmg_dbcon_db.asp" -->
<% 
	
	Set DbConn = Server.CreateObject("ADODB.Connection")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect

%>

<html>
<head>
<title></title>
<style type="text/css">
<!--
.style3 {font-size: 12px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; }
.style4 {font-size: 12px; font-family: "굴림체", "돋움체", Seoul; }
-->
</style>
</head>
<body>
<table width="800" border="0">
  <tr> 
    <td width="800" height="41"><img src="image/k1_excel_upload_mg_title.gif" width="800" height="40"></td>
  </tr>
  <tr> 
    <td width="800" height="6"><table width="800" height="30"  border="1" cellpadding="0" cellspacing="0">
      <tr>
        <td width="180" height="50" align="center" valign="middle" bgcolor="#CCFFFF" class="style3"><div align="center">1. 업로드 EXCEL 파일 선택 </div></td>
        <td height="50" valign="middle"><input name="att_file" type="file" id="att_file2" size="60"></td>
        <td width="70" height="50"><div align="center">
		<a href="k1_file_upload_ok.asp"><img src="image/burton/upbtn.gif" width="55" height="20" border="0"></a></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="3"><div align="center">
		</div></td>
  </tr>
</table>
</body>
</html>
