<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/itft2005_db.asp" -->

<%
Dim Rs
Dim Repeat_Rows

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_wait = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "SELECT count(acpt_no) as wait_no  FROM as_acpt WHERE ( reside_place = '콜센터' ) and ( as_type = '원격처리' ) and ( as_process = '접수' )"
Rs_wait.Open Sql, Dbconn, 1
wait_no = rs_wait("wait_no") + 1

Sql = "SELECT * FROM as_acpt WHERE ( reside_place = '콜센터' ) and ( as_type = '원격처리' ) and ( as_process = '접수' or as_process = '확인' ) ORDER BY acpt_no DESC"
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="include/itft_style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style13 {color: #003366}
.style14 {font-family: "굴림", "돋움", Seoul, "한강체"}
.style15 {font-family: "굴림체", "돋움체", Seoul}
.style22 {font-family: Arial, Seoul}
-->
</style>
</head>

<body>
<table width="650" border="0">
  <tr> 
    <td width="800" height="41"><img src="image/as_acpt_view_title.gif" width="650" height="40"></td>
  </tr>
  <tr> 
    <td width="800"><form action="" method="post" name="form1">
      <table width="650"  border="1" cellpadding="0" cellspacing="0">
        <tr valign="middle" bgcolor="#EFEFEF" class="style11B">
          <td width="120" height="28"><div align="center">접수일자</div></td>
          <td width="55"><div align="center">접수NO</div></td>
          <td width="50"><div align="center">접수</div></td>
          <td width="50"><div align="center">상태</div></td>
          <td width="60"><div align="center">사용자</div></td>
          <td width="70"><div align="center">회사</div></td>
          <td width="135"><div align="center">조직명</div></td>
          <td width="40"><div align="center">대기</div></td>
          <td width="50"><div align="center">FAQ</div></td>
        </tr>
        <%
  	Repeat_rows = 10 
	Repeat_index = 0
	While ((Repeat_Rows <> 0) AND (NOT Rs.EOF)) 
    %>
        <tr valign="middle" class="style11">
    <%
	int date_len 
	date_len=len(rs("acpt_date"))
	as_memo = rs("as_memo")
	if rs("acpt_man") = "인터넷" then
		acpt_type = "인터넷"
	  else
	  	acpt_type = "전화"
	end if
	if rs("as_process") = "확인" then
		wait_no_view = 0
		as_process = "원격중"
	  else
		wait_no = wait_no - 1
		wait_no_view = wait_no
	  	as_process = "접수"
	end if
	%>
          <td width="120" height="27"><div align="center"><%=mid(cstr(rs("acpt_date")),3,date_len-5)%></div></td>
          <td width="55" height="27"><div align="center"><%=rs("acpt_no")%></div></td>
          <td width="50" height="27"><div align="center"><%=acpt_type%></div></td>
          <td width="50" height="27"><div align="center"><%=as_process%></div></td>
          <td width="60" height="27" class="style13"><div align="center">
            <p style="cursor:pointer"><span title="<%=as_memo%>"><%=rs("acpt_user")%></span></p>
          </div></td>
          <td width="70" height="27"><div align="center"><%=rs("company")%></div></td>
          <td width="135" height="27"><div align="center"><%=rs("dept")%></div></td>
          <td width="40" height="27"><div align="center"><%=wait_no_view%></div></td>
          <td width="50" height="27"><div align="center">1 / 1 </div></td>
        </tr>
        <% 
  	Repeat_index=Repeat_index+1
'	Repeat1__numRows=Repeat1__numRows-1
	Repeat_rows = Repeat_rows - 1
	Rs.MoveNext()
	Wend
%>
        <%
intstart = (int((page-1)/10)*10) + 1
intend = intstart + 9
first_page = 1

if intend > total_page then
	intend = total_page
end if
%>
      </table>
      <table width="650">
        <tr>
          <td height="33"><span class="style1"></span>
              <div align="center" class="style12"></div></td>
        </tr>
      </table>
    </form>	</td>
  </tr>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
