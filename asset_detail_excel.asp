<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

company=Request("company")
field_view=Request("field_view")
field_check=Request("field_check")

If field_check = "total" Then
	field_view = ""
End If
curr_date = mid(now(),1,10)

savefilename = "자산대장" + curr_date + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_dept = Server.CreateObject("ADODB.Recordset")
Set rs_group = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

com_sql = " and asset.company = '" + company + "' "

if field_check = "total" then
	condi_sql = ""
  else
	condi_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

order_sql = " order by asset_dept.org_first , asset_dept.org_second , asset_dept.dept_name , asset.gubun , asset.code_seq "

Sql = "SELECT * FROM asset inner join asset_dept on (asset.company = asset_dept.company) and (asset.dept_code = asset_dept.dept_code) where asset.dept_code > '0' and (inst_process = 'Y') " + com_sql + condi_sql + order_sql
Rs.Open Sql, Dbconn, 1

if company = "01" then
	title_01 = "법인명"
	title_02 = "지사명"
	title_03 = "지점명"
  else
	title_01 = "조직명1"
	title_02 = "조직명2"
	title_03 = "조직명3"
end if

etc_code = "75" + company
Sql="select * from etc_code where etc_code = '" + etc_code + "'"
Set rs_etc=DbConn.Execute(SQL)
if rs_etc.eof or rs_etc.bof then
	company_name = "없음"
  else
	company_name = rs_etc("etc_name")
end if
rs_etc.close()						

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
													
<html>
<head>
<title>자산 세부 내역</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<body>
<table width="1200" border="0">
  <tr> 
    <td width="100%" height="30" bgcolor="#FFFFFF" class="style14BW style15">&nbsp;* 자산대장 (<%=now()%>기준) </td>
  </tr>
  <tr> 
    <td width="100%"><form action="k1_org_asset.asp?ck_sw=n" method="post" name="form3">
      <table width="100%">
        <tr>
          <td>            <table  border="1" cellpadding="0" cellspacing="0">
            <tr bgcolor="#EFEFEF" class="style12">
              <td width="50" height="30" bgcolor="#EFEFEF"><div align="center" class="style12">순번</div></td>
              <td width="100" bgcolor="#EFEFEF"><div align="center">소속회사</div></td>
              <td width="100" height="30" bgcolor="#EFEFEF"><div align="center">관리조직</div></td>
              <td width="150" height="30" bgcolor="#EFEFEF"><div align="center"><%=title_01%></div></td>
              <td width="150" height="30" bgcolor="#EFEFEF"><div align="center"><%=title_02%></div></td>
              <td width="100" height="30" bgcolor="#EFEFEF"><div align="center"><%=title_03%></div></td>
              <td width="80" height="30" class="style12"><div align="center">자산코드</div></td>
              <td width="120" height="30" class="style12"><div align="center">자산명</div></td>
              <td width="100" height="30" class="style12"><div align="center">자산번호</div></td>
              <td width="100" height="30" class="style12"><div align="center">시리얼NO</div></td>
              <td width="80" height="30" class="style12"><div align="center">사용자</div></td>
              <td width="80" height="30" class="style12"><div align="center">인터넷NO</div></td>
              <td width="80" height="30" class="style12"><div align="center">발송일자</div></td>
              <td width="80" height="30" class="style12"><div align="center">설치일자</div></td>
              </tr>
        <%
		k = 0
		do until rs.eof
			k = k + 1
			internet_no = "."
			if rs("gubun") = "01" then
				sql = "select * from asset_dept where company = '" + rs("company") + "' and dept_code = '" + rs("dept_code") + "'"
				Set rs_dept=DbConn.Execute(SQL)
				if rs_dept.eof or rs_dept.bof then
					internet_no = "없음"
				  else
					internet_no = rs_dept("internet_no")
				end if
			end if
	   	%>
            <tr valign="middle" class="style12">
              <td height="25"><div align="center" class="style12"><%=k%></div></td>
              <td height="25"><div align="center"><%=company_name%></div></td>
              <td height="25"><div align="center" class="style12"><%=rs("high_org")%></div></td>
              <td height="25"><div align="center"><%=rs("org_first")%></div></td>
              <td height="25"><div align="center"><%=rs("org_second")%></div></td>
              <td height="25"><div align="center"><%=rs("dept_name")%></div></td>
              <td height="25"><div align="center"><%=rs("company")%>-<%=rs("gubun")%>-<%=rs("code_seq")%></div></td>
              <td height="25"><div align="center"><%=rs("asset_name")%></div></td>
              <td height="25"><div align="center"><%=mid(rs("asset_no"),1,2)%>-<%=mid(rs("asset_no"),3,6)%>-<%=right(rs("asset_no"),4)%></div></td>
              <td height="25"><div align="center"><%=rs("serial_no")%></div></td>
              <td height="25"><div align="center"><%=rs("user_name")%></div></td>
              <td height="25"><div align="center"><%=internet_no%></div></td>
              <td height="25"><div align="center"><%=rs("send_date")%></div></td>
              <td height="25"><div align="center"><%=rs("install_date")%></div></td>
              </tr>
            <% 
			rs.movenext()
		loop
		%>
          </table></td></tr>
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
