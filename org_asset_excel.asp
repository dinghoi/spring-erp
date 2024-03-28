<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim asset_cnt(1000,6)
dim asset_dept(1000)
dim asset_tot(6)
for i = 0 to 1000
	asset_dept(i) = "N"
	for j = 0 to 6
		asset_cnt(i,j) = 0
	next
next
for j = 0 to 6
	asset_tot(j) = 0
next

company=Request("company")
field_view=Request("field_view")
field_check=Request("field_check")

If field_check = "total" Then
	field_view = ""
End If
curr_date = mid(now(),1,10)

savefilename = "조직별 자산 현황" + curr_date + ".xls"

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

base_sql = "SELECT asset.dept_code, asset.gubun, count(*) as asset_cnt, asset_dept.high_org, asset_dept.org_first, asset_dept.org_second, asset_dept.dept_name FROM asset INNER JOIN asset_dept ON (asset.dept_code = asset_dept.dept_code) AND (asset.company = asset_dept.company) where asset.inst_process = 'Y' and asset.company = '" + company + "'"
if field_check <> "total" then
	condi_sql = " and ( asset_dept." + field_check + " like '%" + field_view + "%' ) "
  else
  	condi_sql = " "
end if	

group_sql = " group by asset.dept_code, asset.gubun order by asset.dept_code"
sql = base_sql + condi_sql + group_sql
rs_group.Open Sql, Dbconn, 1
bi_dept = "y"
i = 0
do until rs_group.eof
	if	bi_dept = "y" then
		bi_dept = rs_group("dept_code")
	end if
	if  bi_dept <> rs_group("dept_code") then
		i = i + 1
		asset_dept(i) = bi_dept
		for j = 1 to 6
			asset_cnt(i,j) = asset_cnt(i,j) + asset_tot(j)
			asset_cnt(0,j) = asset_cnt(0,j) + asset_tot(j)
			asset_tot(j) = 0
		next		
		bi_dept = rs_group("dept_code")
	end if
	jj = cint(rs_group("gubun"))
	if jj > 4 then
		jj = 5
	end if
	asset_tot(jj) = asset_tot(jj) + cint(rs_group("asset_cnt"))
	asset_tot(6) = asset_tot(6) + cint(rs_group("asset_cnt"))	
	rs_group.movenext()
loop
if  i > 0 then
	i = i + 1
	asset_dept(i) = bi_dept
	for j = 1 to 6
		asset_cnt(i,j) = asset_cnt(i,j) + asset_tot(j)
		asset_cnt(0,j) = asset_cnt(0,j) + asset_tot(j)
	next		
end if

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
<title>조직별 자산 현황</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>

<body>
<table width="1000" border="0">
  <tr> 
    <td width="100%" height="30" bgcolor="#FFFFFF" class="style14BW style15">&nbsp;조직별 자산 현황 </td>
  </tr>
  <tr> 
    <td width="100%"><form action="k1_org_asset.asp?ck_sw=n" method="post" name="form3">
      <table width="100%">
        <tr>
          <td>            <table width="100%"  border="1" cellpadding="0" cellspacing="0">
            <tr bgcolor="#EFEFEF" class="style12">
              <td width="5%" height="30" bgcolor="#EFEFEF"><div align="center" class="style12">순번</div></td>
              <td width="9%" bgcolor="#EFEFEF"><div align="center">소속회사</div></td>
              <td width="8%" height="30" bgcolor="#EFEFEF"><div align="center">관리조직</div></td>
              <td width="14%" height="30" bgcolor="#EFEFEF"><div align="center"><%=title_01%></div></td>
              <td width="14%" height="30" bgcolor="#EFEFEF"><div align="center"><%=title_02%></div></td>
              <td width="14%" height="30" bgcolor="#EFEFEF"><div align="center"><%=title_03%></div></td>
              <td width="6%" height="30" class="style12"><div align="center">데탑</div></td>
              <td width="6%" height="30" class="style12"><div align="center">모니터</div></td>
              <td width="6%" height="30" class="style12"><div align="center">노트북</div></td>
              <td width="6%" height="30" class="style12"><div align="center">프린터</div></td>
              <td width="6%" height="30" class="style12"><div align="center">기타</div></td>
              <td width="6%" height="30" class="style12"><div align="center">소계</div></td>
            </tr>
            <%
		for k = 1 to i
			if asset_dept(k) = "N" then
				exit for
			end if

			Sql="select * from asset_dept where company = '" + company + "' and dept_code = '" + asset_dept(k) + "'"
			Set rs_dept=DbConn.Execute(SQL)
			if rs_dept.eof or rs_dept.bof then
				high_org = "없음"
				org_first = "없음"
				org_second = "없음"
				dept_name = "없음"
			  else
				high_org = rs_dept("high_org")
				org_first = rs_dept("org_first")
				org_second = rs_dept("org_second")
				dept_name = rs_dept("dept_name")
			end if
			rs_dept.close()						
			
    	%>
            <tr valign="middle" class="style12">
              <td width="5%" height="25"><div align="center" class="style12"><%=k%></div></td>
              <td width="9%" height="25"><div align="center"><%=company_name%></div></td>
              <td width="8%" height="25"><div align="center" class="style12"><%=high_org%></div></td>
              <td width="14%" height="25"><div align="center"><%=org_first%></div></td>
              <td width="14%" height="25"><div align="center"><%=org_second%></div></td>
              <td width="14%" height="25"><div align="center"><%=dept_name%></div></td>
              <td width="6%" height="25"><div align="right"><%=formatnumber(clng(asset_cnt(k,1)),0)%></div></td>
              <td width="6%" height="25"><div align="right"><%=formatnumber(clng(asset_cnt(k,2)),0)%></div></td>
              <td width="6%" height="25"><div align="right"><%=formatnumber(clng(asset_cnt(k,3)),0)%></div></td>
              <td width="6%" height="25"><div align="right"><%=formatnumber(clng(asset_cnt(k,4)),0)%></div></td>
              <td width="6%" height="25"><div align="right"><%=formatnumber(clng(asset_cnt(k,5)),0)%></div></td>
              <td width="6%" height="25"><div align="right"><%=formatnumber(clng(asset_cnt(k,6)),0)%></div></td>
            </tr>
            <% 
		next
		%>
          </table></td></tr>
      </table>
    </form>	</td>
  </tr>
</table>
</body>
</html>
<%
Rs_group.Close()
Set Rs_group = Nothing
%>
