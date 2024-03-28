<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/kwon2010.asp" -->
<%

dim company_tab(150)
dim area_tab
area_tab = array("서울","경기","부산","대구","인천","광주","대전","울산","강원","경남","경북","세종","충남","충북","전남","전북","제주")
dim as_cnt(16)
dim as_per(16)

'ck_sw=Request("ck_sw")
c_name = "전체"
c_grade = request.cookies("kwon_user")("coo_grade")
user_id = request.cookies("kwon_user")("coo_id")
mg_group = request.cookies("kwon_user")("coo_mg_group")
c_reside = request.cookies("kwon_user")("coo_reside")
user_name = request.cookies("kwon_user")("coo_name")

'If ck_sw = "n" Then
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	company = request.form("company")
'Else
'	from_date=Request("from_date")
'	to_date=Request("to_date")
'	company = "전체"
'End if

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	company = "전체"
End If

'if company = "" then
'	company = "전체"
'end if

if	c_grade = "5" and c_reside = "0" then
	c_name = request.cookies("kwon_user")("coo_name")
	company = c_name
end if

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if c_name = "전체" then
	k = 0
	company_tab(0) = "전체"
	if	( c_grade = "5" and c_reside = "1" ) then
		Sql="select * from k1_etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' and group_name = '"+user_name+"' order by etc_name asc"
		  else
		Sql="select * from k1_etc_code where etc_type = '51' and used_sw = 'Y' and mg_group = '"+mg_group+"' order by etc_name asc"
	end if
	Rs_etc.Open Sql, Dbconn, 1
	while not rs_etc.eof
		k = k + 1
		company_tab(k) = rs_etc("etc_name")
		rs_etc.movenext()
	Wend
rs_etc.close()						
end if				

grade_sql = ""
if ( c_grade = "5" and c_reside = "1" ) then
	com_sql = "company = '" + company_tab(1) + "'"	
	for kk = 2 to k
		com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
	next
	grade_sql = " and (" + com_sql + ")"
end if

kkk = k

if company = "전체" then
	if  ( c_grade = "5" and c_reside = "1" ) then
		sql = "select count(*) as err_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') " + grade_sql
	  else 
		sql = "select count(*) as err_tot from k1_as_acpt "
		sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
	end if
  else
	sql = "select count(*) as err_tot from k1_as_acpt "
	sql = sql + "WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
	sql = sql + " and company = '" + company + "'"
end if

Rs.Open Sql, Dbconn, 1
err_tot = cint(rs("err_tot"))
if rs.eof then
	err_tot = 0
end if

rs.close()
for i = 0 to 16            
	sido = area_tab(i)
	if company = "전체" then
		if  ( c_grade = "5" and c_reside = "1" ) then
			sql = "select sido,COUNT(*) AS err_cnt FROM k1_as_acpt" 
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')" + grade_sql
			sql = sql + " GROUP BY sido"
			sql = sql + " HAVING (sido = '"+sido+"')"
		  else 
			sql = "select sido,COUNT(*) AS err_cnt FROM k1_as_acpt" 
			sql = sql + " WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
			sql = sql + " GROUP BY sido"
			sql = sql + " HAVING (sido = '"+sido+"')"
		end if
	  else
		sql = "select company,sido,COUNT(*) AS err_cnt FROM k1_as_acpt" 
		sql = sql + " WHERE (mg_group='"+mg_group+"') and (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		sql = sql + " GROUP BY company,sido"
		sql = sql + " HAVING (company = '"+company+"') AND (sido = '"+sido+"')"
	end if
			 
	Rs.Open Sql, Dbconn, 1

	if rs.eof then
		as_cnt(i) = 0
		as_per(i) = 0
		else
		as_cnt(i) = cint(rs("err_cnt"))
		as_per(i) = formatnumber((as_cnt(i)/err_tot * 100),2)
	end if
	rs.close()

next
		

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="include/kwon_style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/java/PopupCalendar.js"></script>
<script>
<!--
function auto_submit()
{
 window.setTimeout("Submit_Function()", 300000);
 return true;
}

function Submit_Function() 
{
document.form1.submit();
}

function MM_openBrWindow(theURL,winName,features) 
{ 
  window.open(theURL,winName,features);
}
//-->
</script>
<title></title>
</head>
<body>
<table width="900" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" height="30" bgcolor="#6699CC">&nbsp;<span class="style14BW">지역별 통계 현황</span></td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellspacing="3">
      <tr>
        <td><form name="form1" method="post" action="k1_waiting.asp?pg_name=k1_area_sum.asp">
          <table width="100%"  border="0">
            <tr>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr valign="middle" class="style12">
                  <td width="10%" height="25" bgcolor="#CCCCCC"><div align="center" class="style6">접수기간</div></td>
                  <td width="45%" height="25">&nbsp;
                      <input name="from_date" type="text" id="from_date2" size="10" value=<%=from_date%>>
                      <span class="style5">
                      <input name="button" type="button" onClick="popUpCalendar(this, from_date, 'yyyy-mm-dd')" value="달력">
                      </span>~
                      <input name="to_date" type="text" id="to_date2" size="10" value=<%=to_date%>>
                      <span class="style5">
                      <input name="button2" type="button" onClick="popUpCalendar(this, to_date, 'yyyy-mm-dd')" value="달력">
                    </span></td>
                  <td width="10%" height="25" bgcolor="#CCCCCC"><div align="center" class="style6">회사</div></td>
                  <td width="25%" height="25">&nbsp;
                      <%
		if c_name = "전체" then
		%>
                      <select name="company" id="company">
                        <% 
					for kk = 0 to kkk
			  	%>
                        <option value='<%=company_tab(kk)%>' <%If company_tab(kk) = company then %>selected<% end if %>><%=company_tab(kk)%></option>
                        <%
					next
				%>
                      </select>
                      <% else %>
                      <%=c_name%>
                      <% end if %>
                  </td>
                  <td width="10%" height="25"><div align="center">
                      <input name="imageField" type="image" src="image/burton/view01.gif" width="55" height="20">
                  </div></td>
                </tr>
              </table></td>
              </tr>
            <tr>
              <td><table width="100%" border="1" cellspacing="0" cellpadding="0">
                <tr valign="middle" bgcolor="#CCFFCC" class="style12">
                  <td width="15%" height="22"><div align="center" class="style7 style8">시도</div></td>
                  <td width="60%" height="20" bgcolor="#CCFFCC"><div align="center" class="style6"></div>
                      <div align="center" class="style6">그 래 프</div>
                      <div align="center" class="style6"></div></td>
                  <td width="10%" height="20"><div align="center" class="style6">건수</div></td>
                  <td width="15%" height="25"><div align="center" class="style6">백분율</div>
                      <div align="center" class="style6"></div></td>
                </tr>
              </table>
                <table width="100%" height="22"  border="1" cellpadding="0" cellspacing="0">
                  <%
		for i = 0 to 16
		%>
                  <tr class="style6">
                    <td width="15%" height="20" bgcolor="#FFFFCC"><div align="center" class="style12"><%=area_tab(i)%></div></td>
                    <td width="60%" height="20" bgcolor="#FFFFFF"><div align="left"><span class="style7">&nbsp;<img src="image/graph02.gif" width="<%=as_per(i)%>%" height="13" align="center"></span></div></td>
                    <td width="10%" height="20" class="style12"><div align="center"><span class="style7"> <a  href="#" onClick="MM_openBrWindow('k1_area_detail.asp?sido=<%=area_tab(i)%>&company=<%=company%>&err_tot=<%=as_cnt(i)%>&from_date=<%=from_date%>&to_date=<%=to_date%>','area_detail_popup','scrollbars=yes,width=840,height=500')"><%=formatnumber(clng(as_cnt(i)),0)%></a></span></div></td>
                    <td width="15%" height="20" class="style12"><div align="center"><%=as_per(i)%>%</div></td>
                  </tr>
                  <%
		next
		%>
                </table>
                <table width="100%"  border="1" cellpadding="0" cellspacing="0">
                  <tr bgcolor="#D6DDEF">
                    <td width="15%" height="22"><div align="center" class="style12B">총계</div></td>
                    <td width="60%">&nbsp;</td>
                    <td width="10%" class="style12B"><div align="center"><%=formatnumber(clng(err_tot),0)%></div></td>
                    <td width="15%" class="style12B"><div align="center">&nbsp;</div></td>
                  </tr>
                </table></td>
              </tr>
          </table>
          </form></td>
      </tr>
    </table>      </td>
  </tr>
</table>
</body>
</html>
<%
dbconn.Close()
Set dbconn = Nothing
%>
