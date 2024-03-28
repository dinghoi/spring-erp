<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim ce_sum(1000,3,11)
dim ce_tab(1000,3)

for i = 1 to 1000
	ce_tab(i,1) = ""
	ce_tab(i,2) = ""
	ce_tab(i,2) = ""
next

for i = 1 to 1000
	for j = 1 to 3
		for k = 1 to 11
			ce_sum(i,j,k) = 0
		next
	next
next
	 
slip_month="201408"
cost_year = mid(slip_month,1,4)
cost_month = mid(slip_month,5)
from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
Set rs_sign = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' 일반비용
sql = "select emp_company,bonbu,saupbu,team,org_name,reside_place,company,account,sum(cost) as cost from general_cost where (slip_gubun = '비용') and (cancel_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,bonbu,saupbu,team,org_name,reside_place,company,account"
rs.Open sql, Dbconn, 1
do until rs.eof
			
	sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
	"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and reside_place ='"&rs("reside_place")& _
	"' and company ='"&rs("company")&"' and cost_id ='일반경비' and cost_detail ='"&rs("account")&"'"
	set rs_cost=dbconn.execute(sql)

	if rs_cost.eof or rs_cost.bof then
		sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,reside_place,company,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','"&rs("reside_place")&"','"&rs("company")&"','일반경비','"&rs("account")&"',"&rs("cost")&")"
		dbconn.execute(sql)
	  else
		sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and reside_place ='"&rs("reside_place")&"' and company ='"&rs("company")&"' and cost_id ='일반경비' and cost_detail ='"&rs("account")&"'"
		dbconn.execute(sql)
	end if		
	rs.movenext()
loop
rs.close()

' 임차료
sql = "select emp_company,bonbu,saupbu,team,org_name,reside_place,company,account,sum(cost) as cost from general_cost where (slip_gubun = '임차료') and (cancel_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,bonbu,saupbu,team,org_name,reside_place,company,account"
rs.Open sql, Dbconn, 1
do until rs.eof
			
	sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
	"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and reside_place ='"&rs("reside_place")& _
	"' and company ='"&rs("company")&"' and cost_id ='임차료' and cost_detail ='"&rs("account")&"'"
	set rs_cost=dbconn.execute(sql)

	if rs_cost.eof or rs_cost.bof then
		sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,reside_place,company,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','"&rs("reside_place")&"','"&rs("company")&"','임차료','"&rs("account")&"',"&rs("cost")&")"
		dbconn.execute(sql)
	  else
		sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and reside_place ='"&rs("reside_place")&"' and company ='"&rs("company")&"' and cost_id ='임차료' and cost_detail ='"&rs("account")&"'"
		dbconn.execute(sql)
	end if		
	rs.movenext()
loop
rs.close()

response.write"<script language=javascript>"
response.write"alert('처리되었습니다');"
response.write"parent.opener.location.reload();"
'response.write"self.close() ;"
'response.write"</script>"
Response.End

dbconn.Close()
Set dbconn = Nothing

%>
