<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/itft2005.asp" -->

<%
user_name = request.Cookies("itft_user")("coo_name")
retire_pay_month=Request("retire_pay_month")

retire_pay_date = mid(cstr(retire_pay_month),1,4) + "-" + mid(cstr(retire_pay_month),5,2) + "-01"
retire_pay_date = dateadd("m",1,retire_pay_date)
retire_pay_date = dateadd("d",-1,retire_pay_date)
enter_date = dateadd("yyyy",-1,retire_pay_date)
enter_date = dateadd("d",1,enter_date)
from_date = dateadd("m",-2,retire_pay_date)
from_month = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
to_month = mid(cstr(retire_pay_date),1,4) + mid(cstr(retire_pay_date),6,2)
Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_pay = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from insa where (retir_date < '2005-02-01' and cast(enter_date as date) <= '"+cstr(enter_date)+"') order by emp_no asc"
Rs.Open Sql, Dbconn, 1
do until rs.eof	
' 1년에 한번 처리
'	if rs("emp_no") <> "200501" and rs("emp_no") <> "201311" then
	if rs("emp_no") <> "201311" then
		yy = 0
		mm = 0
		dd = 0
		
		sql = "select sum(base_pay) as base_pay,sum(jikmu) as jikmu,sum(jikchek) as jikchek,sum(sikdae) as sikdae,sum(over_time) as over_time,sum(incentive) as incentive,sum(etc_pay) as etc_pay"
		sql = sql + " from pay where pay_month >= '" + from_month + "' and pay_month <= '" + to_month + "' and emp_no = '" + rs("emp_no") + "'"
		Set rs_pay = Dbconn.Execute (sql)
'		tot_pay = clng(rs_pay("base_pay"))+clng(rs_pay("jikmu"))+clng(rs_pay("jikchek"))+clng(rs_pay("sikdae"))+clng(rs_pay("over_time"))+clng(rs_pay("incentive"))+clng(rs_pay("etc_pay"))
		tot_pay = clng(rs_pay("base_pay"))+clng(rs_pay("jikmu"))+clng(rs_pay("jikchek"))+clng(rs_pay("sikdae"))+clng(rs_pay("over_time"))+clng(rs_pay("etc_pay"))
	
		average_pay = int(tot_pay / 3)
		if rs("emp_no") = "200501" then
			average_pay = average_pay * 12 * 3 / 10
		end if
	
		if rs("retir_pay_date") = "1900-01-01" or rs("retir_pay_date") < rs("enter_date") then				
			com_sw = "1"
			yy = 1
			mm = 0
			dd = datediff("d",rs("enter_date"),enter_date)
			retire_pay = average_pay + int(dd*average_pay/365)
		  else	
			be_pay_date = dateadd("m",-1,retire_pay_date)
			if be_pay_date = rs("retir_pay_date") then
				com_sw = "2"
				yy = 0
				mm = 1
				dd = 0
				retire_pay = int(average_pay / 12)
			  else
				com_sw = "3"
				yy = 0
				mm = 0
				dd = datediff("d",rs("retir_pay_date"),retire_pay_date)
				retire_pay = int(average_pay / 12) * int(dd/30)
			end if
		 end if
	
		sql="insert into retire_pay (retire_pay_month,emp_no,before_pay_date,retire_pay_date,average_pay,work_year,work_month,work_day,retire_pay,pay_yn,reg_user,reg_date) values ('"&retire_pay_month&"','"&rs("emp_no")&"','"&rs("retir_pay_date")&"','"&retire_pay_date&"',"&average_pay&","&yy&","&mm&","&dd&","&retire_pay&",'N','"&user_name&"',now())"
		dbconn.execute(sql)
	end if
	rs.movenext()	
loop

url = "retire_pay_mg.asp?retire_pay_month="+retire_pay_month
response.write"<script language=javascript>"
response.write"alert('퇴직금 생성이 완료되었습니다...');"		
response.write"location.replace('"&url&"');"
response.write"</script>"
Response.End
%>
