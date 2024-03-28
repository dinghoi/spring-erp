<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

dim saupbu_tab(10,2)

end_month=Request("end_month")
end_yn=Request("end_yn")

from_date = mid(end_month,1,4) + "-" + mid(end_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

mm = mid(end_month,5,2)
cost_year = mid(end_month,1,4)
cost_month = mid(end_month,5)

for i = 1 to 10
	saupbu_tab(i,1) = ""
	saupbu_tab(i,2) = 0
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

response.write"<script language=javascript>"
response.write"alert('마감처리중!!!');"
response.write"</script>"

dbconn.BeginTrans
' 사업부별 인원수 집계
sql = "select * from sales_org where sales_year='" & cost_year & "' Order By saupbu Asc"

Rs.Open Sql, Dbconn, 1
i = 0
tot_person = 0
do until rs.eof 

	sql = "select count(*) from emp_master_month where emp_month = '"&end_month&"' and emp_saupbu ='"&rs("saupbu")&"'" 
	'//2016-08-31 정직원만 인원수 집계
	'sql = sql & " and emp_type='정직' "
	set rs_emp=dbconn.execute(sql)
	if rs_emp(0) = "" or isnull(rs_emp(0)) then
		saupbu_person = 0
	  else
		saupbu_person = clng(rs_emp(0))
	end if
	rs_emp.close()
	i = i + 1
	saupbu_tab(i,1) = rs("saupbu")
	saupbu_tab(i,2) = saupbu_person	
	tot_person = tot_person + saupbu_person
	
	rs.movenext()
loop	
rs.close()

'전사공통비 총액
sql = "select sum(cost_amt_"&mm&") as tot_cost from company_cost where cost_year ='"&cost_year&"' and cost_center = '전사공통비'"
Set rs=DbConn.Execute(SQL)
tot_cost_amt = clng(rs("tot_cost"))
rs.close()

' 사업부별 손익 자료 생성
' 처리전 zero
sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='전사공통비') "
dbconn.execute(sql)
sql = "delete from management_cost where cost_month ='"&end_month&"'"
dbconn.execute(sql)

' 전사공통비 배부
for i = 1 to 10
	if saupbu_tab(i,1) = "" or isnull(saupbu_tab(i,1)) then
		exit for
	end if

' 사업부 매출 총액
	sql = "select sum(sales_amt) from saupbu_sales where sales_month = '"&end_month&"' and saupbu ='"&saupbu_tab(i,1)&"'"
	set rs_cost=dbconn.execute(sql)
	if rs_cost(0) = "" or isnull(rs_cost(0)) then
		saupbu_sales = 0
	  else
		saupbu_sales = clng(rs_cost(0))
	end if
	rs_cost.close()

	saupbu_per = saupbu_tab(i,2) / tot_person
	saupbu_cost_amt = int(tot_cost_amt * saupbu_per)
	
	sql = "select company,sales_amt as cost from saupbu_sales where sales_month = '"&end_month&"' and saupbu ='"&saupbu_tab(i,1)&"'"
	rs_etc.Open sql, Dbconn, 1
	
	k = 0
	do until rs_etc.eof
		k = k + 1
		if saupbu_sales = 0 then
			charge_per = 0
		  else
			charge_per = rs_etc("cost") / saupbu_sales
		end if
		cost_amt = int(charge_per * saupbu_cost_amt)
		
		sql = "INSERT INTO      	" &_
		      " management_cost 	" &_
		      "(                	" &_
		      "   cost_month    	" &_
		      " , saupbu        	" &_
		      " , company					" &_
		      " , tot_person			" &_
		      " , saupbu_person		" &_
		      " , saupbu_per			" &_
		      " , tot_cost_amt		" &_
		      " , saupbu_cost_amt	" &_
		      " , charge_per			" &_
		      " , cost_amt				" &_
		      " , reg_id					" &_
		      " , reg_name				" &_
		      " , reg_date				" &_
		      ") 									" &_
		      "VALUES 						" &_
		      "(              		" &_
		      "   '"&end_month&"'					" &_
		      " , '"&saupbu_tab(i,1)&"'		" &_
		      " , '"&rs_etc("company")&"'	" &_
		      " , "&tot_person							&_
		      " , "&saupbu_tab(i,2)					&_
		      " , "&saupbu_per							&_
		      " , "&tot_cost_amt						&_
		      " , "&saupbu_cost_amt					&_
		      " , "&charge_per							&_
		      " , "&cost_amt								&_
		      " , '"&user_Id								&_
		      " , '"&user_name&"'					" &_
		      " , now()" 										&_
		      ")"
		dbconn.execute(sql)
		rs_etc.movenext()
	loop
' 매출이 제로인 경우
	if k = 0 then
		sql = "INSERT INTO        			"&_
		      "management_cost    			"&_
		      "(                  			"&_
		      "   cost_month      			"&_
		      " , saupbu          			"&_
		      " , company         			"&_
		      " , tot_person      			"&_
		      " , saupbu_person   			"&_
		      " , saupbu_per      			"&_
		      " , tot_cost_amt    			"&_
		      " , saupbu_cost_amt 			"&_
		      " , charge_per      			"&_
		      " , cost_amt        			"&_
		      " , reg_id          			"&_
		      " , reg_name        			"&_
		      " , reg_date)       			"&_
		      " values            			"&_
		      "(  '"&end_month&"' 			"&_
		      " , '"&saupbu_tab(i,1)&"' "&_
		      " , '' 										"&_
		      " , "&tot_person 					 &_
		      " , "&saupbu_tab(i,2)			 &_
		      " , "&saupbu_per 					 &_
		      " , "&tot_cost_amt				 &_
		      " , "&saupbu_cost_amt			 &_
		      " , 1 										"&_
		      " , "&saupbu_cost_amt			 &_
		      " , '"&user_Id&"' 				"&_
		      " , '"&user_name&"' 			"&_
		      " , now() 								"&_
		      ")"
		dbconn.execute(sql)
	end if
	rs_etc.close()



	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '전사공통비' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof

		cost = int(saupbu_per * clng(rs_etc("cost")))

		sql = "select * from saupbu_profit_loss where cost_year ='"&cost_year&"' and saupbu ='"&saupbu_tab(i,1)&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
			
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&saupbu_tab(i,1)&"','전사공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and saupbu ='"&saupbu_tab(i,1)&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		

		rs_etc.movenext()
	loop
	rs_etc.close()
next
' 전사공통비 배부 끝

' 고객사별 손익 자료 생성
' 전사공통비 배부
' 처리전 zero
sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='전사공통비') "
dbconn.execute(sql)

sql = " select company,saupbu_per, sum(charge_per) as charge_per from management_cost Where (cost_month = '"&end_month&"') GROUP BY company"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	charge_per = rs("charge_per")

	sql = "select * from trade where trade_name = '"&rs("company")&"'"
	set rs_trade=dbconn.execute(sql)
	if rs_trade.eof or rs_trade.bof then
		group_name = "Error"
	  else
		group_name = rs_trade("group_name")
	end if		  

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '전사공통비' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * clng(rs_etc("cost")) * rs("saupbu_per"))

		sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','전사공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()

' 고객사별 직접비 배부
' 처리전 zero
sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='직접비') "
dbconn.execute(sql)

sql = " select saupbu,company, sum(charge_per) as charge_per from management_cost Where (cost_month = '"&end_month&"') GROUP BY saupbu,company"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	charge_per = rs("charge_per")

	sql = "select * from trade where trade_name = '"&rs("company")&"'"
	set rs_trade=dbconn.execute(sql)
	if rs_trade.eof or rs_trade.bof then
		group_name = "Error"
	  else
		group_name = rs_trade("group_name")
	end if		  

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '직접비' ) and (saupbu = '"&rs("saupbu")&"' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * clng(rs_etc("cost")))

		sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='직접비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','직접비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='직접비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()
' 고객사별 직접비 배부 끝


if end_yn = "C" then
	sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
	"' and saupbu = '공통비배분'"
  else
	sql="insert into cost_end (end_month,saupbu,end_yn,batch_yn,bonbu_yn,ceo_yn,reg_id,reg_name,reg_date) values ('"&end_month& _
	"','공통비배분','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
end if
dbconn.execute(sql)

if Err.number <> 0 then
	dbconn.RollbackTrans 
	end_msg = emp_msg + "처리중 Error가 발생하였습니다...."
else    
	dbconn.CommitTrans
	end_msg = emp_msg + "마감처리 되었습니다...."
end if

response.write"<script language=javascript>"
response.write"alert('"&end_msg&"');"
response.write"location.replace('cost_end_mg.asp');"
response.write"</script>"
Response.End

dbconn.Close()
Set dbconn = Nothing

%>

