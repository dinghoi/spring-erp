<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

Server.ScriptTimeOut = 500

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

' ���� 5%, �湮 95%
won_per = 5
bang_per = 95

for i = 1 to 10
	saupbu_tab(i,1) = ""
	saupbu_tab(i,2) = 0
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from cost_end where end_month = '"&end_month&"' and (end_yn = 'Y') and (saupbu = '���ֺ��')"
Set rs_check = Dbconn.Execute (sql)	
if rs_check.eof or rs_check.bof then
	check_sw = "N"
  else
  	check_sw = "Y"
end if

if check_sw = "N" then
	response.write"<script language=javascript>"
	response.write"alert('���ֺ�� ������ �����ϼž� �մϴ� !!');"
	response.write"location.replace('cost_end_mg.asp');"
	response.write"</script>"
	Response.End
  else		
	response.write"<script language=javascript>"
	response.write"alert('����ó����!!!');"
	response.write"</script>"

dbconn.BeginTrans

sql = "select sum(cost_amt_"&mm&") as tot_cost from company_cost where cost_year ='"&cost_year&"' and cost_center = '�ι������'"
Set rs=DbConn.Execute(SQL)
tot_cost = clng(rs("tot_cost"))
rs.close()

' ����
sql = " select count(*) as tot_cnt from as_acpt Where acpt_man in ('���μ�','�ֿ���','�Ѽ���','����ȯ') and (Cast(visit_date as date) >= '" + from_date + "' and Cast(visit_date as date) <= '"+to_date+"')and company not in('�ڿ���','������ũ��','������ǰ','�Ե���Ż')"  
Set rs=DbConn.Execute(SQL)
won_cnt = clng(rs("tot_cnt"))
if won_cnt = "" or isnull(tot_cnt) then
	won_cnt = 0
end if
rs.close()

sql = " select company, count(*) as as_cnt from as_acpt Where acpt_man in ('���μ�','�ֿ���','�Ѽ���','����ȯ') and (Cast(visit_date as date) >= '" + from_date + "' and Cast(visit_date as date) <= '"+to_date+"') and company not in('�ڿ���','������ũ��','������ǰ','�Ե���Ż') GROUP BY company Order By company Asc"
Rs.Open Sql, Dbconn, 1

do until rs.eof 

	sql = "select saupbu from trade where trade_name = '"&rs("company")&"'" 
	set rs_trade=dbconn.execute(sql)
	if rs_trade.eof or rs_trade.bof then
		saupbu = "Error"
	  else
	  	saupbu = rs_trade("saupbu")
	end if
	rs_trade.close()
	
	charge_per = clng(rs("as_cnt")) / won_cnt * won_per / 100
	cost_amt = int(charge_per * tot_cost)
	sql="insert into company_as (as_month,as_company,saupbu,remote_cnt,charge_per,cost_amt,reg_id,reg_name,reg_date) values ('"&end_month&"','"&rs("company")&"','"&saupbu&"','"&rs("as_cnt")&"','"&charge_per&"',"&cost_amt&",'"&user_id&"','"&user_name&"',now())"
	dbconn.execute(sql)
	
	rs.movenext()
loop	
rs.close()

' ���ݿ�
sql = " select count(*) as tot_cnt "
sql = sql & " from as_acpt a inner join emp_master_month b on a.mg_ce_id=b.emp_no and b.emp_month='" & end_month & "'"
sql = sql & " Where (as_type <> '����ó��' and as_process <> '���' and as_type <> '��Ư��') "
'sql = sql & " and ( team not  like '%����%') and (company <> '��Ÿ' and company <> '���̿��������') "
'sql = sql & " and  mg_ce not in ('�Ķ�','�ھ���','��ÿ�','������','������','������') "
sql = sql & " and reside='0'  and reside_place=' ' "
sql = sql & " and (Cast(visit_date as date) >= '" + from_date + "' and Cast(visit_date as date) <= '"+to_date+"')"
sql = sql & " and b.cost_center='�ι������' "

Set rs=DbConn.Execute(SQL)
bang_cnt = clng(rs("tot_cnt"))
if bang_cnt = "" or isnull(tot_cnt) then
	bang_cnt = 0
end if
rs.close()

sql = " select company, count(*) as as_cnt "
sql = sql & " from as_acpt a inner join emp_master_month b on a.mg_ce_id=b.emp_no and b.emp_month='" & end_month & "'"
sql = sql & " Where (as_type <> '����ó��' and as_process <> '���' and as_type <> '��Ư��') "
'sql = sql & " and ( team not  like '%����%') and (company <> '��Ÿ' and company <> '���̿��������') "
'sql = sql & " and  mg_ce not in ('�Ķ�','�ھ���','��ÿ�','������','������','������') 
sql = sql & " and reside='0' and reside_place=' ' "
sql = sql & " and (Cast(visit_date as date) >= '" + from_date + "' and Cast(visit_date as date) <= '"+to_date+"') "
sql = sql & " and b.cost_center='�ι������' "
sql = sql & " GROUP BY company Order By company Asc"
Rs.Open Sql, Dbconn, 1

do until rs.eof 

	sql = "select saupbu from trade where trade_name = '"&rs("company")&"'" 
	set rs_trade=dbconn.execute(sql)
	if rs_trade.eof or rs_trade.bof then
		saupbu = "Error"
	  else
	  	saupbu = rs_trade("saupbu")
	end if
	rs_trade.close()
	
	sql = "select * from company_as where as_month = '"&end_month&"' and as_company = '"&rs("company")&"'"
	set rs_etc=dbconn.execute(sql)
	if rs_etc.eof or rs_etc.bof then
		charge_per = clng(rs("as_cnt")) / bang_cnt * bang_per / 100
		cost_amt = int(charge_per * tot_cost)
		sql="INSERT INTO company_as (as_month,as_company,saupbu,visit_cnt,charge_per,cost_amt,reg_id,reg_name,reg_date) "&_
		    " VALUES ('"&end_month&"','"&rs("company")&"','"&saupbu&"','"&rs("as_cnt")&"','"&charge_per&"',"&cost_amt&",'"&user_id&"','"&user_name&"',now())"
		dbconn.execute(sql)
	else
		charge_per = clng(rs("as_cnt")) / bang_cnt * bang_per / 100 + rs_etc("charge_per")
		cost_amt = int(charge_per * tot_cost)
		sql = "UPDATE company_as SET visit_cnt='"&rs("as_cnt")&"', charge_per='"&charge_per&"', cost_amt="&cost_amt&_
		      " WHERE as_company='" &rs("company")& "' and as_month = '" &end_month& "'"
		dbconn.execute(sql)	  
	end if
	
	rs.movenext()
loop	
rs.close()
rs_etc.close()
' ����κ� ���� �ڷ� ����
' �ι������ ���
' ó���� zero
sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='�ι������') "
dbconn.execute(sql)

sql = " select saupbu, sum(charge_per) as charge_per from company_as Where (as_month = '"&end_month&"') GROUP BY saupbu"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	charge_per = rs("charge_per")

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '�ι������' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * clng(rs_etc("cost")))

		sql = "select * from saupbu_profit_loss where cost_year ='"&cost_year&"' and saupbu ='"&rs("saupbu")&"' and cost_center ='�ι������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("saupbu")&"','�ι������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and saupbu ='"&rs("saupbu")&"' and cost_center ='�ι������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()
' �κа���� ��� ��

' ���纰 ���� �ڷ� ����
' �ι������ ���
' ó���� zero
sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='�ι������') "
dbconn.execute(sql)

sql = " select as_company as company, sum(charge_per) as charge_per from company_as Where (as_month = '"&end_month&"') GROUP BY as_company"
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

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '�ι������' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * clng(rs_etc("cost")))

		sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='�ι������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','�ι������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='�ι������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()
' �κа���� ��� ��


' �߰� ����
' ����κ� �ο��� ����
sql = " select saupbu from sales_org where sales_year='" & cost_year & "' order by saupbu asc"
Rs.Open Sql, Dbconn, 1
i = 0
tot_person = 0
do until rs.eof 

	'sql = "select count(*) from pay_month_give where pmg_id = '1' and pmg_yymm = '"&end_month&"' and mg_saupbu ='"&rs("saupbu")&"'"
	
	''
	'' KDC����� ��� ����ó��
	''
	
	' KDC����ο� �̸��� ���� ���� ���̿�������ſ� �Ҽӵ� ����� cost_except = '2' �� �����Ѵ�.
	sql = "SELECT emp_name, count(*)               "&chr(13)&_
        "  FROM                                  "&chr(13)&_
        "(                                       "&chr(13)&_
        "  SELECT B.*                            "&chr(13)&_
        "    FROM pay_month_give A               "&chr(13)&_
        "        ,emp_master_month B             "&chr(13)&_
        "   WHERE A.pmg_id = '1'                 "&chr(13)&_
        "     AND A.pmg_emp_no = B.emp_no        "&chr(13)&_
        "     AND B.cost_except in ('0','1','2') "&chr(13)&_
        "     AND A.pmg_yymm  = '"&end_month&"'  "&chr(13)&_
        "     AND B.emp_month = '"&end_month&"'  "&chr(13)&_
        "     AND A.mg_saupbu = 'KDC�����'      "&chr(13)&_
        ") A                                     "&chr(13)&_
        "GROUP BY emp_name                       "&chr(13)&_
        "  HAVING count(*) = 2                   "
  'Response.write "<pre>"& sql &"</pre><br>"
  set rs_emp = dbconn.execute(sql)
  do until rs_emp.eof 
    emp_name = rs_emp("emp_name")
    
    sql = "UPDATE emp_master_month              "&chr(13)&_    
          "   SET cost_except = '2'             "&chr(13)&_
          " WHERE emp_name    = '"&emp_name&"'  "&chr(13)&_
          "   AND emp_month   = '"&end_month&"' "&chr(13)&_
          "   AND emp_company = '���̿��������'"
    'Response.write "<pre>"& sql &"</pre><br>"
    dbconn.execute(sql)
     
    rs_emp.movenext()
  loop  
  rs_emp.close()
  
  ' KDC����ο� ���������� �ش��ϴ� ����� cost_except = '2' �� �����Ѵ�.
	sql = "UPDATE emp_master_month                               "&chr(13)&_
        "   SET cost_except = '2'                              "&chr(13)&_
        " WHERE emp_month   = '"&end_month&"'                  "&chr(13)&_
        "   AND cost_center = '����������'                     "&chr(13)&_
        "   AND emp_saupbu  = 'KDC�����'                      "&chr(13)&_
        "   AND emp_no IN ( SELECT pmg_emp_no                  "&chr(13)&_
        "                     FROM pay_month_give              "&chr(13)&_
        "                    WHERE pmg_id    = 1               "&chr(13)&_
        "                      AND pmg_yymm  = '"&end_month&"' "&chr(13)&_
        "                      AND mg_saupbu ='KDC�����'      "&chr(13)&_
        "                 )                                    "
  'Response.write "<pre>"& sql &"</pre><br>"        
  dbconn.execute(sql)

	''
	'' KDC����� ��� ����ó�� _ ��
	''

	
	'����� ��α��� ���� ó��(2016-01-15)
	sql = "select count(*) from pay_month_give  A ,emp_master_month B "
	sql = sql & "where A.pmg_id = '1'  "
	sql = sql & "and A.pmg_yymm = '"&end_month&"' "
	sql = sql & "and A.mg_saupbu ='"&rs("saupbu")&"' "
	sql = sql & "and A.pmg_emp_no=  B.emp_no "
	sql = sql & "and B.cost_except in('0','1') "
	sql = sql & "and B.emp_month ='"&end_month&"' "
	
	'Response.write sql&"<br>"

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

'�������� �Ѿ�
sql = "select sum(cost_amt_"&mm&") as tot_cost from company_cost where cost_year ='"&cost_year&"' and cost_center = '��������'"
Set rs=DbConn.Execute(SQL)
tot_cost_amt = clng(rs("tot_cost"))

rs.close()

' ����κ� ���� �ڷ� ����
' ó���� zero
sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='��������') "
dbconn.execute(sql)
sql = "delete from management_cost where cost_month ='"&end_month&"'"
'Response.write sql&"<br>"
dbconn.execute(sql)


' �������� ���

for i = 1 to 10
	if saupbu_tab(i,1) = "" or isnull(saupbu_tab(i,1)) then
		exit for
	end if

' ����� ���� �Ѿ�
	sql = "select sum(cost_amt) from saupbu_sales where substring(sales_date,1,7) = '"&sales_month&"' and saupbu ='"&saupbu_tab(i,1)&"'"
	set rs_cost=dbconn.execute(sql)
	if rs_cost(0) = "" or isnull(rs_cost(0)) then
		saupbu_sales = 0
	  else
		saupbu_sales = CCur(rs_cost(0))
	end if
	rs_cost.close()

	saupbu_per = saupbu_tab(i,2) / tot_person
	saupbu_cost_amt = int(tot_cost_amt * saupbu_per)

sql = "select company,sum(cost_amt) as cost from saupbu_sales where substring(sales_date,1,7) = '"&sales_month&"' and saupbu ='"&saupbu_tab(i,1)&"' group by saupbu,company"
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
		
		sql = "insert into management_cost (cost_month,saupbu,company,tot_person,saupbu_person,saupbu_per,tot_cost_amt,saupbu_cost_amt,charge_per,cost_amt,reg_id,reg_name,reg_date) values ('"&end_month&"','"&saupbu_tab(i,1)&"','"&rs_etc("company")&"',"&tot_person&","&saupbu_tab(i,2)&","&saupbu_per&","&tot_cost_amt&","&saupbu_cost_amt&","&charge_per&","&cost_amt&",'"&user_Id&"','"&user_name&"',now())"
		'Response.write sql&"<br>"
		dbconn.execute(sql)
		rs_etc.movenext()
	loop
' ������ ������ ���
	if k = 0 then
		sql = "insert into management_cost (cost_month,saupbu,company,tot_person,saupbu_person,saupbu_per,tot_cost_amt,saupbu_cost_amt,charge_per,cost_amt,reg_id,reg_name,reg_date) values ('"&end_month&"','"&saupbu_tab(i,1)&"','',"&tot_person&","&saupbu_tab(i,2)&","&saupbu_per&","&tot_cost_amt&","&saupbu_cost_amt&",1,"&saupbu_cost_amt&",'"&user_Id&"','"&user_name&"',now())"
		'Response.write sql&"<br>"
		dbconn.execute(sql)
	end if
	rs_etc.close()



	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '��������' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof

		cost = int(saupbu_per * clng(rs_etc("cost")))

		sql = "select * from saupbu_profit_loss where cost_year ='"&cost_year&"' and saupbu ='"&saupbu_tab(i,1)&"' and cost_center ='��������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
			
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&saupbu_tab(i,1)&"','��������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and saupbu ='"&saupbu_tab(i,1)&"' and cost_center ='��������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		

		rs_etc.movenext()
	loop
	rs_etc.close()
next
' �������� ��� ��

' ���纰 ���� �ڷ� ����
' �������� ���
' ó���� zero
sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='��������') "
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

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '��������' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * clng(rs_etc("cost")) * rs("saupbu_per"))

		sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='��������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','��������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='��������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()

' ���纰 ������ ���
' ó���� zero
sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='������') "
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

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '������' ) and (saupbu = '"&rs("saupbu")&"' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * Cdbl(rs_etc("cost")))

		sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','������','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='������' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()
' ���纰 ������ ��� ��

if end_yn = "C" then
	sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
	"' and saupbu = '�����/��������'"
  else
	sql="insert into cost_end (end_month,saupbu,end_yn,batch_yn,bonbu_yn,ceo_yn,reg_id,reg_name,reg_date) values ('"&end_month& _
	"','�����/��������','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
end if
dbconn.execute(sql)

if Err.number <> 0 then
	dbconn.RollbackTrans 
	end_msg = emp_msg + "ó���� Error�� �߻��Ͽ����ϴ�...."
else    
	dbconn.CommitTrans
	end_msg = emp_msg + "����ó�� �Ǿ����ϴ�...."
end if

response.write"<script language=javascript>"
response.write"alert('"&end_msg&"');"
response.write"location.replace('cost_end_mg.asp');"
response.write"</script>"
Response.End

dbconn.Close()
Set dbconn = Nothing

end if
%>