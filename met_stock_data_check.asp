<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

view_condi = "���̿��������"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

j = 0

'����â�� ���
'Sql = "select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_level = 'ȸ��') ORDER BY org_company,org_code ASC"

'��â����(���̿� �켱)
'Sql = "select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_level = '��') and (org_company = '"+view_condi+"') ORDER BY org_company,org_code ASC"

'��â����(���̿��� ������ �׷�� ����� : ���̸��ߺ� ����)
Sql = "select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_level = '��') ORDER BY org_company,org_code ASC"

'���� ����ó�� â��� ��Ͻ� ���(����ó�� ���� ���� ������� �̸��ߺ� üũ�Ұ�)
'Sql = "select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01' or org_end_date = '0000-00-00') and (org_level = '����ó') and (org_company = '"+view_condi+"') ORDER BY org_company,org_code ASC"


Rs.Open Sql, Dbconn, 1
if not Rs.eof then
   do until Rs.eof

    org_code = rs("org_code")
	org_level = rs("org_level")
	if org_level = "ȸ��" then
	       org_level = "����"
	end if
	org_company = rs("org_company")
	org_bonbu = rs("org_bonbu")
	org_saupbu = rs("org_saupbu")
	org_team = rs("org_team")
	org_name = rs("org_name")
	org_empno = rs("org_empno")
	org_emp_name = rs("org_emp_name")
	org_date = rs("org_date")
	stock_end_date = "1900-01-01"
	
    if org_level = "����" or org_level = "��" then 
	
	   Sql = "select * from met_stock_code where stock_name = '"+org_name+"'"
	   Set Rs_stock = DbConn.Execute(SQL)
	   if  Rs_stock.EOF or Rs_stock.BOF then

'	Sql = "select * from emp_master where emp_no = '"+org_empno+"'"
'	Set Rs_emp = DbConn.Execute(SQL)
'	if not Rs_emp.EOF or not Rs_emp.BOF then
'	        emp_grade = rs_emp("emp_grade")
'		    emp_position = rs_emp("emp_position")
'		    emp_company = rs_emp("emp_company")
'			emp_bonbu = rs_emp("emp_bonbu")
'			emp_saupbu = rs_emp("emp_saupbu")
'			emp_team = rs_emp("emp_team")
'			emp_org_code = rs_emp("emp_org_code")
'			emp_org_name = rs_emp("emp_org_name")
'			emp_reside_place = rs_emp("emp_reside_place")
'			emp_reside_company = rs_emp("emp_reside_company")

            j = j + 1
		   
	        sql = "insert into met_stock_code (stock_code,stock_level,stock_name,stock_company,stock_bonbu,stock_saupbu,stock_team,stock_open_date,stock_end_date,stock_manager_code,stock_manager_name"
		    sql = sql + ",reg_date,reg_user) values "
		    sql = sql + " ('"&org_code&"','"&org_level&"','"&org_name&"','"&org_company&"','"&org_bonbu&"','"&org_saupbu&"','"&org_team&"','"&org_date&"','"&stock_end_date&"','"&org_empno&"','"&org_emp_name&"',now(),'"&user_name&"')"        
			
			dbconn.execute(sql)	
			
		 end if	
	end if	 
'	    Rs_emp.close()	
	    Rs.MoveNext()
  loop		
		response.write"<script language=javascript>"
		response.write"alert('â�� �����Ͱ� ����� �����ϴ�......"&j&"');"		
		response.write"location.replace('met_goods_code_mg.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert('��ϵ� ���� �����̾����ϴ�...');"		
		response.write"location.replace('met_goods_code_mg.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
