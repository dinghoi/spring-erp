<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	reg_user = request.cookies("nkpmg_user")("coo_user_name")
	mod_user = request.cookies("nkpmg_user")("coo_user_name")
    user_id = request.cookies("nkpmg_user")("coo_user_id")

    be_pg = request.form("be_pg")
	app_date = request.form("app_date")
	app_id = request.form("app_id")
	
	be_month = cstr(mid(app_date,1,4)) + cstr(mid(app_date,6,2))
	
	emp_no = request.form("emp_no")
	emp_name = request.form("emp_name")
	app_grade = request.form("app_grade")
	app_position = request.form("app_position")
	app_job = request.form("app_job")
	app_to_company = request.form("app_to_company")
	app_to_bonbu = request.form("app_to_bonbu")
	app_to_saupbu = request.form("app_to_saupbu")
	app_to_team = request.form("app_to_team")
	app_org = request.form("app_org")
	app_org_name = request.form("app_org_name")
	
' �̵��߷� accept
	if app_id = "�̵��߷�" then
	    sms_msg = emp_no + "-" + emp_name + "- �̵��߷�"
    	app_be_orgcode = request.form("app_be_orgcode")
	    app_be_org = request.form("app_be_org")
	    app_company = request.form("app_company")
		app_bonbu = request.form("app_bonbu")
		app_saupbu = request.form("app_saupbu")
		app_team = request.form("app_team")
	    app_mv_comment = request.form("app_mv_comment")
		emp_stay_code = request.form("emp_stay_code")
		app_reside_place = request.form("app_reside_place")
		app_reside_company = request.form("app_reside_company")
		stay_name = request.form("stay_name")
		app_jikmu = request.form("emp_jikmu")
		cost_center = request.form("cost_center")
		app_org_level = request.form("app_org_level")
		cost_group = request.form("app_cost_group")
		if app_org_level = "����ó" then
	          reside = "1"
	       else 
	          reside = "0"
        end if
	end if
' �����߷� accept	
    if app_id = "�����߷�" then
	    sms_msg = emp_no + "-" + emp_name + "- �����߷�"
		app_be_enddate = request.form("app_date")
	   ' app_be_enddate = request.form("app_be_enddate")
	    app_end_type = request.form("app_end_type")
	    app_end_comment = request.form("app_end_comment")
	end if	
' �����߷� accept	
    if app_id = "�����߷�" then	
	    sms_msg = emp_no + "-" + emp_name + "- �����߷�"
	    app_be_grade = request.form("app_be_grade")
	    app_gr_type = request.form("app_gr_type")
	    app_gr_comment = request.form("app_gr_comment")
	end if	
' ��å���� accept	
	if app_id = "��å����" then	
	    sms_msg = emp_no + "-" + emp_name + "- ��å����"	
	    app_be_position = request.form("app_be_position")
	    app_bm_company = request.form("app_bm_company")
	    app_bm_bonbu = request.form("app_bm_bonbu")
	    app_bm_saupbu = request.form("app_bm_saupbu")
	    app_bm_team = request.form("app_bm_team")
	    app_bm_orgcode = request.form("app_bm_orgcode")
	    app_bm_org = request.form("app_bm_org")
		app_bm_reside_place = request.form("app_bm_reside_place")
		app_bm_reside_company = request.form("app_bm_reside_company")
	    app_bm_comment = request.form("app_bm_comment")
		app_bm_org_level = request.form("app_bm_org_level")
		if app_bm_org_level = "����ó" then
	          reside = "1"
	       else 
	          reside = "0"
        end if
	end if
' ��å���� accept	
    if app_id = "��å����" then	
	    sms_msg = emp_no + "-" + emp_name + "- ��å����"		
	    app_hm_type = request.form("app_hm_type")
	    app_hm_comment = request.form("app_hm_comment")
	end if	
' �����߷� accept
    if app_id = "�����߷�" then	
	    sms_msg = emp_no + "-" + emp_name + "- �����߷�"				
	    app_hu_type = request.form("app_hu_type")
	    app_hustart_date = request.form("app_hustart_date")
	    app_hufinish_date = request.form("app_hufinish_date")	
	    app_hu_comment = request.form("app_hu_comment")
	end if	
' ¡��߷� accept
    if app_id = "¡��߷�" then	
	    sms_msg = emp_no + "-" + emp_name + "- ¡��߷�"				
	    app_di_type = request.form("app_di_type")
	    app_distart_date = request.form("app_distart_date")
	    app_difinish_date = request.form("app_difinish_date")	
	    app_di_comment = request.form("app_di_comment")	
	end if	
' ����߷� accept
    if app_id = "����߷�" then	
	    sms_msg = emp_no + "-" + emp_name + "- ����߷�"				
	    app_rw_type = request.form("app_rw_type")
	    app_rw_comment = request.form("app_rw_comment")
	end if		

' db update and insert....

	set dbconn = server.CreateObject("adodb.connection")
	
    Set Rs = Server.CreateObject("ADODB.Recordset")
    Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_memb = Server.CreateObject("ADODB.Recordset")
	Set Rs_sawo = Server.CreateObject("ADODB.Recordset")
	Set Rs_mon = Server.CreateObject("ADODB.Recordset")
    Set rs_max = Server.CreateObject("ADODB.Recordset")
	Set rs_stock = Server.CreateObject("ADODB.Recordset")
	
	dbconn.open dbconnect

	dbconn.BeginTrans

	if app_id = "�̵��߷�" then
	
'       �λ縶���� ��������
        sql = "select * from emp_master where emp_no = '"&emp_no&"'"
		set rs_emp = dbconn.execute(sql)
		
		if	rs_emp.eof or rs_emp.bof then
		    response.write"<script language=javascript>"
			response.write"alert('��ϵ� ������ �ƴմϴ�....');"		
			response.write"self.opener.location.reload();"		
	        response.write"window.close();"		
	        response.write"</script>"
	        Response.End
	        dbconn.Close()	
		    else 
     	    sql = "update emp_master set emp_jikmu ='"+app_jikmu+"',emp_company ='"+app_company+"',emp_bonbu ='"+app_bonbu+"',emp_saupbu ='"+app_saupbu+"',emp_team ='"+app_team+"',emp_org_code ='"+app_be_orgcode+"',emp_org_name ='"+app_be_org+"',emp_org_baldate ='"+app_date+"',emp_reside_place ='"+app_reside_place+"',emp_stay_code ='"+emp_stay_code+"',emp_reside_company ='"+app_reside_company+"',cost_center ='"+cost_center+"',cost_group ='"+cost_group+"',emp_mod_user = '"+mod_user+"',emp_mod_date = now() where emp_no = '"&emp_no&"'"
	
	        dbconn.execute(sql)

'       memb ����				
            sql="select * from memb where user_id='"&emp_no&"'"
	        set rs_memb=dbconn.execute(sql)			
		    if not rs_memb.eof then
		       sql = "update memb set emp_company='"&app_company&"',bonbu='"&app_bonbu&"',saupbu='"&app_saupbu&"',team='"&app_team&"',org_name='"&app_be_org&"',reside_place='"&app_reside_place&"',reside_company='"&app_reside_company&"',reside='"&reside&"',insa_grade='1',pay_grade='1' where user_id='"&emp_no&"'"

		       'response.write sql
		
		        dbconn.execute(sql)	 
		    end if
			
'       â���ڵ� ����				
            sql="select * from met_stock_code where stock_code='"&emp_no&"'"
	        set rs_stock=dbconn.execute(sql)			
		    if not rs_stock.eof then
		       sql = "update met_stock_code set stock_company='"&app_company&"',stock_bonbu='"&app_bonbu&"',stock_saupbu='"&app_saupbu&"',stock_team='"&app_team&"' where stock_code='"&emp_no&"'"

		       'response.write sql
		
		        dbconn.execute(sql)	 
		    end if			
			
		end if

'       �λ縶���� ��������(���� �߷� ����)
        sql = "select * from emp_master_month where emp_month ='"&be_month&"' and emp_no = '"&emp_no&"'"
		set Rs_mon = dbconn.execute(sql)
		
		if	not Rs_mon.eof then
     	    sql = "update emp_master_month set emp_jikmu ='"+app_jikmu+"',emp_company ='"+app_company+"',emp_bonbu ='"+app_bonbu+"',emp_saupbu ='"+app_saupbu+"',emp_team ='"+app_team+"',emp_org_code ='"+app_be_orgcode+"',emp_org_name ='"+app_be_org+"',emp_org_baldate ='"+app_date+"',emp_reside_place ='"+app_reside_place+"',emp_stay_code ='"+emp_stay_code+"',emp_reside_company ='"+app_reside_company+"',cost_center ='"+cost_center+"',cost_group ='"+cost_group+"',emp_mod_user = '"+mod_user+"',emp_mod_date = now() where emp_month ='"&be_month&"' and emp_no = '"&emp_no&"'"
	
	        dbconn.execute(sql)
        end if
		
		sql="select max(app_seq) as max_seq from emp_appoint where app_empno='" + emp_no + "'"
		set rs_max=dbconn.execute(sql)
		
		if	isnull(rs_max("max_seq"))  then
			app_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
			app_seq = right(max_seq,3)
			rs_max.Close()
		end if		

        sql = "insert into emp_appoint (app_empno,app_seq,app_id,app_date,app_emp_name,app_to_company,app_to_orgcode,app_to_org,app_to_grade,app_to_job,app_to_position,app_be_company,app_be_orgcode,app_be_org,app_be_grade,app_be_job,app_be_position,app_comment,app_reg_date) values "
		sql = sql +	" ('"&emp_no&"','"&app_seq&"','"&app_id&"','"&app_date&"','"&emp_name&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"','"&app_company&"','"&app_be_orgcode&"','"&app_be_org&"','"&app_grade&"','"&app_job&"','"&app_position&"','"&app_mv_comment&"',now())"	
		
		'response.write sql		
		
		dbconn.execute(sql)	
		
	end if

	if app_id = "�����߷�" then
	
'       �λ縶���� ��������
        sql = "select * from emp_master where emp_no = '"&emp_no&"'"
		set rs_emp = dbconn.execute(sql)
		
		if	rs_emp.eof or rs_emp.bof then
		    response.write"<script language=javascript>"
			response.write"alert('��ϵ� ������ �ƴմϴ�....');"		
			response.write"self.opener.location.reload();"		
	        response.write"window.close();"		
	        response.write"</script>"
	        Response.End
	        dbconn.Close()	
		    else 
     	    sql = "update emp_master set emp_end_date ='"+app_be_enddate+"',emp_pay_id ='2',emp_mod_user = '"+mod_user+"',emp_mod_date = now()  where emp_no = '"&emp_no&"'"
	
	        dbconn.execute(sql)
			
'       memb ���� - �α����� ���ϵ��� �����ؾ� ��(grade�� 6����)
            sql="select * from memb where user_id='"&emp_no&"'"
	        set rs_memb=dbconn.execute(sql)			
		    if not rs_memb.eof then
		       sql = "update memb set grade='6' where user_id='"&emp_no&"'"

		       'response.write sql
		
		        dbconn.execute(sql)	 
		    end if						

'       sawo_memb 
            sql="select * from emp_sawo_mem where sawo_empno='"&emp_no&"'"
	        set rs_sawo=dbconn.execute(sql)			
		    if not rs_sawo.eof then
		       sql = "update emp_sawo_mem set sawo_out='����',sawo_out_date='"+app_be_enddate+"' where sawo_empno='"&emp_no&"'"

		       'response.write sql
		
		        dbconn.execute(sql)	 
		    end if						
			
		end if
		sql="select max(app_seq) as max_seq from emp_appoint where app_empno='" + emp_no + "'"
		set rs_max=dbconn.execute(sql)
		
		if	isnull(rs_max("max_seq"))  then
			app_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
			app_seq = right(max_seq,3)
			rs_max.Close()
		end if		

        sql = "insert into emp_appoint (app_empno,app_seq,app_id,app_date,app_emp_name,app_id_type,app_to_company,app_to_orgcode,app_to_org,app_to_grade,app_to_job,app_to_position,app_be_enddate,app_comment,app_reg_date) values "
		sql = sql +	" ('"&emp_no&"','"&app_seq&"','"&app_id&"','"&app_date&"','"&emp_name&"','"&app_end_type&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"','"&app_be_enddate&"','"&app_end_comment&"',now())"	
		
		'response.write sql		
		
		dbconn.execute(sql)	
		
	end if

	if app_id = "�����߷�" then
	
'       �λ縶���� ��������
        sql = "select * from emp_master where emp_no = '"&emp_no&"'"
		set rs_emp = dbconn.execute(sql)
		
		if  app_be_grade = "�븮1" or app_be_grade = "�븮2" then
		    app_be_job = "�븮"
			else 
			app_be_job = app_be_grade
		end if
		
		if	rs_emp.eof or rs_emp.bof then
		    response.write"<script language=javascript>"
			response.write"alert('��ϵ� ������ �ƴմϴ�....');"		
			response.write"self.opener.location.reload();"		
	        response.write"window.close();"		
	        response.write"</script>"
	        Response.End
	        dbconn.Close()	
		    else 
     	    sql = "update emp_master set emp_grade ='"+app_be_grade+"',emp_grade_date ='"+app_date+"',emp_job ='"+app_be_job+"',emp_mod_user = '"+mod_user+"',emp_mod_date = now() where emp_no = '"&emp_no&"'"
			dbconn.execute(sql)
			
'       memb ���� - ���޺���				
            sql="select * from memb where user_id='"&emp_no&"'"
	        set rs_memb=dbconn.execute(sql)			
		    if not rs_memb.eof then
		       sql = "update memb set user_grade='"&app_be_job&"' where user_id='"&emp_no&"'"

		       'response.write sql
		
		        dbconn.execute(sql)	 
		    end if

		end if
		sql="select max(app_seq) as max_seq from emp_appoint where app_empno='" + emp_no + "'"
		set rs_max=dbconn.execute(sql)
		
		if	isnull(rs_max("max_seq"))  then
			app_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
			app_seq = right(max_seq,3)
			rs_max.Close()
		end if		

        sql = "insert into emp_appoint (app_empno,app_seq,app_id,app_date,app_emp_name,app_id_type,app_to_company,app_to_orgcode,app_to_org,app_to_grade,app_to_job,app_to_position,app_be_company,app_be_orgcode,app_be_org,app_be_grade,app_be_job,app_be_position,app_comment,app_reg_date) values "
		sql = sql +	" ('"&emp_no&"','"&app_seq&"','"&app_id&"','"&app_date&"','"&emp_name&"','"&app_gr_type&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_be_grade&"','"&app_be_job&"','"&app_position&"','"&app_gr_comment&"',now())"	
		
		'response.write sql		
		
		dbconn.execute(sql)	
		
	end if

	if app_id = "��å����" then
	
'       �λ縶���� ��������
        sql = "select * from emp_master where emp_no = '"&emp_no&"'"
		set rs_emp = dbconn.execute(sql)
		
		if	rs_emp.eof or rs_emp.bof then
		    response.write"<script language=javascript>"
			response.write"alert('��ϵ� ������ �ƴմϴ�....');"		
			response.write"self.opener.location.reload();"		
	        response.write"window.close();"		
	        response.write"</script>"
	        Response.End
	        dbconn.Close()	
		    else 
     	    sql = "update emp_master set emp_position ='"+app_be_position+"',emp_company ='"+app_bm_company+"',emp_bonbu ='"+app_bm_bonbu+"',emp_saupbu ='"+app_bm_saupbu+"',emp_team ='"+app_bm_team+"',emp_org_code ='"+app_bm_orgcode+"',emp_org_name ='"+app_bm_org+"',emp_org_baldate ='"+app_date+"',emp_reside_place ='"+app_bm_reside_place+"',emp_reside_company ='"+app_bm_reside_company+"',emp_mod_user = '"+mod_user+"',emp_mod_date = now() where emp_no = '"&emp_no&"'"
	
	        dbconn.execute(sql)
			
'       memb ���� - ��å����				
            sql="select * from memb where user_id='"&emp_no&"'"
	        set rs_memb=dbconn.execute(sql)			
		    if not rs_memb.eof then
			   sql = "update memb set position='"&app_be_position&"',emp_company='"&app_bm_company&"',bonbu='"&app_bm_bonbu&"',saupbu='"&app_bm_saupbu&"',team='"&app_bm_team&"',org_name='"&app_bm_org&"',reside_place='"&app_bm_reside_place&"',reside_company='"&app_bm_reside_company&"',reside='"&reside&"',insa_grade='1',pay_grade='1' where user_id='"&emp_no&"'"

		       'response.write sql
		
		        dbconn.execute(sql)	 
		    end if			
			
		end if
		
'       ���������� ��������	
        if 	app_be_position = "����" or  app_be_position = "�������" or app_be_position = "������" or app_be_position = "��ǥ�̻�"  then
		        sql = "update emp_org_mst set org_empno ='"+emp_no+"',org_emp_name ='"+emp_name+"',org_mod_date =now(),org_mod_user ='"+mod_user+"' where org_code = '"&app_bm_orgcode&"'"
		
		        dbconn.execute(sql)	
		end if 		
		
		sql="select max(app_seq) as max_seq from emp_appoint where app_empno='" + emp_no + "'"
		set rs_max=dbconn.execute(sql)
		
		if	isnull(rs_max("max_seq"))  then
			app_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
			app_seq = right(max_seq,3)
			rs_max.Close()
		end if		

        sql = "insert into emp_appoint (app_empno,app_seq,app_id,app_date,app_emp_name,app_to_company,app_to_orgcode,app_to_org,app_to_grade,app_to_job,app_to_position,app_be_company,app_be_orgcode,app_be_org,app_be_grade,app_be_job,app_be_position,app_comment,app_reg_date) values "
		sql = sql +	" ('"&emp_no&"','"&app_seq&"','"&app_id&"','"&app_date&"','"&emp_name&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"','"&app_bm_company&"','"&app_bm_orgcode&"','"&app_bm_org&"','"&app_grade&"','"&app_job&"','"&app_be_position&"','"&app_bm_comment&"',now())"	

		dbconn.execute(sql)	
		
	end if

	if app_id = "��å����" then
	
'       �λ縶���� ��������
        sql = "select * from emp_master where emp_no = '"&emp_no&"'"
		set rs_emp = dbconn.execute(sql)
		
		if	rs_emp.eof or rs_emp.bof then
		    response.write"<script language=javascript>"
			response.write"alert('��ϵ� ������ �ƴմϴ�....');"		
			response.write"self.opener.location.reload();"		
	        response.write"window.close();"		
	        response.write"</script>"
	        Response.End
	        dbconn.Close()	
		    else 
     	    sql = "update emp_master set emp_position ='����',emp_mod_user = '"+mod_user+"',emp_mod_date = now() where emp_no = '"&emp_no&"'"
	
	        dbconn.execute(sql)
		end if
		
'       memb ���� - ��å����				
            sql="select * from memb where user_id='"&emp_no&"'"
	        set rs_memb=dbconn.execute(sql)			
		    if not rs_memb.eof then
		       sql = "update memb set position='����' where user_id='"&emp_no&"'"

		        dbconn.execute(sql)	 
		    end if				
		
'       ���������� ��������		
		sql = "update emp_org_mst set org_empno ='',org_emp_name ='',org_mod_date =now(),org_mod_user ='"+mod_user+"' where org_code = '"&app_org&"'"
		        dbconn.execute(sql)	 
				
		sql="select max(app_seq) as max_seq from emp_appoint where app_empno='" + emp_no + "'"
		set rs_max=dbconn.execute(sql)
		
		if	isnull(rs_max("max_seq"))  then
			app_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
			app_seq = right(max_seq,3)
			rs_max.Close()
		end if		

        sql = "insert into emp_appoint (app_empno,app_seq,app_id,app_date,app_emp_name,app_id_type,app_to_company,app_to_orgcode,app_to_org,app_to_grade,app_to_job,app_to_position,app_be_company,app_be_orgcode,app_be_org,app_be_grade,app_be_job,app_be_position,app_comment,app_reg_date) values "
		sql = sql +	" ('"&emp_no&"','"&app_seq&"','"&app_id&"','"&app_date&"','"&emp_name&"','"&app_hm_type&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','����','"&app_hm_comment&"',now())"	
		
		'response.write sql		
		
		dbconn.execute(sql)	
		
	end if

	if app_id = "�����߷�" then
	
'       �λ縶���� ��������
        sql = "select * from emp_master where emp_no = '"&emp_no&"'"
		set rs_emp = dbconn.execute(sql)
		
		if	rs_emp.eof or rs_emp.bof then
		    response.write"<script language=javascript>"
			response.write"alert('��ϵ� ������ �ƴմϴ�....');"		
			response.write"self.opener.location.reload();"		
	        response.write"window.close();"		
	        response.write"</script>"
	        Response.End
	        dbconn.Close()	
		    else 
     	    sql = "update emp_master set emp_pay_id ='1',emp_mod_user = '"+mod_user+"',emp_mod_date = now() where emp_no = '"&emp_no&"'"
	
	        dbconn.execute(sql)
		end if
		
		sql="select max(app_seq) as max_seq from emp_appoint where app_empno='" + emp_no + "'"
		set rs_max=dbconn.execute(sql)
		
		if	isnull(rs_max("max_seq"))  then
			app_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
			app_seq = right(max_seq,3)
			rs_max.Close()
		end if		

        sql = "insert into emp_appoint (app_empno,app_seq,app_id,app_date,app_emp_name,app_id_type,app_to_company,app_to_orgcode,app_to_org,app_to_grade,app_to_job,app_to_position,app_start_date,app_finish_date,app_comment,app_bokjik_id,app_be_company,app_be_orgcode,app_be_org,app_be_grade,app_be_job,app_be_position,app_reg_date) values "
		sql = sql +	" ('"&emp_no&"','"&app_seq&"','"&app_id&"','"&app_date&"','"&emp_name&"','"&app_hu_type&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"','"&app_hustart_date&"','"&app_hufinish_date&"','"&app_hu_comment&"','N','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"',now())"	
		
		'response.write sql		
		
		dbconn.execute(sql)	
		
	end if

	if app_id = "¡��߷�" then
	
		sql="select max(app_seq) as max_seq from emp_appoint where app_empno='" + emp_no + "'"
		set rs_max=dbconn.execute(sql)
		
		if	isnull(rs_max("max_seq"))  then
			app_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
			app_seq = right(max_seq,3)
			rs_max.Close()
		end if		

        sql = "insert into emp_appoint (app_empno,app_seq,app_id,app_date,app_emp_name,app_id_type,app_to_company,app_to_orgcode,app_to_org,app_to_grade,app_to_job,app_to_position,app_start_date,app_finish_date,app_comment,app_be_company,app_be_orgcode,app_be_org,app_be_grade,app_be_job,app_be_position,app_reg_date) values "
		sql = sql +	" ('"&emp_no&"','"&app_seq&"','"&app_id&"','"&app_date&"','"&emp_name&"','"&app_di_type&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"','"&app_distart_date&"','"&app_difinish_date&"','"&app_di_comment&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"',now())"	
		
		'response.write sql		
		
		dbconn.execute(sql)	
		
	end if

	if app_id = "����߷�" then
	
		sql="select max(app_seq) as max_seq from emp_appoint where app_empno='" + emp_no + "'"
		set rs_max=dbconn.execute(sql)
		
		if	isnull(rs_max("max_seq"))  then
			app_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
			app_seq = right(max_seq,3)
			rs_max.Close()
		end if		

        sql = "insert into emp_appoint (app_empno,app_seq,app_id,app_date,app_emp_name,app_id_type,app_to_company,app_to_orgcode,app_to_org,app_to_grade,app_to_job,app_to_position,app_reward,app_reg_date) values "
		sql = sql +	" ('"&emp_no&"','"&app_seq&"','"&app_id&"','"&app_date&"','"&emp_name&"','"&app_rw_type&"','"&app_to_company&"','"&app_org&"','"&app_org_name&"','"&app_grade&"','"&app_job&"','"&app_position&"','"&app_rw_comment&"',now())"	
		
		'response.write sql		
		
		dbconn.execute(sql)	
		
	end if

' url = "as_list_ce.asp?page="+page+"&view_sort="+view_sort
  url = "insa_appoint_mg.asp"
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "����� Error�� �߻��Ͽ����ϴ�...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "��ϵǾ����ϴ�...."
	end if
	
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
'	response.write"alert('��� �Ϸ� �Ǿ����ϴ�....');"		
	response.write"location.replace('"&url&"');"
'	response.write"history.go(-2);"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
