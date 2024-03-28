<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	pmg_company = request.form("pmg_company")
	pmg_yymm = request.form("pmg_yymm")
'	pmg_date = request.form("pmg_date")
	objFile = request.form("objFile")

	w_cnt = 0

    emp_user = request.cookies("nkpmg_user")("coo_user_name")

	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")

	Set DbConn = Server.CreateObject("ADODB.Connection")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_org = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
	Set Rs_give = Server.CreateObject("ADODB.Recordset")
    Set Rs_dct = Server.CreateObject("ADODB.Recordset")
	Set rs_com = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect

	dbconn.BeginTrans

	cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
	rs.Open "select * from [1:10000]",cn,"0"

	rowcount=-1
	xgr = rs.getrows
	rowcount = ubound(xgr,2)
	fldcount = rs.fields.count

	tot_cnt = rowcount + 1
    if rowcount > -1 then
		for i=0 to rowcount
			if xgr(1,i) = "" or isnull(xgr(1,i)) then
				exit for
			end if
			pmg_company = xgr(7,i)
	        pmg_yymm = xgr(1,i)
	        pmg_date = xgr(2,i)
	' 사번체크
			Sql = "select * from emp_master where emp_no = '"&xgr(3,i)&"'"
			Set rs_emp = DbConn.Execute(SQL)
			if rs_emp.eof then
                emp_name = ""
			  else
				emp_no             = xgr(3,i)
				emp_name           = rs_emp("emp_name")
				emp_company        = rs_emp("emp_company")
				emp_bonbu          = rs_emp("emp_bonbu")
				emp_saupbu         = rs_emp("emp_saupbu")
				emp_team           = rs_emp("emp_team")
				emp_org_code       = rs_emp("emp_org_code")
				emp_org_name       = rs_emp("emp_org_name")
				emp_reside_place   = rs_emp("emp_reside_place")
				emp_reside_company = rs_emp("emp_reside_company")
				emp_in_date        = rs_emp("emp_in_date")
				emp_grade          = rs_emp("emp_grade")
				emp_position       = rs_emp("emp_position")
				emp_type           = rs_emp("emp_type")
				cost_center        = rs_emp("cost_center")
				cost_group         = rs_emp("cost_group")
				'//2017-06-09 비용구분(cost_group)이 잘못 등록된 경우(ex. 전사공통비, 전사공통비) 콤마 에러 이유가 됨. 앞부분만 등록처리
				IF Trim(cost_center&"")<>"" And InStr(cost_center,",")>0 Then
					cost_center = Left(cost_center, InStr(cost_center,","))
				End If

			end if
			w_cnt = w_cnt + 1

        ' 지급항목
	    pmg_base_pay	  = toString(xgr(12,i),"0")
        pmg_meals_pay	  = toString(xgr(13,i),"0")
        pmg_research_pay  = toString(xgr(14,i),"0")	'연구수당 (신규추가)
        pmg_postage_pay	  = toString(xgr(15,i),"0")
        pmg_re_pay		  = toString(xgr(16,i),"0")
        pmg_overtime_pay  = toString(xgr(17,i),"0")
        pmg_car_pay		  = toString(xgr(18,i),"0")
        pmg_position_pay  = toString(xgr(19,i),"0")
        pmg_custom_pay	  = toString(xgr(20,i),"0")
        pmg_job_pay		  = toString(xgr(21,i),"0")
        pmg_job_support	  = toString(xgr(22,i),"0")
        pmg_jisa_pay	  = toString(xgr(23,i),"0")
        pmg_long_pay	  = toString(xgr(24,i),"0")
        pmg_disabled_pay  = toString(xgr(25,i),"0")
        pmg_family_pay    = 0
        pmg_school_pay    = 0
        pmg_qual_pay      = 0
        pmg_other_pay1    = 0
        pmg_other_pay2    = 0
        pmg_other_pay3    = 0
        pmg_tax_yes       = 0
        pmg_tax_no        = 0
        pmg_tax_reduced   = 0

			pmg_give_total = pmg_base_pay + pmg_meals_pay + pmg_research_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_car_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay
			'pmg_give_total = xgr(25,i)

			meals_pay = pmg_meals_pay
			car_pay   = pmg_car_pay
			meals_tax_pay = 0
	    car_tax_pay = 0
	    if  meals_pay > 100000 then
	    	meals_tax_pay = meals_pay - 100000
				meals_pay =  100000
	    end if
	    if car_pay > 200000 then
	    	car_tax_pay = car_pay - 200000
				car_pay =  200000
	    end if

	    pmg_tax_yes = pmg_base_pay + pmg_postage_pay + pmg_re_pay + pmg_overtime_pay + pmg_position_pay + pmg_custom_pay + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay + meals_tax_pay + car_tax_pay

	    pmg_tax_no = meals_pay + car_pay

	    ' 공제항목
        de_nps_amt			= toString(xgr(27,i),"0")
        de_nhis_amt			= toString(xgr(28,i),"0")
        de_epi_amt			= toString(xgr(29,i),"0")
        de_longcare_amt		= toString(xgr(30,i),"0")
        de_income_tax		= toString(xgr(31,i),"0")
        de_wetax			= toString(xgr(32,i),"0")
        de_year_incom_tax	= toString(xgr(33,i),"0")
        de_year_wetax		= toString(xgr(34,i),"0")
        de_year_incom_tax2	= toString(xgr(35,i),"0")
        de_year_wetax2		= toString(xgr(36,i),"0")
        de_other_amt1		= toString(xgr(37,i),"0")
        de_special_tax		= 0
        de_saving_amt		= 0
        de_sawo_amt			= toString(xgr(38,i),"0")
        de_johab_amt		= 0
        de_school_amt		= toString(xgr(39,i),"0")
        de_nhis_bla_amt		= toString(xgr(40,i),"0")
        de_long_bla_amt		= toString(xgr(41,i),"0")
        de_hyubjo_amt		= toString(xgr(42,i),"0")

	    '공제항목 변경전
	    'de_nps_amt			= toString(xgr(26,i),"0")
        'de_nhis_amt		= toString(xgr(27,i),"0")
        'de_epi_amt			= toString(xgr(28,i),"0")
        'de_longcare_amt	= toString(xgr(29,i),"0")
        'de_income_tax		= toString(xgr(30,i),"0")
        'de_wetax			= toString(xgr(31,i),"0")
	    'de_year_incom_tax	= toString(xgr(32,i),"0")
        'de_year_wetax		= toString(xgr(33,i),"0")
	    'de_year_incom_tax2 = toString(xgr(34,i),"0")
        'de_year_wetax2		= toString(xgr(35,i),"0")
	    'de_other_amt1		= toString(xgr(36,i),"0")
        'de_special_tax	   	= 0
        'de_saving_amt		= 0
        'de_sawo_amt		= toString(xgr(37,i),"0")
        'de_johab_amt		= 0
        'de_school_amt		= toString(xgr(38,i),"0")
        'de_nhis_bla_amt	= toString(xgr(39,i),"0")
        'de_long_bla_amt	= toString(xgr(40,i),"0")
	    'de_hyubjo_amt		= toString(xgr(41,i),"0")

		'de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax + de_year_incom_tax + de_year_wetax + de_year_incom_tax2 + de_year_wetax2 + de_other_amt1 + de_sawo_amt + de_school_amt + de_nhis_bla_amt + de_long_bla_amt + de_hyubjo_amt
        'de_deduct_total = xgr(38,i)

        ' 2019.03.15 윤성희,박정신 계산에 의한 공제액 합계가 아니라 엑셀 컬럼의 내용을 그대로 계산없이 설정한다.
        de_deduct_total = xgr(43,i)

        Sql = "SELECT * FROM pay_bank_account where emp_no = '"&emp_no&"'"
        Set rs_bnk = DbConn.Execute(SQL)
        if not rs_bnk.eof then
            bank_name = rs_bnk("bank_name")
            account_no = rs_bnk("account_no")
            account_holder = rs_bnk("account_holder")
	    else
      	    bank_name = ""
	    	account_no = ""
		    account_holder = ""
        end if
      rs_bnk.close()

        sql = "select * from pay_month_give where pmg_yymm = '"&pmg_yymm&"' and pmg_id = '1' and pmg_emp_no = '"&emp_no&"' and pmg_company = '"&pmg_company&"'"
'Response.write sql&"<br>"
        set Rs_give=dbconn.execute(sql)
        if not (Rs_give.eof or Rs_give.bof) then
            sql = "delete from pay_month_give where pmg_yymm = '"&pmg_yymm&"' and pmg_id = '1' and pmg_emp_no = '"&emp_no&"' and pmg_company = '"&pmg_company&"'"
'Response.write sql&"<br>"
            dbconn.execute(sql)
        end if

        sql="INSERT INTO pay_month_give "&_
            "(  pmg_yymm, pmg_id, pmg_emp_no, pmg_company, pmg_date "&_
            " , pmg_in_date, pmg_emp_name, pmg_emp_type, pmg_org_code, pmg_org_name "&_
            " , pmg_bonbu, pmg_saupbu, pmg_team, pmg_reside_place, pmg_reside_company "&_
            " , pmg_grade, pmg_position, pmg_base_pay, pmg_meals_pay, pmg_postage_pay "&_
            " , pmg_re_pay, pmg_overtime_pay, pmg_car_pay, pmg_position_pay, pmg_custom_pay  "&_
            " , pmg_job_pay, pmg_job_support, pmg_jisa_pay, pmg_long_pay, pmg_disabled_pay "&_
            " , pmg_family_pay, pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2 "&_
            " , pmg_other_pay3, pmg_tax_yes, pmg_tax_no, pmg_tax_reduced, pmg_give_total "&_
            " , pmg_bank_name, pmg_account_no, pmg_account_holder, cost_group, cost_center "&_
            " , pmg_reg_date, pmg_reg_user, pmg_research_pay "&_
            ")  "&_
            "VALUES "&_
            "(  '"&pmg_yymm&"', '1', '"&emp_no&"', '"&pmg_company&"', '"&pmg_date&"' "&_
            " , '"&emp_in_date&"', '"&emp_name&"', '"&emp_type&"', '"&emp_org_code&"', '"&emp_org_name&"' "&_
            " , '"&emp_bonbu&"', '"&emp_saupbu&"', '"&emp_team&"', '"&emp_reside_place&"', '"&emp_reside_company&"' "&_
            " , '"&emp_grade&"','"&emp_position&"','"&pmg_base_pay&"','"&pmg_meals_pay&"','"&pmg_postage_pay&"' "&_
            " , '"&pmg_re_pay&"', '"&pmg_overtime_pay&"', '"&pmg_car_pay&"', '"&pmg_position_pay&"', '"&pmg_custom_pay&"' "&_
            " , '"&pmg_job_pay&"', '"&pmg_job_support&"', '"&pmg_jisa_pay&"', '"&pmg_long_pay&"', '"&pmg_disabled_pay&"' "&_
            " , '"&pmg_family_pay&"', '"&pmg_school_pay&"', '"&pmg_qual_pay&"', '"&pmg_other_pay1&"', '"&pmg_other_pay2&"' "&_
            " , '"&pmg_other_pay3&"', '"&pmg_tax_yes&"', '"&pmg_tax_no&"', '"&pmg_tax_reduced&"', '"&pmg_give_total&"' "&_
            " , '"&bank_name&"', '"&account_no&"', '"&account_holder&"', '"&cost_group&"', '"&cost_center&"' "&_
            " , now(),'"&emp_user&"', '"&pmg_research_pay&"')"
'Response.write sql&"<br>"
		dbconn.execute(sql)


		sql = "select * from pay_month_deduct where de_yymm = '"&pmg_yymm&"' and de_id = '1' and de_emp_no = '"&emp_no&"' and de_company = '"&pmg_company&"'"
'Response.write sql&"<br>"
		set Rs_dct=dbconn.execute(sql)
        if not (Rs_dct.eof or Rs_dct.bof) then
            sql = "delete from pay_month_deduct where de_yymm = '"&pmg_yymm&"' and de_id = '1' and de_emp_no = '"&emp_no&"' and de_company = '"&pmg_company&"'"
'Response.write sql&"<br>"
            dbconn.execute(sql)
        end if

        sql = "INSERT INTO pay_month_deduct "&_
                "(  de_yymm, de_id, de_emp_no, de_company, de_date                                  "&_
                " , de_emp_name, de_emp_type, de_org_code, de_org_name, de_bonbu                    "&_
                " , de_saupbu, de_team, de_reside_place, de_reside_company, de_grade                "&_
                " , de_position, de_nps_amt, de_nhis_amt, de_epi_amt, de_longcare_amt               "&_
                " , de_income_tax, de_wetax, de_year_incom_tax, de_year_wetax, de_year_incom_tax2   "&_
                " , de_year_wetax2, de_other_amt1, de_saving_amt, de_sawo_amt, de_johab_amt         "&_
                " , de_hyubjo_amt, de_school_amt, de_nhis_bla_amt, de_long_bla_amt, de_deduct_total "&_
                " , cost_group, cost_center, de_reg_date, de_reg_user                               "&_
                " )                                                                                 "&_
                " values                                                                            "&_
                "( '"&pmg_yymm&"', '1', '"&emp_no&"', '"&pmg_company&"', '"&pmg_date&"'             "&_
                " , '"&emp_name&"', '"&emp_type&"', '"&emp_org_code&"', '"&emp_org_name&"', '"&emp_bonbu&"' "&_
                " , '"&emp_saupbu&"', '"&emp_team&"', '"&emp_reside_place&"', '"&emp_reside_company&"', '"&emp_grade&"' "&_
                " , '"&emp_position&"', '"&de_nps_amt&"', '"&de_nhis_amt&"', '"&de_epi_amt&"', '"&de_longcare_amt&"' "&_
                " , '"&de_income_tax&"', '"&de_wetax&"', '"&de_year_incom_tax&"', '"&de_year_wetax&"','"&de_year_incom_tax2&"' "&_
                " , '"&de_year_wetax2&"', '"&de_other_amt1&"', '"&de_saving_amt&"', '"&de_sawo_amt&"', '"&de_johab_amt&"' "&_
                " , '"&de_hyubjo_amt&"', '"&de_school_amt&"', '"&de_nhis_bla_amt&"', '"&de_long_bla_amt&"', '"&de_deduct_total&"' "&_
                " , '"&cost_group&"', '"&cost_center&"', now(), '"&emp_user&"')"
'Response.write sql&"<br>"
        dbconn.execute(sql)
        if Err.number <> 0 then
            Response.write "(ErrDesc=" & err.Description & "&ErrCode=" & err.number & ")" & " [sql : " & sql4 & "]<br>"
        end if

		next
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = "변경중 Error가 발생하였습니다...."
	else
		dbconn.CommitTrans
		end_msg = cstr(w_cnt) +" 건 등록 완료되었습니다...."
	end if

	'err_msg = cstr(rowcount+1) + " 건 처리되었습니다..."
	Response.write"<script language=javascript>"
	Response.write"alert('"&end_msg&"');"
	Response.write"location.replace('insa_pay_month_pay_mg.asp');"
	Response.write"</script>"
	Response.End

	rs.close
	cn.close
	'rs_etc.close
	set rs = nothing
	set cn = nothing
	'set rs_etc = nothing
%>