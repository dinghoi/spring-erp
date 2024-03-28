<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	Server.ScriptTimeOut = 1200

	end_month=Request("end_month")
	end_yn=Request("end_yn")
		
	cost_year = mid(end_month,1,4)
	cost_month = mid(end_month,5)
	
	from_date = mid(end_month,1,4) + "-" + mid(end_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))
	start_date = dateadd("m",-1,from_date)

	org_company = "케이원정보통신"

	reside_sw = "Y"

	sql = "select count(*) from tax_bill where (bill_id = '1') and (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and cost_reg_yn = 'N'"
	Set rscount = Dbconn.Execute (sql)	
	total_record = cint(rscount(0)) 'Result.RecordCount
	
	if total_record > 0 then
		reside_sw = "N"
	end if

	sql = "select count(*) from cost_end where end_month = '"&end_month&"' and saupbu <> '상주비용'"
	Set rscount = Dbconn.Execute (sql)	
	total_record = cint(rscount(0)) 'Result.RecordCount
	
	if total_record > 0 then
		sql = "select count(*) from cost_end where end_month = '"&end_month&"' and (end_yn = 'N' or end_yn = 'C') and (saupbu <> '상주비용' and saupbu <> '공통비/직접비배분')"
		Set rscount = Dbconn.Execute (sql)	
		total_record = cint(rscount(0)) 'Result.RecordCount
		if total_record > 0 then
			reside_sw = "N"
		end if
	end if
	
	if reside_sw = "N" then
		response.write"<script language=javascript>"
		response.write"alert('전체 비용 마감이 되어 있지 않습니다 !!');"
		response.write"location.replace('cost_end_mg.asp');"
		response.write"</script>"
		Response.End
    else		
		response.write"<script language=javascript>"
		response.write"alert('마감처리중!!!');"
		response.write"</script>"
		
		dbconn.BeginTrans

		response.write(now())
	
' 인사마스터 및 급여DATA에 관리사업부 지정

'		sql = "select emp_saupbu from emp_master_month where (emp_month ='"&end_month&"') group by emp_saupbu"		
		sql = "select emp_saupbu from emp_master_month where (emp_month ='"&end_month&"') and (cost_center <> '손익제외') group by emp_saupbu"		
		rs.Open sql, Dbconn, 1
	
		do until rs.eof		  
			saupbu = rs("emp_saupbu")
			sql = "select * from sales_org where saupbu = '"&saupbu&"'"
			set rs_etc=dbconn.execute(sql)
			if rs_etc.eof or rs_etc.bof then							
				saupbu = ""
			end if
			sql = "update emp_master_month set mg_saupbu = '"&saupbu&"' where emp_month ='"&end_month&"' and emp_saupbu = '"&rs("emp_saupbu")&"'"
			'dbconn.execute(sql)
Response.write "<br>1::"&sql

			sql = "update pay_month_give set mg_saupbu = '"&saupbu&"' where pmg_yymm ='"&end_month&"' and pmg_saupbu = '"&rs("emp_saupbu")&"'"
			'dbconn.execute(sql)
Response.write "<br>2::"&sql
			rs.movenext()
		loop
		rs.close()

'		sql = "select emp_reside_company from emp_master_month where (emp_month ='"&end_month&"') and (mg_saupbu = '') and (emp_reside_company <> '') group by emp_reside_company"		
		sql = "select emp_reside_company from emp_master_month where (emp_month ='"&end_month&"') and (mg_saupbu = '') and (emp_reside_company <> '') and (cost_center <> '손익제외') group by emp_reside_company"		
		rs.Open sql, Dbconn, 1
	
		do until rs.eof		  
			sql = "select * from trade where trade_name = '"&rs("emp_reside_company")&"'"
			set rs_trade=dbconn.execute(sql)
			if rs_trade.eof or rs_trade.bof then
				saupbu = "Error"		
			  else
				saupbu = rs_trade("saupbu")
			end if		  
			sql = "update emp_master_month set mg_saupbu = '"&saupbu&"' where emp_month ='"&end_month&"' and mg_saupbu = '' and emp_reside_company = '"&rs("emp_reside_company")&"'"
			'dbconn.execute(sql)
Response.write "<br>3::"&sql

			sql = "update pay_month_give set mg_saupbu = '"&saupbu&"' where pmg_yymm ='"&end_month&"' and mg_saupbu = '' and pmg_reside_company = '"&rs("emp_reside_company")&"'"
			'dbconn.execute(sql)
Response.write "<br>4::"&sql

			rs.movenext()
		loop
		rs.close()

' 알바비용 관리사업부 및 비용유형 지정
' 초기값 Clear
'		sql = "update pay_alba_cost set mg_saupbu = '', cost_center = '' where rever_yymm ='"&end_month&"'"
'		dbconn.execute(sql)

		sql = "update pay_alba_cost set cost_center = '상주직접비' where (cost_company <> '공통' and cost_company <> '전사' and cost_company <> '부문' and cost_company <> '기타' and cost_company <> '본사' and cost_company <> '케이원정보통신' and cost_company <> '') and (rever_yymm ='"&end_month&"')"
		'dbconn.execute(sql)
Response.write "<br>5::"&sql

		sql = "select company,org_name from pay_alba_cost where (cost_company = '공통' Or cost_company <> '전사' or cost_company <> '부문' or cost_company = '기타' or cost_company = '본사' or cost_company = '케이원정보통신' or cost_company = '') and (rever_yymm ='"&end_month&"') group by company,org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof
	
			sql = "select * from emp_org_mst_month where org_month = '"&end_month&"' and org_company = '"&rs("company")&"' and org_name = '"&rs("org_name")&"'"
			set rs_org=dbconn.execute(sql)
			if rs_org.eof or rs_org.bof then
				cost_center = "전사공통비"
				cost_company = ""
				group_name = ""	
				bill_trade_name = ""
			  else
				cost_center = rs_org("org_cost_center")
				cost_company = ""
				group_name = ""	
				bill_trade_name = ""
			end if		  

			sql = "update pay_alba_cost set cost_center = '"&cost_center&"' where (cost_company = '공통' or cost_company = '기타' or cost_company = '본사' or cost_company = '케이원정보통신' or cost_company = '') and (rever_yymm ='"&end_month&"') and org_name = '"&rs("org_name")&"'"
			'dbconn.execute(sql)
Response.write "<br>6::"&sql

			rs.movenext()
		loop
		rs.close()

' 알바비용 관리사업부 지정
		sql = "select saupbu,cost_company from pay_alba_cost where (cost_center = '상주직접비') and (rever_yymm ='"&end_month&"') group by saupbu,cost_company"
		rs.Open sql, Dbconn, 1
		do until rs.eof		  
			saupbu = rs("saupbu")
			sql = "select * from sales_org where saupbu = '"&saupbu&"'"
			set rs_etc=dbconn.execute(sql)
			if rs_etc.eof or rs_etc.bof then							
				if rs("cost_company") = "" or isnull(rs("cost_company")) then
					saupbu = ""
				  else
					sql = "select * from trade where trade_name = '"&rs("cost_company")&"'"
					set rs_trade=dbconn.execute(sql)
					if rs_trade.eof or rs_trade.bof then
						saupbu = "Error"		
					  else
						saupbu = rs_trade("saupbu")
					end if		  
				end if
			end if

			sql = "update pay_alba_cost set mg_saupbu = '"&saupbu&"' where (cost_center = '상주직접비') and (rever_yymm ='"&end_month&"') and (saupbu = '"&rs("saupbu")&"') and (cost_company = '"&rs("cost_company")&"')"
			'dbconn.execute(sql)
Response.write "<br>7::"&sql

			rs.movenext()
		loop
		rs.close()

' 알바비용 직접비 관리사업부 지정
		sql = "select saupbu from pay_alba_cost where (cost_center = '직접비') and (rever_yymm ='"&end_month&"') group by saupbu"
		rs.Open sql, Dbconn, 1
		do until rs.eof		  
			sql = "update pay_alba_cost set mg_saupbu = '"&rs("saupbu")&"' where (cost_center = '직접비') and (rever_yymm ='"&end_month&"') and (saupbu = '"&rs("saupbu")&"')"
			'dbconn.execute(sql)
Response.write "<br>8::"&sql

			rs.movenext()
		loop
		rs.close()

' 일반비용 관리사업부 및 비용유형 지정
' 초기값 Clear
		sql = "update general_cost set mg_saupbu = '', cost_center = '' where (tax_bill_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "
		'dbconn.execute(sql)
Response.write "<br>9::"&sql

' 세금계산서는 입력시 관리사업부 지정하게 변경
		sql = "update general_cost set cost_center = '' where (tax_bill_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
		'dbconn.execute(sql)
Response.write "<br>a::"&sql

' 비용유형 셋팅
		sql = "update general_cost set cost_center = '상주직접비' where (pl_yn = 'Y') and (company <> '공통' and company <> '기타' and company <> '본사' and company <> '케이원정보통신' and company <> '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
		'dbconn.execute(sql)
Response.write "<br>b::"&sql

' 공통비 비용 유형세팅 ( 비용 )
		sql = "select emp_company,org_name from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof
	
			sql = "select * from emp_org_mst_month where org_month = '"&end_month&"' and org_company = '"&rs("emp_company")&"' and org_name = '"&rs("org_name")&"'"
			set rs_org=dbconn.execute(sql)
			if rs_org.eof or rs_org.bof then
				cost_center = "전사공통비"
				cost_company = ""
				group_name = ""	
				bill_trade_name = ""
			  else
				cost_center = rs_org("org_cost_center")
				cost_company = ""
				group_name = ""	
				bill_trade_name = ""
			end if		  

			sql = "update general_cost set cost_center = '"&cost_center&"' where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"
			'dbconn.execute(sql)
Response.write "<br>c::"&sql

			rs.movenext()
		loop
		rs.close()

' 공통비 비용 유형세팅 ( 세금계산서 )
' 관리사업부 있는경우
		sql = "select emp_company,mg_saupbu from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu <> '') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,mg_saupbu"
		rs.Open sql, Dbconn, 1
		do until rs.eof
	
			sql = "select * from emp_org_mst_month where org_month = '"&end_month&"' and org_company = '"&rs("emp_company")&"' and org_name = '"&rs("mg_saupbu")&"'"
			set rs_org=dbconn.execute(sql)
			if rs_org.eof or rs_org.bof then
				cost_center = "전사공통비"
				cost_company = ""
				group_name = ""	
				bill_trade_name = ""
			  else
				cost_center = rs_org("org_cost_center")
				cost_company = ""
				group_name = ""	
				bill_trade_name = ""
			end if		  

			sql = "update general_cost set cost_center = '"&cost_center&"' where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu <> '') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (mg_saupbu = '"&rs("mg_saupbu")&"')"
			'dbconn.execute(sql)
Response.write "<br>d::"&sql

			rs.movenext()
		loop
		rs.close()
' 관리사업부가 없는경우
		sql = "select emp_company,org_name from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu = '') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof
	
			sql = "select * from emp_org_mst_month where org_month = '"&end_month&"' and org_company = '"&rs("emp_company")&"' and org_name = '"&rs("org_name")&"'"
			set rs_org=dbconn.execute(sql)
			if rs_org.eof or rs_org.bof then
				cost_center = "전사공통비"
				cost_company = ""
				group_name = ""	
				bill_trade_name = ""
			  else
				cost_center = rs_org("org_cost_center")
				cost_company = ""
				group_name = ""	
				bill_trade_name = ""
			end if		  

			sql = "update general_cost set cost_center = '"&cost_center&"' where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu = '') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"
			'dbconn.execute(sql)
Response.write "<br>e::"&sql

			rs.movenext()
		loop
		rs.close()

' 일반비용 관리사업부 지정
		sql = "select saupbu,company from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (cost_center = '상주직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by saupbu,company"
		rs.Open sql, Dbconn, 1
		do until rs.eof		  
			saupbu = rs("saupbu")
			sql = "select * from sales_org where saupbu = '"&saupbu&"'"
			set rs_etc=dbconn.execute(sql)
			if rs_etc.eof or rs_etc.bof then							
				if rs("company") = "" or isnull(rs("company")) then
					saupbu = ""
				  else
					sql = "select * from trade where trade_name = '"&rs("company")&"'"
					set rs_trade=dbconn.execute(sql)
					if rs_trade.eof or rs_trade.bof then
						saupbu = "Error"		
					  else
						saupbu = rs_trade("saupbu")
					end if		  
				end if
			end if

			sql = "update general_cost set mg_saupbu = '"&saupbu&"' where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (saupbu = '"&rs("saupbu")&"') and (company = '"&rs("company")&"')"
			'dbconn.execute(sql)
Response.write "<br>f::"&sql

			rs.movenext()
		loop
		rs.close()

' 세금계산서 비용 관리사업부 지정
'		sql = "select company from general_cost where (tax_bill_yn = 'Y') and (cost_center = '상주직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by company"
'		rs.Open sql, Dbconn, 1
'		do until rs.eof		  
'			sql = "select * from trade where trade_name = '"&rs("company")&"'"
'			set rs_trade=dbconn.execute(sql)
'			if rs_trade.eof or rs_trade.bof then
'				saupbu = "Error"		
'			  else
'				saupbu = rs_trade("saupbu")
'			end if		  

'			sql = "update general_cost set mg_saupbu = '"&saupbu&"' where (tax_bill_yn = 'Y') and (cost_center = '상주직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (company = '"&rs("company")&"')"
'			dbconn.execute(sql)

'			rs.movenext()
'		loop
'		rs.close()

'		sql = "update general_cost set cost_center = '전사공통비' where (tax_bill_yn = 'Y') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
'		dbconn.execute(sql)

' 비용 직접비 관리사업부 지정
		sql = "select saupbu from general_cost where (pl_yn = 'Y') and (cost_center = '직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by saupbu"
		rs.Open sql, Dbconn, 1
		do until rs.eof		  
			sql = "update general_cost set mg_saupbu = '"&rs("saupbu")&"' where (pl_yn = 'Y') and (cost_center = '직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (saupbu = '"&rs("saupbu")&"')"
			'dbconn.execute(sql)
Response.write "<br>g::"&sql

			rs.movenext()
		loop
		rs.close()

' 사간거래 체크
		sql = "select * from general_cost where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and tax_bill_yn = 'Y'"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			sql = "select trade_id from trade where trade_no = '"&rs("customer_no")&"'"
			set rs_trade=dbconn.execute(sql)
			if rs_trade.eof or rs_trade.bof then
				cost_center = ""
			  else
			  	if rs_trade("trade_id") = "계열사" then
					sql = "update general_cost set cost_center = '회사간거래' where slip_date ='"&rs("slip_date")&"' and slip_seq = '"&rs("slip_seq")&"'"
					'dbconn.execute(sql)
Response.write "<br>h::"&sql

				end if
			end if
			
			rs.movenext()
		loop
		rs.close()

' 일반비용 관리사업부와 비용유형 지정 끝

' 카드사용 관리사업부 및 비용유형 지정
' 초기값 Clear
'		sql = "update card_slip set mg_saupbu = '', cost_center = '' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
'		dbconn.execute(sql)
' 비용유형 지정
		sql = "update card_slip set cost_center = '상주직접비' where (pl_yn = 'Y') and (reside_company <> '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
		'dbconn.execute(sql)
Response.write "<br>j::"&sql

		sql = "select org_name from card_slip where (pl_yn = 'Y') and (reside_company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			sql = "select org_cost_center from emp_org_mst_month where org_month = '"&end_month&"' and org_name = '"&rs("org_name")&"' group by org_name"
			set rs_org=dbconn.execute(sql)
			sql = "update card_slip set cost_center = '"&rs_org("org_cost_center")&"' where (pl_yn = 'Y') and (reside_company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and org_name = '"&rs("org_name")&"'"
			'dbconn.execute(sql)
Response.write "<br>k::"&sql

			rs.movenext()
		loop
		rs.close()
' 카드사용 직접비 관리사업부 지정
		sql = "select saupbu from card_slip where (pl_yn = 'Y') and (cost_center = '직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by saupbu"
		rs.Open sql, Dbconn, 1
		do until rs.eof		  
			sql = "update card_slip set mg_saupbu = '"&rs("saupbu")&"' where (pl_yn = 'Y') and (cost_center = '직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (saupbu = '"&rs("saupbu")&"')"
			'dbconn.execute(sql)
Response.write "<br>l::"&sql

			rs.movenext()
		loop
		rs.close()

' 카드사용 상주직접비 관리사업부 지정
		sql = "select reside_company from card_slip where (pl_yn = 'Y') and (cost_center = '상주직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by reside_company"
		rs.Open sql, Dbconn, 1
		do until rs.eof		  
			sql = "select * from trade where trade_name = '"&rs("reside_company")&"'"
			set rs_trade=dbconn.execute(sql)
			if rs_trade.eof or rs_trade.bof then
				saupbu = "Error"		
			  else
				saupbu = rs_trade("saupbu")
			end if		  

			sql = "update card_slip set mg_saupbu = '"&saupbu&"' where (pl_yn = 'Y') and (cost_center = '상주직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (reside_company = '"&rs("reside_company")&"')"
			'dbconn.execute(sql)
Response.write "<br>m::"&sql

			rs.movenext()
		loop
		rs.close()

' 카드사용 관리사업부 및 비용유형 지정 끝

' 차량관리비 관리사업부 및 비용유형 지정
' 초기값 Clear
'		sql = "update transit_cost set mg_saupbu = '', cost_center = '' where (run_date >='"&from_date&"' and run_date <='"&to_date&"')"
'		dbconn.execute(sql)
' 차량관리비 비용유형 지정	
		sql = "update transit_cost set cost_center = '상주직접비' where (company <> '공통' and company <> '기타' and company <> '케이원정보통신' and company <> '') and (run_date >='"&from_date&"' and run_date <='"&to_date&"')"
		'dbconn.execute(sql)
Response.write "<br>n::"&sql

		sql = "select org_name from transit_cost where (company = '공통' or company = '기타' or company = '케이원정보통신' or company = '') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			sql = "select org_cost_center from emp_org_mst_month where org_month = '"&end_month&"' and org_name = '"&rs("org_name")&"' group by org_name"
			set rs_org=dbconn.execute(sql)
			sql = "update transit_cost set cost_center = '"&rs_org("org_cost_center")&"' where (company = '공통' or company = '기타' or company = '케이원정보통신' or company = '') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and org_name = '"&rs("org_name")&"'"
			'dbconn.execute(sql)
Response.write "<br>o::"&sql

			rs.movenext()
		loop
		rs.close()
' 상주이면서 기타로 입력시 기타를 상주 회사로 변경
		sql = "select org_name from transit_cost where (company = '공통' or company = '기타' or company = '케이원정보통신' or company = '') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (cost_center = '상주직접비') group by org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			sql = "select org_reside_company from emp_org_mst_month where org_month = '"&end_month&"' and org_name = '"&rs("org_name")&"' group by org_name"
			set rs_org=dbconn.execute(sql)
			sql = "update transit_cost set company = '"&rs_org("org_reside_company")&"' where (company = '공통' or company = '기타' or company = '케이원정보통신' or company = '') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (cost_center = '상주직접비') and org_name = '"&rs("org_name")&"'"
			'dbconn.execute(sql)
Response.write "<br>p::"&sql

			rs.movenext()
		loop
		rs.close()

' 차량관리비 관리사업부 지정
		sql = "select saupbu,company from transit_cost where (cost_center = '상주직접비') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by saupbu,company"
		rs.Open sql, Dbconn, 1
		do until rs.eof		  
			sql = "select * from trade where trade_name = '"&rs("company")&"'"
			set rs_trade=dbconn.execute(sql)
			if rs_trade.eof or rs_trade.bof then
				saupbu = "Error"		
			  else
				saupbu = rs_trade("saupbu")
			end if		  

			sql = "update transit_cost set mg_saupbu = '"&saupbu&"' where (cost_center = '상주직접비') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (saupbu = '"&rs("saupbu")&"') and (company = '"&rs("company")&"')"
			'dbconn.execute(sql)
Response.write "<br>q::"&sql

			rs.movenext()
		loop
		rs.close()

' 차량관리비 직접비 관리사업부 지정
		sql = "select saupbu from transit_cost where (cost_center = '직접비') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by saupbu"
		rs.Open sql, Dbconn, 1
		do until rs.eof		  
			sql = "update transit_cost set mg_saupbu = '"&rs("saupbu")&"' where (cost_center = '직접비') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (saupbu = '"&rs("saupbu")&"')"
			'dbconn.execute(sql)
Response.write "<br>r::"&sql

			rs.movenext()
		loop
		rs.close()
		response.write(now())

' 비용구분 Marking 종료

' 회사별 비용 마감전 기존 데이터 Clear	
		sql = "update company_cost set cost_amt_"&cost_month&"='0' where cost_year ='"&cost_year&"'"
		'dbconn.execute(sql)
Response.write "<br>s::"&sql


' 4대보험율과 기타 인건비율 검색
		sql = "select * from insure_per where insure_year = '"&cost_year&"'"
		set rs_etc=dbconn.execute(sql)
		insure_tot_per = rs_etc("insure_tot_per")
		income_tax_per = rs_etc("income_tax_per")
		annual_pay_per = rs_etc("annual_pay_per")
		retire_pay_per = rs_etc("retire_pay_per")
		rs_etc.close()

' 급여 SUM
' 1. 상주자 인건비	
		sql = "select mg_saupbu,cost_center,pmg_reside_company,pmg_id,sum(pmg_give_total) as tot_cost,sum(pmg_base_pay) as base_pay,sum(pmg_meals_pay) as meals_pay,sum(pmg_overtime_pay) as overtime_pay,sum(pmg_tax_no) as tax_no from pay_month_give where (pmg_yymm ='"&end_month&"') and (cost_center <> '손익제외') group by mg_saupbu,cost_center,pmg_reside_company,pmg_id"
		
		rs.Open sql, Dbconn, 1
	
		do until rs.eof
		  
      	  	if rs("pmg_id") = "1" or rs("pmg_id") = "2" then			
				if rs("pmg_id") = "1" then
					sort_seq = 0
					cost_detail = "급여"
				  elseif rs("pmg_id") = "2" then
					sort_seq = 2
					cost_detail = "상여"
				  elseif rs("pmg_id") = "4" then
					sort_seq = 3
					cost_detail = "연차수당"
				  else
					sort_seq = 9
					cost_detail = "기타"
				end if		  		
	
				group_name = ""
				bill_trade_name = ""				
				if rs("cost_center") = "상주직접비" then
					sql = "select * from trade where trade_name = '"&rs("pmg_reside_company")&"'"
					set rs_trade=dbconn.execute(sql)
					if rs_trade.eof or rs_trade.bof then
						group_name = "Error"
						bill_trade_name = "Error"		
					  else
						group_name = rs_trade("group_name")
						bill_trade_name = rs_trade("bill_trade_name")
					end if		  
				end if
					
				sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='"&cost_detail&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
				set rs_cost=dbconn.execute(sql)
			
				if rs_cost.eof or rs_cost.bof then
					sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','인건비','"&cost_detail&"','"&rs("mg_saupbu")&"',"&rs("tot_cost")&","&sort_seq&")"
					'dbconn.execute(sql)
				  else
					sql = "update company_cost set cost_amt_"&cost_month&"="&rs("tot_cost")&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='"&cost_detail&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
					'dbconn.execute(sql)
				end if		
				Response.write "<br/>" & sql
      	  		if rs("pmg_id") = "1" then			
' 4대보험 요율 적용
                    'insure_tot = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * insure_tot_per / 100)	
                    insure_tot = clng((clng(rs("tot_cost"))) * insure_tot_per / 100)	
					sort_seq = 2

					sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='4대보험' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
					set rs_cost=dbconn.execute(sql)
				
					if rs_cost.eof or rs_cost.bof then
						sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','인건비','4대보험','"&rs("mg_saupbu")&"',"&insure_tot&","&sort_seq&")"
						'dbconn.execute(sql)
					  else
						sql = "update company_cost set cost_amt_"&cost_month&"="&insure_tot&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='4대보험' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
						'dbconn.execute(sql)
					end if		

				Response.write "<br/>" & sql
' 소득세 종업원분
                    'income_tax = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * income_tax_per / 100)		
                    income_tax = clng((clng(rs("tot_cost"))) * income_tax_per / 100)		
					sort_seq = 3

					sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='소득세종업원분' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
					set rs_cost=dbconn.execute(sql)
				
					if rs_cost.eof or rs_cost.bof then
						sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','인건비','소득세종업원분','"&rs("mg_saupbu")&"',"&income_tax&","&sort_seq&")"
						'dbconn.execute(sql)
					  else
						sql = "update company_cost set cost_amt_"&cost_month&"="&income_tax&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='소득세종업원분' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
						'dbconn.execute(sql)
					end if		
					
				Response.write "<br/>" & sql
' 연차수당
					annual_pay = clng((clng(rs("base_pay"))+clng(rs("meals_pay"))+clng(rs("overtime_pay"))) * annual_pay_per / 100)		
					sort_seq = 4

					sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='연차수당' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
					set rs_cost=dbconn.execute(sql)
				
					if rs_cost.eof or rs_cost.bof then
						sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','인건비','연차수당','"&rs("mg_saupbu")&"',"&annual_pay&","&sort_seq&")"
						'dbconn.execute(sql)
					  else
						sql = "update company_cost set cost_amt_"&cost_month&"="&annual_pay&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='연차수당' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
						'dbconn.execute(sql)
					end if		
					
				Response.write "<br/>" & sql
' 퇴직충당금
					retire_pay = clng((clng(rs("base_pay"))+clng(rs("meals_pay"))+clng(rs("overtime_pay"))) * retire_pay_per / 100)		
					sort_seq = 5

					sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='퇴직충당금' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
					set rs_cost=dbconn.execute(sql)
				
					if rs_cost.eof or rs_cost.bof then
						sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','인건비','퇴직충당금','"&rs("mg_saupbu")&"',"&retire_pay&","&sort_seq&")"
						'dbconn.execute(sql)
					  else
						sql = "update company_cost set cost_amt_"&cost_month&"="&retire_pay&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='인건비' and cost_detail ='퇴직충당금' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
						'dbconn.execute(sql)
					end if		
					
				Response.write "<br/>" & sql
				end if
			end if		

			rs.movenext()
		loop
		rs.close()
	
' 알바비
		sql = "select cost_center,mg_saupbu,cost_company,sum(alba_give_total) as cost from pay_alba_cost where (rever_yymm ='"&end_month&"') group by cost_center,mg_saupbu,cost_company"
		rs.Open sql, Dbconn, 1
		do until rs.eof	
	
			group_name = ""
			bill_trade_name = ""		
			if rs("cost_center") = "상주직접비" then
				sql = "select * from trade where trade_name = '"&rs("cost_company")&"'"
				set rs_trade=dbconn.execute(sql)
				if rs_trade.eof or rs_trade.bof then
					group_name = "Error"
					bill_trade_name = "Error"		
				  else
					group_name = rs_trade("group_name")
					bill_trade_name = rs_trade("bill_trade_name")
				end if		  		  
			end if
	
			sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and cost_center ='"&rs("cost_center")&"' and company ='"&rs("cost_company")&"' and cost_id ='인건비' and cost_detail ='알바비' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"'"
			set rs_cost=dbconn.execute(sql)
		
			sort_seq = 8
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("cost_company")&"','"&bill_trade_name&"','"&group_name&"','인건비','알바비','"&rs("mg_saupbu")&"',"&rs("cost")&","&sort_seq&")"
				'dbconn.execute(sql)
			  else
				sum_cost = int(rs_cost("cost")) + clng(rs("cost"))
				sql = "update company_cost set cost_amt_"&cost_month&"="&sum_cost&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and cost_center ='"&rs("cost_center")&"' and company ='"&rs("cost_company")&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and cost_id ='인건비' and cost_detail ='알바비' and saupbu ='"&rs("mg_saupbu")&"'"
				'dbconn.execute(sql)
			end if		
			
				Response.write "<br/>" & sql
			rs.movenext()
		loop
		rs.close()
' 알바비 종료

' 비용 SUM
		sql = "select slip_gubun,cost_center,mg_saupbu,company,account,sum(cost) as cost from general_cost where (pl_yn = 'Y') and (cancel_yn = 'N') and (skip_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by slip_gubun,cost_center,mg_saupbu,company,account"
		rs.Open sql, Dbconn, 1
		do until rs.eof	
	
			cost_id = rs("slip_gubun")
			if cost_id = "비용" then
				cost_id = "일반경비"
			end if
			group_name = ""
			bill_trade_name = ""		
			if rs("cost_center") = "상주직접비" then
				sql = "select * from trade where trade_name = '"&rs("company")&"'"
				set rs_trade=dbconn.execute(sql)
				if rs_trade.eof or rs_trade.bof then
					group_name = "Error"
					bill_trade_name = "Error"		
				  else
					group_name = rs_trade("group_name")
					bill_trade_name = rs_trade("bill_trade_name")
				end if		  		  
			end if
	
			sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and cost_center ='"&rs("cost_center")&"' and company ='"&rs("company")&"' and cost_id ='"&cost_id&"' and cost_detail ='"&rs("account")&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"'"
			set rs_cost=dbconn.execute(sql)
		
			sort_seq = 8
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("company")&"','"&bill_trade_name&"','"&group_name&"','"&cost_id&"','"&rs("account")&"','"&rs("mg_saupbu")&"',"&rs("cost")&","&sort_seq&")"
				'dbconn.execute(sql)
			  else
				sum_cost = int(rs_cost("cost")) + Cdbl(rs("cost"))
				sql = "update company_cost set cost_amt_"&cost_month&"="&sum_cost&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and cost_center ='"&rs("cost_center")&"' and company ='"&rs("company")&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and cost_id ='"&cost_id&"' and cost_detail ='"&rs("account")&"' and saupbu ='"&rs("mg_saupbu")&"'"
				'dbconn.execute(sql)
			end if	
			
				Response.write "<br/>" & sql
			rs.movenext()
		loop
		rs.close()
' 비용 SUM 종료
	
' 카드비용 집계
		sql = "select mg_saupbu,cost_center,reside_company as company,account,sum(cost) as cost from card_slip where (pl_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (card_type not like '%주유%' or com_drv_yn = 'Y') group by  mg_saupbu,cost_center,reside_company,account"
		rs.Open sql, Dbconn, 1
		do until rs.eof
								
			group_name = ""
			bill_trade_name = ""		
			if rs("cost_center") = "상주직접비" then
				sql = "select * from trade where trade_name = '"&rs("company")&"'"
				set rs_trade=dbconn.execute(sql)
				if rs_trade.eof or rs_trade.bof then
					group_name = "Error"
					bill_trade_name = "Error"		
				  else
					group_name = rs_trade("group_name")
					bill_trade_name = rs_trade("bill_trade_name")
				end if		  
			end if
	
			sql = "select * from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='법인카드' and cost_detail ='"&rs("account")&"' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into company_cost (cost_year,cost_center,company,group_name,bill_trade_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("company")&"','"&group_name&"','"&bill_trade_name&"','법인카드','"&rs("account")&"','"&rs("mg_saupbu")&"',"&rs("cost")&")"
				'dbconn.execute(sql)
			  else
				sql = "update company_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='법인카드' and cost_detail ='"&rs("account")&"' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
				'dbconn.execute(sql)
			end if		
			
				Response.write "<br/>" & sql
			rs.movenext()
		loop
		rs.close()
' 카드비용 집계 끝


' 차량관리비 집계
' 유류비,주차비,톨비,대중교통비
		sql = "select mg_saupbu,cost_center,company,car_owner,sum(somopum+oil_price+fare+parking+toll) as cost from transit_cost where (cancel_yn = 'N') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by mg_saupbu,cost_center,company,car_owner"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			group_name = ""
			bill_trade_name = ""		
			if rs("cost_center") = "상주직접비" then
				sql = "select * from trade where trade_name = '"&rs("company")&"'"
				set rs_trade=dbconn.execute(sql)
				if rs_trade.eof or rs_trade.bof then
					group_name = "Error"
					bill_trade_name = "Error"		
				  else
					group_name = rs_trade("group_name")
					bill_trade_name = rs_trade("bill_trade_name")
				end if		  
			end if
	
			sql = "select * from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='교통비' and cost_detail ='"&rs("car_owner")&"' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into company_cost (cost_year,cost_center,company,group_name,bill_trade_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("company")&"','"&group_name&"','"&bill_trade_name&"','교통비','"&rs("car_owner")&"','"&rs("mg_saupbu")&"',"&rs("cost")&")"
				'dbconn.execute(sql)
			  else
				sql = "update company_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='교통비' and cost_detail ='"&rs("car_owner")&"' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
				'dbconn.execute(sql)
			end if		
			
				Response.write "<br/>" & sql
			rs.movenext()
		loop
		rs.close()

' 차량수리비
		sql = "select mg_saupbu,cost_center,company,car_owner,sum(repair_cost) as cost from transit_cost where (cancel_yn = 'N') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by mg_saupbu,cost_center,company,car_owner"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			group_name = ""
			bill_trade_name = ""		
			if rs("cost_center") = "상주직접비" then
				sql = "select * from trade where trade_name = '"&rs("company")&"'"
				set rs_trade=dbconn.execute(sql)
				if rs_trade.eof or rs_trade.bof then
					group_name = "Error"
					bill_trade_name = "Error"		
				  else
					group_name = rs_trade("group_name")
					bill_trade_name = rs_trade("bill_trade_name")
				end if		  
			end if
	
			sql = "select * from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='교통비' and cost_detail ='차량수리비' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into company_cost (cost_year,cost_center,company,group_name,bill_trade_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("company")&"','"&group_name&"','"&bill_trade_name&"','교통비','차량수리비','"&rs("mg_saupbu")&"',"&rs("cost")&")"
				'dbconn.execute(sql)
			  else
				sql = "update company_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and bill_trade_name ='"&bill_trade_name&"' and cost_id ='교통비' and cost_detail ='차량수리비' and cost_center ='"&rs("cost_center")&"' and saupbu ='"&rs("mg_saupbu")&"'"
				'dbconn.execute(sql)
			end if		
			
				Response.write "<br/>" & sql
			rs.movenext()
		loop
		rs.close()
								
' 차량관리비 집계 끝
	
' 사업부별 손익 자료 생성
' 처리전 zero
		sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='상주직접비' or cost_center ='직접비') "
		'dbconn.execute(sql)
Response.write "<br>v::"&sql

' 상주직접비 와 직접비 업데이트
		sql = "select saupbu,cost_center,cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '상주직접비' or cost_center = '직접비') and cost_year ='"&cost_year&"' group by saupbu,cost_center,cost_id,cost_detail"
		rs.Open sql, Dbconn, 1
		do until rs.eof

			sql = "select * from saupbu_profit_loss where cost_year ='"&cost_year&"' and saupbu ='"&rs("saupbu")&"' and cost_center ='"&rs("cost_center")&"' and cost_id ='"&rs("cost_id")&"' and cost_detail ='"&rs("cost_detail")&"'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("saupbu")&"','"&rs("cost_center")&"','"&rs("cost_id")&"','"&rs("cost_detail")&"',"&rs("cost")&")"
				'dbconn.execute(sql)
			  else
				sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and saupbu ='"&rs("saupbu")&"' and cost_center ='"&rs("cost_center")&"' and cost_id ='"&rs("cost_id")&"' and cost_detail ='"&rs("cost_detail")&"'"
				'dbconn.execute(sql)
			end if		

				Response.write "<br/>" & sql
			rs.movenext()		
		loop
		rs.close()
' 사업부별 손익 자료 생성 종료

' 회사별별 손익 자료 생성
' 처리전 zero
		sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='상주직접비') "
		'dbconn.execute(sql)
Response.write "<br>x::"&sql

' 상주직접비 업데이트
		sql = "select company,group_name,cost_center,cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '상주직접비') and cost_year ='"&cost_year&"' group by company,group_name,cost_center,cost_id,cost_detail"
		rs.Open sql, Dbconn, 1
		do until rs.eof

			sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&rs("group_name")&"' and cost_center ='"&rs("cost_center")&"' and cost_id ='"&rs("cost_id")&"' and cost_detail ='"&rs("cost_detail")&"'"
			set rs_cost=dbconn.execute(sql)
		
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&rs("group_name")&"','"&rs("cost_center")&"','"&rs("cost_id")&"','"&rs("cost_detail")&"',"&rs("cost")&")"
				'dbconn.execute(sql)
			  else
				sql = "update company_profit_loss set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&rs("group_name")&"' and cost_center ='"&rs("cost_center")&"' and cost_id ='"&rs("cost_id")&"' and cost_detail ='"&rs("cost_detail")&"'"
				'dbconn.execute(sql)
			end if		

				Response.write "<br/>" & sql
			rs.movenext()		
		loop
		rs.close()
' 회사별 손익 자료 생성 종료

		if end_yn = "C" then
			sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
			"' and saupbu = '상주비용'"
		  else
			sql="insert into cost_end (end_month,saupbu,end_yn,batch_yn,bonbu_yn,ceo_yn,reg_id,reg_name,reg_date) values ('"&end_month& _
			"','상주비용','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
		end if
		'dbconn.execute(sql)

				Response.write "<br/>" & sql
		if Err.number <> 0 then
			dbconn.RollbackTrans 
			end_msg = emp_msg + "처리중 Error가 발생하였습니다...."
		else    
			dbconn.CommitTrans
			end_msg = emp_msg + "마감처리 되었습니다...."
		end if
		response.write(now())
	
		response.write"<script language=javascript>"
		response.write"alert('"&end_msg&"');"
		'response.write"location.replace('cost_end_mg.asp');"
		response.write"</script>"
		Response.End
		
		dbconn.Close()
		Set dbconn = Nothing
	end if
%>


