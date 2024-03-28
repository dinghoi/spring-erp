<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

Server.ScriptTimeOut = 1200

org_company	=	Request("org_company")
saupbu		=	Request("saupbu")
end_month	=	Request("end_month")
end_yn		=	Request("end_yn")
'response.write(org_company)

cost_year 	= mid(end_month,1,4)
cost_month 	= mid(end_month,5)

from_date = mid(end_month,1,4) + "-" + mid(end_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
start_date = dateadd("m",-1,from_date)

response.write"<script language=javascript>"
response.write"alert('마감처리중!!!');"
response.write"</script>"

dbconn.BeginTrans


sql = "SELECT * "&_
      "  FROM oil_unit "&_
      " WHERE oil_unit_month = '"&end_month&"'"
Set rs_oil = DbConn.Execute(Sql)
'Response.write Sql & "<br>"


if rs_oil.eof or rs_oil.bof then
	response.write"<script language=javascript>"
	response.write"alert('유류비 단가가 입력되어 있지 않아 마감을 할 수 없습니다.');"
	response.write"location.replace('cost_end_mg.asp');"
	response.write"</script>"
	Response.End
else
' 유류비 단가 및 유류비 계산
	sql = "SELECT * "&_
	      "  FROM transit_cost "&_
	      " WHERE (run_date >='"&from_date&"' AND run_date <='"&to_date&"') "&_
	      "   AND (car_owner ='개인') "&_
	      "   AND (far > 0) "&_
	      "   AND saupbu = '"&saupbu&"'"
	rs.Open sql, Dbconn, 1

	do until rs.eof

		if (rs("team") = "본사팀" or rs("team") = "SM1팀" or rs("team") = "Repair팀" or rs("team") = "SM2팀") then
			oil_unit_id = "1"
		else
			oil_unit_id = "2"
		end if

		sql = "SELECT * "&_
		      "  FROM emp_master_month "&_
		      " WHERE emp_month = '"&end_month&"' "&_
		      "   AND emp_no = '"&rs("mg_ce_id")&"'"
		Set rs_emp=DbConn.Execute(Sql)
'		response.write(sql)
'		response.write(rs_emp("emp_reside_company"))
        if  (not rs_emp.eof) then
            if rs_emp("emp_reside_company") = "한화화약" then
                liter = 8
            else
                liter = 10
            end if
        else
            liter = 10
        end if
		rs_emp.close()

		if rs("oil_kind") = "가스" then
			liter = 7
		end if

		sql = "SELECT * "&_
		      "  FROM oil_unit "&_
		      " WHERE oil_unit_month = '"&end_month&"' "&_
		      "   AND oil_unit_id = '"&oil_unit_id&"' "&_
		      "   AND oil_kind = '"&rs("oil_kind")&"'"
		Set rs_etc=DbConn.Execute(Sql)
		oil_unit_average = rs_etc("oil_unit_average")
		rs_etc.close()

		oil_price = round(int(rs("far")) * oil_unit_average / liter)
		sql = "UPDATE  transit_cost "&_
		      "   SET  oil_unit		=	 "&oil_unit_average	 &_
		      "      , oil_price	=	 "&oil_price	 	 &_
		      " WHERE  mg_ce_id		= '"&rs("mg_ce_id")&"'		"&_
		      "   AND  run_date 	= '"&rs("run_date")&"'		"&_
		      "   AND run_seq 		=	'"&rs("run_seq")&"'"
		dbconn.execute(sql)

		rs.movenext()
	loop
	rs.close()
' 유류비 단가 및 유류비 계산 끝

' 개인별 비용 정산
'	sql = "delete from person_cost where cost_month = '"&end_month&"' and saupbu = '"&saupbu&"'"
'	dbconn.execute(sql)

' 조직 업데이트
	sql = "SELECT * 																		"&_
	      "  FROM emp_master_month 											"&_
	      " WHERE emp_month		= '"&end_month&"' 				"&_
	      "   AND emp_saupbu	= '"&saupbu&"' 						"&_
	      "   AND (   emp_end_date = '1900-01-01' 			"&_
	      "        OR isnull(emp_end_date)							"&_
	      "        OR emp_end_date >= '"&start_date&"') "
  'Response.write sql &"<br>" '--------------------------------------------------------------------------'
	rs_emp.Open sql, Dbconn, 1
	emp_cnt = 0
	do until rs_emp.eof
		emp_cnt = emp_cnt + 1
		' 일반비용
'		sql = "update general_cost set emp_company='"&rs_emp("emp_company")&"',bonbu='"&rs_emp("emp_bonbu")&"',saupbu='"&rs_emp("emp_saupbu")&"',team='"&rs_emp("emp_team")&"',org_name='"&rs_emp("emp_org_name")&"',reside_place='"&rs_emp("emp_reside_place")&"' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (slip_gubun = '비용') and (tax_bill_yn = 'N' or isnull(tax_bill_yn)) and (emp_no='"&rs_emp("emp_no")&"')"
'		dbconn.execute(sql)

		' 교통비
		'Response.write "====================교통비===================="&"<br>"
		sql = "UPDATE  transit_cost																		"&_
		      "   SET  emp_company='"&rs_emp("emp_company")&"'				"&_
		      "      , bonbu='"&rs_emp("emp_bonbu")&"'								"&_
		      "      , saupbu='"&rs_emp("emp_saupbu")&"'							"&_
		      "      , team='"&rs_emp("emp_team")&"'									"&_
		      "      , org_name='"&rs_emp("emp_org_name")&"'					"&_
		      "      , reside_place='"&rs_emp("emp_reside_place")&"'	"&_
		      " WHERE  (    run_date >='"&from_date&"'								"&_
		      "         AND run_date <='"&to_date&"')									"&_
		      "   AND  (mg_ce_id='"&rs_emp("emp_no")&"')"
		dbconn.execute(sql)
		'Response.write sql &"<br>"

		' 야특근
		'Response.write "====================야특근=================="&"<br>"
		sql = "UPDATE  overtime 																				"&_
		      "   SET  emp_company	= '"&rs_emp("emp_company")&"'				"&_
		      "      , bonbu				= '"&rs_emp("emp_bonbu")&"'					"&_
		      "      , saupbu				= '"&rs_emp("emp_saupbu")&"'				"&_
		      "      , team					= '"&rs_emp("emp_team")&"'					"&_
		      "      , org_name			= '"&rs_emp("emp_org_name")&"'			"&_
		      "      , reside_place	= '"&rs_emp("emp_reside_place")&"'	"&_
		      " WHERE  (    work_date >='"&from_date&"'									"&_
		      "         AND work_date <='"&to_date&"')									"&_
		      "   AND  (mg_ce_id		= '"&rs_emp("emp_no")&"')"
		dbconn.execute(sql)
		'Response.write sql &"<br>"

		' 카드전표
		'Response.write "====================카드전표===================="&"<br>"
		sql = "UPDATE  card_slip  "&_
		      "   SET  emp_company		= '"&rs_emp("emp_company")&"'					"&_
		      "      , bonbu					= '"&rs_emp("emp_bonbu")&"'						"&_
		      "      , saupbu					= '"&rs_emp("emp_saupbu")&"'					"&_
		      "      , team						= '"&rs_emp("emp_team")&"'						"&_
		      "      , org_name				= '"&rs_emp("emp_org_name")&"'				"&_
		      "      , reside_place		= '"&rs_emp("emp_reside_place")&"'		"&_
		      "      , reside_company	= '"&rs_emp("emp_reside_company")&"'	"&_
		      " WHERE  (    slip_date >='"&from_date&"'											"&_
		      "         AND slip_date <='"&to_date&"')											"&_
		      "   AND  (emp_no				= '"&rs_emp("emp_no")&"')"
		dbconn.execute(sql)
		'Response.write sql &"<br>"
' 조직 업데이트 끝

' 퇴사 여부 체크
		if (rs_emp("emp_end_date") = "1900-01-01" or isnull(rs_emp("emp_end_date")) or rs_emp("emp_end_date") >= from_date) then
			emp_end = "근무"
		else
			emp_end = "퇴사"
		end if

		' 일반비용
		'Response.write "====================일반비용===================="&"<br>"
		general_cnt = 0
		general_cost = 0
		general_pre_cnt = 0
		general_pre_cost = 0

		sql = "SELECT  pay_yn "&_
		      "      , COUNT(slip_seq) AS c_cnt "&_
		      "      , SUM(cost) AS cost  "&_
		      "  FROM  general_cost  "&_
		      " WHERE  (emp_no='"&rs_emp("emp_no")&"') "& _
		      "   AND  (slip_gubun = '비용')  "&_
		      "   AND  (tax_bill_yn = 'N' OR isnull(tax_bill_yn))  "&_
		      "   AND  (cancel_yn = 'N')  "&_
		      "   AND  (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')  "&_
		      " GROUP  BY pay_yn"
		rs.Open sql, Dbconn, 1
		'Response.write sql & "<br>"
		do until rs.eof
			if rs("pay_yn") = "N" then
				general_cnt  = general_cnt + cint(rs("c_cnt"))
				general_cost = general_cost + cdbl(rs("cost"))
			else
				general_pre_cnt  = general_pre_cnt + cint(rs("c_cnt"))
				general_pre_cost = general_pre_cost + cdbl(rs("cost"))
			end if
			rs.movenext()
		loop
		rs.close()

		' 야특근
		'Response.write "====================야특근===================="&"<br>"
		overtime_cnt = 0
		overtime_cost = 0
		sql = "SELECT  cancel_yn														"&_
		      "      , COUNT(work_date) AS c_cnt						"&_
		      "      , SUM(overtime_amt) AS cost						"&_
		      "  FROM  overtime															"&_
		      " WHERE  (mg_ce_id	='"&rs_emp("emp_no")&"')	"& _
		      "   AND  (    work_date >='"&from_date&"'			"&_
		      "         AND work_date <='"&to_date&"')			"&_
		      "   AND  (cancel_yn = 'N')										"&_
		      " GROUP  BY cancel_yn"
		rs.Open sql, Dbconn, 1
		'Response.Write(sql)

		do until rs.eof
			overtime_cnt  = overtime_cnt + cint(rs("c_cnt"))
			overtime_cost = overtime_cost + cdbl(rs("cost"))
			rs.movenext()
		loop
		rs.close()

		' 교통비
		'Response.write "====================교통비===================="&"<br>"
		gas_km = 0
		gas_unit = 0
		gas_cost = 0
		diesel_km = 0
		diesel_unit = 0
		diesel_cost = 0
		gasol_km = 0
		gasol_unit = 0
		gasol_cost = 0
		somopum_cost = 0
		fare_cnt = 0
		fare_cost = 0
		oil_cash_cost = 0
		repair_cost = 0
		repair_pre_cost = 0
		parking_cost = 0
		toll_cost = 0

		sql = "SELECT *																	"&_
		      "  FROM transit_cost											"&_
		      " WHERE (mg_ce_id='"&rs_emp("emp_no")&"')	"&_
		      "   AND (    run_date >='"&from_date&"'		"&_
		      "        AND run_date <='"&to_date&"')		"&_
		      "   AND (cancel_yn = 'N')"
		rs.Open sql, Dbconn, 1

		do until rs.eof
			if rs("car_owner") = "대중교통" then
				fare_cnt = fare_cnt + 1
				fare_cost = fare_cost + rs("fare")
			end if
			if rs("car_owner") = "개인" then
				if rs("oil_kind") = "휘발유" then
					gasol_km = gasol_km + rs("far")
				end if
				if rs("oil_kind") = "디젤" then
					diesel_km = diesel_km + rs("far")
				end if
				if rs("oil_kind") = "가스" then
					gas_km = gas_km + rs("far")
				end if
			end if

			if rs("car_owner") = "회사" then
				oil_cash_cost = oil_cash_cost + rs("oil_price")
				repair_cost = repair_cost + rs("repair_cost")
			end if

			parking_cost = parking_cost + rs("parking")
			toll_cost = toll_cost + rs("toll")
			rs.movenext()
		loop
		rs.close()
		if rs_emp("emp_team") = "본사팀" or rs_emp("emp_team") = "Repair팀" or rs_emp("emp_team") = "SM1팀" or rs_emp("emp_team") = "SM2팀" then
			oil_unit_id = "1"
		else
			oil_unit_id = "2"
		end if

		sql = "SELECT *																	"&_
		      "  FROM oil_unit													"&_
		      " WHERE oil_unit_month = '"&end_month&"'	"&_
		      "   AND oil_unit_id    = '"&oil_unit_id&"'"
'		response.write(sql)
		rs_etc.Open sql, Dbconn, 1
		do until rs_etc.eof
			if rs_etc("oil_kind") = "휘발유" then
				gasol_unit = rs_etc("oil_unit_average")
			elseif rs_etc("oil_kind") = "가스" then
				gas_unit = rs_etc("oil_unit_average")
			else
				diesel_unit = rs_etc("oil_unit_average")
			end if
			rs_etc.movenext()
		loop
		rs_etc.close()

		if (rs_emp("emp_reside_company") = "한화화약") then
			liter = 8
		else
			liter = 10
		end if

		tot_km = gas_km + diesel_km + gasol_km
		somopum_cost = (tot_km) * 25

		gas_cost = round(gas_km * gas_unit / 7)
		diesel_cost = round(diesel_km * diesel_unit / liter)
		gasol_cost = round(gasol_km * gasol_unit / liter)
		tot_cost = gas_cost + diesel_cost + gasol_cost

		' 주유 카드사용
		juyoo_card_cnt = 0
		juyoo_card_cost = 0
		juyoo_card_cost_vat = 0
		juyoo_card_price = 0
		'Response.write "====================주유 카드사용====================" &"<br>"
		sql = "SELECT  COUNT(*)  AS c_cnt								"&_
		      "      , SUM(cost) AS cost								"&_
		      "      , SUM(cost_vat) AS cost_vat				"&_
		      "  FROM  card_slip												"&_
		      " WHERE  (emp_no ='"&rs_emp("emp_no")&"')	"&_
		      "   AND  (    slip_date >='"&from_date&"' "&_
		      "         AND slip_date <='"&to_date&"')  "&_
		      "   AND  card_type LIKE '%주유%'"
		Set rs = Dbconn.Execute (sql)
		'Response.write sql &"<br>"

		if cint(rs("c_cnt")) <>  0 then
			juyoo_card_cnt = juyoo_card_cnt + cint(rs("c_cnt"))
			juyoo_card_cost = juyoo_card_cost + cdbl(rs("cost"))
			juyoo_card_cost_vat = juyoo_card_cost_vat + cdbl(rs("cost_vat"))
		end if
		rs.close()
		juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat

		' 카드사용
		'Response.write "====================카드사용====================" &"<br>"
		card_cnt = 0
		card_cost = 0
		card_cost_vat = 0
		card_price = 0
		sql = "SELECT  COUNT(*) AS c_cnt								"&_
		      "      , SUM(cost) as cost								"&_
		      "      , SUM(cost_vat) as cost_vat				"&_
		      "  FROM  card_slip												"&_
		      " WHERE  (emp_no ='"&rs_emp("emp_no")&"')	"&_
		      "   AND  (    slip_date >='"&from_date&"' "&_
		      "         AND slip_date <='"&to_date&"')  "&_
		      "   AND  card_type not like '%주유%'"
'		sql = "select count(*) as c_cnt,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (emp_no='"&rs_emp("emp_no")&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
		Set rs = Dbconn.Execute (sql)
		'Response.write sql &"<br>"

		if (cint(rs("c_cnt")) <>  0) then
			card_cnt = card_cnt + cint(rs("c_cnt"))
			card_cost = card_cost + cdbl(rs("cost"))
			card_cost_vat = card_cost_vat + cdbl(rs("cost_vat"))
		end if
		rs.close()
		card_price = card_cost + card_cost_vat

		cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost

' 차량 정보
		sql = "SELECT *  "&_
		      "  FROM car_info  "&_
		      " WHERE owner_emp_no ='"&rs_emp("emp_no")&"'"
		set rs_car=dbconn.execute(sql)
		if rs_car.eof then
			car_owner = "없음"
		else
			car_owner = rs_car("car_owner")
		end if

'		if tot_km <> 0 then
		if car_owner = "개인" then
			return_cash = cash_tot_cost - juyoo_card_price
		else
			return_cash = cash_tot_cost
		end if

		sql = "SELECT *  "&_
		      "  FROM person_cost  "&_
		      " WHERE cost_month ='"&end_month&"'  "&_
		      "   AND emp_no ='"&rs_emp("emp_no")&"'"
		set rs_person=dbconn.execute(sql)
'		Response.write sql & "<br>"
		if rs_person.eof then
			variation_memo = ""
		else
			variation_memo = rs_person("variation_memo")
		end if
		rs_person.close()

		sql = "DELETE FROM person_cost  "&_
		      " WHERE cost_month ='"&end_month&"'  "&_
		      "   AND emp_no ='"&rs_emp("emp_no")&"'"
		dbconn.execute(sql)

		sql = "INSERT INTO person_cost                "&_
		      "VALUES                                 "&_
		      "(  '"&end_month&"'                     "&_
		      " , '"&rs_emp("emp_no")&"'              "&_
		      " , '"&rs_emp("emp_name")&"'            "&_
		      " , '"&rs_emp("emp_job")&"'             "&_
		      " , '"&emp_end&"'                       "&_
		      " , '"&car_owner&"'                     "&_
		      " , '"&rs_emp("emp_company")&"'					"&_
		      " , '"&rs_emp("emp_bonbu")&"'						"&_
		      " , '"&rs_emp("emp_saupbu")&"'					"&_
		      " , '"&rs_emp("emp_team")&"'						"&_
		      " , '"&rs_emp("emp_org_name")&"'				"&_
		      " , '"&rs_emp("emp_reside_place")&"'		"&_
		      " , '"&rs_emp("emp_reside_company")&"'	"&_
		      " ,  "&general_cnt				               &_
		      " ,  "&general_cost				               &_
		      " ,  "&general_pre_cnt		               &_
		      " ,  "&general_pre_cost		               &_
		      " ,  "&overtime_cnt				               &_
		      " ,  "&overtime_cost			               &_
		      " ,  "&gas_km							               &_
		      " ,  "&gas_unit						               &_
		      " ,  "&gas_cost						               &_
		      " ,  "&diesel_km					               &_
		      " ,  "&diesel_unit				               &_
		      " ,  "&diesel_cost				               &_
		      " ,  "&gasol_km						               &_
		      " ,  "&gasol_unit					               &_
		      " ,  "&gasol_cost					               &_
		      " ,  "&tot_km							               &_
		      " ,  "&tot_cost						               &_
		      " ,  "&somopum_cost				               &_
		      " ,  "&fare_cnt						               &_
		      " ,  "&fare_cost					               &_
		      " ,  "&oil_cash_cost			               &_
		      " ,  "&repair_cost				               &_
		      " ,  "&repair_pre_cost		               &_
		      " ,  "&parking_cost				               &_
		      " ,  "&toll_cost					               &_
		      " ,  "&juyoo_card_cnt			               &_
		      " ,  "&juyoo_card_cost		               &_
		      " ,  "&juyoo_card_cost_vat               &_
		      " ,  "&card_cnt 					               &_
		      " ,  "&card_cost 					               &_
		      " ,  "&card_cost_vat 			               &_
		      " ,  "&return_cash 				               &_
		      " , '"&variation_memo&"'	              "&_
		      " , now()                               "&_
		      " , 0)                                  "
		'Response.write sql & "<br>"
		dbconn.execute(sql)


'		if car_owner = "개인" then
'			sql = "update card_slip set skip_yn='Y' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and emp_no ='"&rs_emp("emp_no")&"' "
'			dbconn.execute(sql)
'		end if
		rs_emp.movenext()
	loop
	rs_emp.close()
'개인별 비용전산 끝

' 월별 인사마스터 구성 여부 파악
	if emp_cnt > 0 then

	' 4대보험율과 기타 인건비율 검색
		sql = "SELECT * "&_
		      "  FROM insure_per "&_
		      " WHERE insure_year = '"&cost_year&"'"
		set rs_etc=dbconn.execute(sql)
		insure_tot_per = rs_etc("insure_tot_per")
		income_tax_per = rs_etc("income_tax_per")
		annual_pay_per = rs_etc("annual_pay_per")
		retire_pay_per = rs_etc("retire_pay_per")
		rs_etc.close()

		sql = "UPDATE org_cost "&_
		      "   SET cost_amt_"&cost_month&"= '0' "&_
		      " WHERE cost_year ='"&cost_year&"' "&_
		      "   AND (saupbu = '"&saupbu&"')"
		dbconn.execute(sql)

	' 급여 SUM
		sql = "SELECT  pmg_company "&_
		      "      , pmg_bonbu "&_
		      "      , pmg_saupbu "&_
		      "      , pmg_team "&_
		      "      , pmg_org_name "&_
		      "      , pmg_id "&_
		      "      , SUM(pmg_give_total) AS tot_cost "&_
		      "      , SUM(pmg_base_pay) AS base_pay "&_
		      "      , SUM(pmg_meals_pay) AS meals_pay "&_
		      "      , SUM(pmg_overtime_pay) AS overtime_pay "&_
		      "      , SUM(pmg_research_pay) AS research_pay "&_
		      "      , SUM(pmg_tax_no) AS tax_no  "&_
		      "  FROM  pay_month_give WHERE (pmg_saupbu = '"&saupbu&"')  "&_
		      "   AND  (pmg_yymm ='"&end_month&"')  "&_
		      "   AND  (pmg_id ='1')  "&_
		      " GROUP  BY pmg_company, pmg_bonbu, pmg_saupbu, pmg_team, pmg_org_name"
		rs.Open sql, Dbconn, 1

		do until rs.eof

			sort_seq = 0
			cost_detail = "급여"

			sql = "SELECT *  																		"&_
		      "    FROM org_cost															"&_
		      "   WHERE cost_year ='"&cost_year&"'						"&_
		      "     AND emp_company ='"&rs("pmg_company")&"'	"&_
		      "     AND bonbu ='"&rs("pmg_bonbu")&"'					"&_
		      "     AND saupbu ='"&rs("pmg_saupbu")&"'				"&_
		      "     AND team ='"&rs("pmg_team")&"'						"&_
		      "     AND org_name ='"&rs("pmg_org_name")&"'		"&_
		      "     AND cost_id ='인건비'												"&_
		      "     AND cost_detail ='"&cost_detail&"'"
			set rs_cost=dbconn.execute(sql)
			'Response.write sql & "<br>"

			if rs_cost.eof or rs_cost.bof then
				sql = "INSERT INTO org_cost					"&_
		          "(  cost_year 								"&_
		          " , emp_company								"&_
		          " , bonbu											"&_
		          " , saupbu										"&_
		          " , team											"&_
		          " , org_name									"&_
		          " , cost_id										"&_
		          " , cost_detail								"&_
		          " , cost_amt_"&cost_month  		 &_
		          " , sort_seq 									"&_
		          ")  													"&_
		          "VALUES  											"&_
		          "(  '"&cost_year&"' 					"&_
		          " , '"&rs("pmg_company")&"'		"&_
		          " , '"&rs("pmg_bonbu")&"' 		"&_
		          " , '"&rs("pmg_saupbu")&"'		"&_
		          " , '"&rs("pmg_team")&"'			"&_
		          " , '"&rs("pmg_org_name")&"'	"&_
		          " , '인건비' 										"&_
		          " , '"&cost_detail&"' 				"&_
		          " , "&rs("tot_cost") 					 &_
		          " , "&sort_seq 								 &_
		          ")"
				dbconn.execute(sql)
			else
				sql = "UPDATE  org_cost  																"&_
		          "   SET  cost_amt_"&cost_month&"="&rs("tot_cost")	 &_
		          "      , sort_seq="&sort_seq											 &_
		          " WHERE  cost_year ='"&cost_year&"'  							"&_
		          "   AND  emp_company = '"&rs("pmg_company")&"'  	"&_
		          "   AND  bonbu ='"&rs("pmg_bonbu")&"'  						"&_
		          "   AND  saupbu ='"&rs("pmg_saupbu")&"'  					"&_
		          "   AND  team ='"&rs("pmg_team")&"'  							"&_
		          "   AND  org_name ='"&rs("pmg_org_name")&"'  			"&_
		          "   AND  cost_id ='인건비'  												"&_
		          "   AND  cost_detail ='"&cost_detail&"'"
				dbconn.execute(sql)
			end if
	' 2015-04-27
	' 4대보험료
            'insure_tot = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * insure_tot_per / 100)
            insure_tot = clng((clng(rs("tot_cost"))) * insure_tot_per / 100)
			sort_seq = 2

			sql = "SELECT * 																		"&_
			      "  FROM org_cost 															"&_
			      " WHERE cost_year ='"&cost_year&"'						"&_
			      "   AND emp_company ='"&rs("pmg_company")&"'	"&_
			      "   AND bonbu ='"&rs("pmg_bonbu")&"'					"&_
			      "   AND saupbu ='"&rs("pmg_saupbu")&"'				"&_
			      "   AND team ='"&rs("pmg_team")&"'						"&_
			      "   AND org_name ='"&rs("pmg_org_name")&"'		"&_
			      "   AND cost_id ='인건비'												"&_
			      "   AND cost_detail ='4대보험'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "INSERT INTO org_cost "&_
				      "( "&_
				      "   cost_year "&_
				      " , emp_company "&_
				      " , bonbu "&_
				      " , saupbu "&_
				      " , team "&_
				      " , org_name "&_
				      " , cost_id "&_
				      " , cost_detail "&_
				      " , cost_amt_"&cost_month &_
				      " , sort_seq "&_
				      ") "&_
				      "VALUES  "&_
				      "(  '"&cost_year&"' "&_
				      " , '"&rs("pmg_company")&"' "&_
				      " , '"&rs("pmg_bonbu")&"' "&_
				      " , '"&rs("pmg_saupbu")&"' "&_
				      " , '"&rs("pmg_team")&"' "&_
				      " , '"&rs("pmg_org_name")&"' "&_
				      " , '인건비' "&_
				      " , '4대보험' "&_
				      " , "&insure_tot &_
				      " , "&sort_seq &_
				      ")"
				dbconn.execute(sql)
			else
				sql = "UPDATE  org_cost																"&_
				      "   SET  cost_amt_"&cost_month&" = "&insure_tot  &_
				      "      , sort_seq = "&sort_seq 									 &_
				      " WHERE  cost_year = '"&cost_year&"'  					"&_
				      "   AND  emp_company = '"&rs("pmg_company")&"'	"&_
				      "   AND  bonbu = '"&rs("pmg_bonbu")&"'					"&_
				      "   AND  saupbu ='"&rs("pmg_saupbu")&"'					"&_
				      "   AND  team ='"&rs("pmg_team")&"'							"&_
				      "   AND  org_name ='"&rs("pmg_org_name")&"'			"&_
				      "   AND  cost_id ='인건비'												"&_
				      "   AND  cost_detail ='4대보험'"
				dbconn.execute(sql)
			end if

		' 소득세 종업원분
            'income_tax = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * income_tax_per / 100)
            income_tax = clng((clng(rs("tot_cost"))) * income_tax_per / 100)
			sort_seq = 3

			sql = "SELECT *																			"&_
			      "  FROM org_cost															"&_
			      "  WHERE cost_year ='"&cost_year&"'						"&_
			      "    AND emp_company ='"&rs("pmg_company")&"'	"&_
			      "    AND bonbu ='"&rs("pmg_bonbu")&"'					"&_
			      "    AND saupbu ='"&rs("pmg_saupbu")&"'				"&_
			      "    AND team ='"&rs("pmg_team")&"'						"&_
			      "    AND org_name ='"&rs("pmg_org_name")&"'		"&_
			      "    AND cost_id ='인건비'											"&_
			      "    AND cost_detail ='소득세종업원분'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "INSERT INTO org_cost "&_
				      "( "&_
				      "   cost_year "&_
				      " , emp_company "&_
				      " , bonbu "&_
				      " , saupbu "&_
				      " , team "&_
				      " , org_name "&_
				      " , cost_id "&_
				      " , cost_detail "&_
				      " , cost_amt_"&cost_month	 &_
				      " , sort_seq							"&_
				      ")												"&_
				      "VALUES										"&_
				      "( "&_
				      "   '"&cost_year&"' "&_
				      " , '"&rs("pmg_company")&"' "&_
				      " , '"&rs("pmg_bonbu")&"' "&_
				      " , '"&rs("pmg_saupbu")&"' "&_
				      " , '"&rs("pmg_team")&"' "&_
				      " , '"&rs("pmg_org_name")&"' "&_
				      " , '인건비' "&_
				      " , '소득세종업원분' "&_
				      " , "&income_tax &_
				      " , "&sort_seq &_
				      ")"
				dbconn.execute(sql)
			  else
				sql = "UPDATE org_cost SET cost_amt_"&cost_month&"="&income_tax&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='인건비' and cost_detail ='소득세종업원분'"
				dbconn.execute(sql)
			end if
		' 연차수당
			annual_pay = clng((clng(rs("base_pay"))+clng(rs("meals_pay"))+clng(rs("overtime_pay"))) * annual_pay_per / 100)
			sort_seq = 4

			sql = "SELECT * FROM org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='인건비' and cost_detail ='연차수당'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("pmg_company")&"','"&rs("pmg_bonbu")&"','"&rs("pmg_saupbu")&"','"&rs("pmg_team")&"','"&rs("pmg_org_name")&"','인건비','연차수당',"&annual_pay&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&annual_pay&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='인건비' and cost_detail ='연차수당'"
				dbconn.execute(sql)
			end if
		' 퇴직충당금
			retire_pay = clng((clng(rs("base_pay"))+clng(rs("meals_pay"))+clng(rs("overtime_pay"))) * retire_pay_per / 100)
			sort_seq = 5

			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='인건비' and cost_detail ='퇴직충당금'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("pmg_company")&"','"&rs("pmg_bonbu")&"','"&rs("pmg_saupbu")&"','"&rs("pmg_team")&"','"&rs("pmg_org_name")&"','인건비','퇴직충당금',"&retire_pay&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&retire_pay&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='인건비' and cost_detail ='퇴직충당금'"
				dbconn.execute(sql)
			end if

	' 2015-04-27 End
			rs.movenext()
		loop
		rs.close()
	' 상여 SUM
		sql = "select pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name,pmg_id,sum(pmg_give_total) as cost from pay_month_give where (pmg_saupbu = '"&saupbu&"') and (pmg_yymm ='"&end_month&"') and (pmg_id ='2') group by pmg_company,pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_name,pmg_id"
		rs.Open sql, Dbconn, 1
		do until rs.eof

			sort_seq = 1
			cost_detail = "상여"

			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='인건비' and cost_detail ='"&cost_detail&"'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("pmg_company")&"','"&rs("pmg_bonbu")&"','"&rs("pmg_saupbu")&"','"&rs("pmg_team")&"','"&rs("pmg_org_name")&"','인건비','"&cost_detail&"',"&rs("cost")&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("pmg_company")&"' and bonbu ='"&rs("pmg_bonbu")&"' and saupbu ='"&rs("pmg_saupbu")&"' and team ='"&rs("pmg_team")&"' and org_name ='"&rs("pmg_org_name")&"' and cost_id ='인건비' and cost_detail ='"&cost_detail&"'"
				dbconn.execute(sql)
			end if
			rs.movenext()
		loop
		rs.close()

	' 알바비
		sql = "select company,bonbu,saupbu,team,org_name,sum(alba_give_total) as cost from pay_alba_cost where (saupbu = '"&saupbu&"') and (rever_yymm ='"&end_month&"') group by company,bonbu,saupbu,team,org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof

			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='인건비' and cost_detail ='알바비'"
			set rs_cost=dbconn.execute(sql)

			sort_seq = 8
			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','인건비','알바비',"&rs("cost")&","&sort_seq&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='인건비' and cost_detail ='알바비'"
				dbconn.execute(sql)
			end if
			rs.movenext()
		loop
		rs.close()

	'야특근 마감
		sql = "UPDATE overtime											"&_
		      "   SET end_yn='Y'										"&_
		      " WHERE work_date >= '"&from_date&"'	"&_
		      "   AND work_date <= '"&to_date&"'		"&_
		      "   AND saupbu 		='"&saupbu&"'"
		dbconn.execute(sql)

	'일반비용
		sql = "UPDATE general_cost "&_
		      "   SET end_yn='Y' "&_
		      " WHERE (slip_date >= '"&from_date&"' AND slip_date <= '"&to_date&"') "&_
		      "   AND saupbu ='"&saupbu&"'"
		dbconn.execute(sql)
	' DB SUM 처리 (비용)
		sql = "SELECT  emp_company "&_
		      "      , bonbu "&_
		      "      , saupbu "&_
		      "      , team "&_
		      "      , org_name "&_
		      "      , account "&_
		      "      , SUM(cost) AS cost  "&_
		      "  FROM  general_cost  "&_
		      " WHERE  (slip_gubun = '비용')  "&_
		      "   AND  (cancel_yn = 'N')  "&_
		      "   AND  (saupbu = '"&saupbu&"')  "&_
		      "   AND  (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"')  "&_
		      " GROUP  BY emp_company, bonbu, saupbu, team, org_name, account"
		rs.Open sql, Dbconn, 1
		do until rs.eof

			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='일반경비' and cost_detail ='"&rs("account")&"'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','일반경비','"&rs("account")&"',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='일반경비' and cost_detail ='"&rs("account")&"'"
				dbconn.execute(sql)
			end if
			rs.movenext()
		loop
		rs.close()
	' DB SUM 처리 (비용 외)
		sql = "select slip_gubun,emp_company,bonbu,saupbu,team,org_name,account,sum(cost) as cost from general_cost where (slip_gubun <> '비용') and (cancel_yn = 'N') and (saupbu = '"&saupbu&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by slip_gubun,emp_company,bonbu,saupbu,team,org_name,account"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			cost_id = rs("slip_gubun")
			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='"&cost_id&"' and cost_detail ='"&rs("account")&"'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','"&cost_id&"','"&rs("account")&"',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='"&cost_id&"' and cost_detail ='"&rs("account")&"'"
				dbconn.execute(sql)
'				Response.write sql
			end if
			rs.movenext()
		loop
		rs.close()

	'교통비
		sql = "Update transit_cost set end_yn='Y' where (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and saupbu ='"&saupbu&"'"
		dbconn.execute(sql)

	' DB SUM 교통비
		sql = "select emp_company,bonbu,saupbu,team,org_name,car_owner,sum(somopum+oil_price+fare+parking+toll) as cost from transit_cost where (cancel_yn = 'N') and (saupbu = '"&saupbu&"') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by emp_company,bonbu,saupbu,team,org_name,car_owner"
		rs.Open sql, Dbconn, 1
		do until rs.eof

			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='교통비' and cost_detail ='"&rs("car_owner")&"'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','교통비','"&rs("car_owner")&"',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='교통비' and cost_detail ='"&rs("car_owner")&"'"
				dbconn.execute(sql)
			end if
			rs.movenext()
		loop
		rs.close()

	' DB SUM 교통비 (차량수리비)
		sql = "select emp_company,bonbu,saupbu,team,org_name,sum(repair_cost) as cost from transit_cost where (cancel_yn = 'N') and (repair_cost > 0) and (saupbu = '"&saupbu&"') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by emp_company,bonbu,saupbu,team,org_name"
		rs.Open sql, Dbconn, 1
		do until rs.eof

			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='교통비' and cost_detail ='차량수리비'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','교통비','차량수리비',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='교통비' and cost_detail ='차량수리비'"
				dbconn.execute(sql)
			end if
			rs.movenext()
		loop
		rs.close()

' 회사 차량 운행 주유카드 셋팅
		sql = "select mg_ce_id from transit_cost where (car_owner = '회사') and (saupbu = '"&saupbu&"') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by mg_ce_id"
		rs.Open sql, Dbconn, 1
		do until rs.eof
			sql = "update card_slip set com_drv_yn='Y' where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_no='"&rs("mg_ce_id")&"')"
			dbconn.execute(sql)

			rs.movenext()
		loop
		rs.close

	' 카드비용 집계
	'	sql = "select emp_company,bonbu,saupbu,team,org_name,reside_place,reside_company,account,sum(cost) as cost from card_slip where (end_sw = 'Y') and (saupbu = '"&saupbu&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,bonbu,saupbu,team,org_name,reside_place,reside_company,account"
		sql = "select owner_company as emp_company,bonbu,saupbu,team,org_name,account,sum(cost) as cost from card_slip where (saupbu = '"&saupbu&"') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (card_type not like '%주유%' or com_drv_yn = 'Y')  group by owner_company,bonbu,saupbu,team,org_name,account"
		rs.Open sql, Dbconn, 1
		do until rs.eof

			sql = "select * from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")& _
			"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='법인카드' and cost_detail ='"&rs("account")&"'"
			set rs_cost=dbconn.execute(sql)

			if rs_cost.eof or rs_cost.bof then
				sql = "insert into org_cost (cost_year,emp_company,bonbu,saupbu,team,org_name,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','법인카드','"&rs("account")&"',"&rs("cost")&")"
				dbconn.execute(sql)
			  else
	'			sum_cost = clng(rs("cost")) + clng(rs_cost(9+cost_month))
				sql = "update org_cost set cost_amt_"&cost_month&"="&rs("cost")&" where cost_year ='"&cost_year&"' and emp_company ='"&rs("emp_company")&"' and bonbu ='"&rs("bonbu")&"' and saupbu ='"&rs("saupbu")&"' and team ='"&rs("team")&"' and org_name ='"&rs("org_name")&"' and cost_id ='법인카드' and cost_detail ='"&rs("account")&"'"
				dbconn.execute(sql)
			end if
			rs.movenext()
		loop
		rs.close()

	end if

		if end_yn = "C" then
			sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
			"' and saupbu = '"&saupbu&"'"
		else
			sql="DELETE FROM cost_end WHERE end_month = '"&end_month&"' and saupbu = '"&saupbu&"' "
			dbconn.execute(sql)

			sql="insert into cost_end (end_month,saupbu,end_yn,batch_yn,bonbu_yn,ceo_yn,reg_id,reg_name,reg_date) values ('"&end_month& _
			"','"&saupbu&"','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
		end if
		dbconn.execute(sql)

	if emp_cnt = 0 then
		emp_msg = "인사마스터 마감이 되지 않았습니다 "
	  else
		emp_msg = ""
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans
		end_msg = emp_msg + "처리중 Error가 발생하였습니다...."
	else
		dbconn.CommitTrans
		end_msg = emp_msg + "마감처리 되었습니다...."
	end if

	rs_oil.close()

	dbconn.Close()
	Set dbconn = Nothing

	Response.write"<script language=javascript>"
	Response.write"alert('"&end_msg&"');"
	Response.write"location.replace('cost_end_mg.asp');"
	Response.write"</script>"
	Response.End

end if

%>
