<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<%
	id = request.form("id")
	pass = request.form("pass")
	save_id = request.form("save_id")
			
	if	id = "" or isnull(id) or id = " " or (id > "700000" and id < "799999") then		
		response.write"<script language=javascript>"
		response.write"alert('아이디가 등록되어 있지 않습니다....');"		
		response.write"history.go(-1);"
		response.write"</script>"
	end if
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open DbConnect

    ' 개행문자 제거 박영주 요청 2019-06-04
    ' 가끔 들어가는 경우가 있어 아예 로그인시 전체처리.. as등록시 작업자 추가에서 오류가 남..
    sql="UPDATE memb SET org_name = REPLACE(org_name, '\r', '') WHERE org_name LIKE '%\r%' " 
    dbconn.execute(sql)

    sql="UPDATE memb SET org_name = REPLACE(org_name, '\n', '') WHERE org_name LIKE '%\n%' "
    dbconn.execute(sql)
    ' 개행문자 제거 박영주 요청 2019-06-04 

	sql="select * from memb where user_id='"&id&"'"
	set rs=dbconn.execute(sql)
	
	if	rs.eof or rs.bof then		
		response.write"<script language=javascript>"
		response.write"alert('아이디가 등록되어 있지 않습니다....');"		
		response.write"history.go(-1);"
		response.write"</script>"
	ElseIf rs("pass") <> pass Then
		response.write"<script language=javascript>"
		response.write"alert('비밀번호가 다릅니다....');"		
		response.write"history.go(-1);"
		response.write"</script>"	
	ElseIf rs("grade") = "6" Then
		response.write"<script language=javascript>"
		response.write"alert('NKP 사용권한이 없거나 조직변경으로 사번이 바뀔수 있습니다.');"		
		response.write"history.go(-1);"
		response.write"</script>"	
	ElseIf rs("grade") = "" or isnull(rs("grade")) Then
		response.write"<script language=javascript>"
		response.write"alert('로그인 할수 없습니다. 관리자에게 문의 바랍니다....');"		
		response.write"history.go(-1);"
		response.write"</script>"	
	Else	
		Response.Cookies("nkpmg_user")("coo_user_id") = rs("user_id")
		Response.Cookies("nkpmg_user")("coo_user_name") = rs("user_name")
		Response.Cookies("nkpmg_user")("coo_emp_no") = rs("emp_no")
		
		if isnull(rs("emp_company")) then
			Response.Cookies("nkpmg_user")("coo_emp_company") = ""
		else			
			Response.Cookies("nkpmg_user")("coo_emp_company") = rs("emp_company")
		end if
		if isnull(rs("saupbu")) then
			Response.Cookies("nkpmg_user")("coo_saupbu") = ""
		else			
			Response.Cookies("nkpmg_user")("coo_saupbu") = rs("saupbu")
		end if
		if isnull(rs("bonbu")) then
			Response.Cookies("nkpmg_user")("coo_bonbu") = ""
		else			
			Response.Cookies("nkpmg_user")("coo_bonbu") = rs("bonbu")
		end if
		if isnull(rs("team")) then
			Response.Cookies("nkpmg_user")("coo_team") = ""
		else			
			Response.Cookies("nkpmg_user")("coo_team") = rs("team")
		end if
		if isnull(rs("reside_place")) then
			Response.Cookies("nkpmg_user")("coo_reside_place") = ""
		else			
			Response.Cookies("nkpmg_user")("coo_reside_place") = rs("reside_place")
		end if
		if isnull(rs("reside_company")) then
			Response.Cookies("nkpmg_user")("coo_reside_company") = ""
		else			
			Response.Cookies("nkpmg_user")("coo_reside_company") = rs("reside_company")
		end if
		Response.Cookies("nkpmg_user")("coo_hp") = rs("hp")
		Response.Cookies("nkpmg_user")("coo_grade") = rs("grade")
		Response.Cookies("nkpmg_user")("coo_cost_grade") = rs("cost_grade")
		Response.Cookies("nkpmg_user")("coo_insa_grade") = rs("insa_grade")
		Response.Cookies("nkpmg_user")("coo_pay_grade") = rs("pay_grade")
		Response.Cookies("nkpmg_user")("coo_met_grade") = rs("met_grade")
		Response.Cookies("nkpmg_user")("coo_account_grade") = rs("account_grade")
		Response.Cookies("nkpmg_user")("coo_sales_grade") = rs("sales_grade")
		Response.Cookies("nkpmg_user")("coo_user_grade") = rs("user_grade")
		Response.Cookies("nkpmg_user")("coo_mg_group") = rs("mg_group")
		Response.Cookies("nkpmg_user")("coo_reside") = rs("reside")
		Response.Cookies("nkpmg_user")("coo_help_yn") = rs("help_yn")
		
		' email을 검사해서 @ 앞을것을 그룹웨어의 id 로 삼는다.
		if isnull(rs("email")) then
  		Response.Cookies("nkpmg_user")("coo_groupware_id") = ""
  	else
  	  if (rs("email")="") then
  	    Response.Cookies("nkpmg_user")("coo_groupware_id") = ""
  	  else
  	    Response.Cookies("nkpmg_user")("coo_groupware_id") = split(rs("email"),"@")(0)
  	  end if
		end if		
		
		if isnull(rs("position")) then
			Response.Cookies("nkpmg_user")("coo_position") = ""
		else			
			Response.Cookies("nkpmg_user")("coo_position") = rs("position")
		end if
		if isnull(rs("org_name")) then
			Response.Cookies("nkpmg_user")("coo_org_name") = ""
		else			
			Response.Cookies("nkpmg_user")("coo_org_name") = rs("org_name")	
		end if
'		Response.Cookies("nkpmg_user")("coo_org_name") = rs("org_name")		
		if isnull(rs("asset_company")) then
			Response.Cookies("nkpmg_user")("coo_asset_company") = "00"	
		else
			Response.Cookies("nkpmg_user")("coo_asset_company") = rs("asset_company")		
		end if

		login_cnt = int(rs("login_cnt")) + 1
		sql = "update memb set login_cnt="&login_cnt&", login_date=now() where user_id='"&id&"'"
		dbconn.execute(sql)	  

		if (rs("mg_group") > "5" or rs("grade") = "5" or rs("team") = "외주관리") and isnull(rs("asset_company")) then
			response.write"<script language=javascript>"
			response.write"location.replace('as_list_ce_user.asp');"
			response.write"</script>"			
		  elseIf rs("grade") = "5" and isnull(rs("asset_company"))  Then
			response.write"<script language=javascript>"
			response.write"location.replace('as_list_ce_user.asp');"
			response.write"</script>"			
		  elseIf rs("grade") = "5" and rs("asset_company") > "00" and rs("reside") = "2" Then
			response.write"<script language=javascript>"
			response.write"location.replace('asset_process_mg.asp');"
			response.write"</script>"			
		  elseIf rs("grade") = "5" and rs("asset_company") > "00" and rs("reside") = "3" Then
			response.write"<script language=javascript>"
			response.write"location.replace('as_list_asset.asp');"
			response.write"</script>"			
		  elseIf rs("grade") = "5" and rs("asset_company") > "00" and rs("reside") < "2" Then
			response.write"<script language=javascript>"
'				response.write"location.replace('k1_asset_user_sum.asp');"
			response.write"location.replace('as_list_asset.asp');"
			response.write"</script>"			
		  else
			response.write"<script language=javascript>"
			response.write"location.replace('nkp_main.asp?first_sw=y');"
			response.write"</script>"			
		end if
	end if
	
	Response.End
	dbconn.Close()
	Set dbconn = Nothing
	
%>
