<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

	dim title_name
	dim company_tab(50)

	title_name = array("접수번호", "접수일자", "접수자", "사용자", "협업구분", "전화번호", "핸드폰", "회사", "조직명", "주소", "CE명", "CE사번", "CE소속팀", "장애내역", "요청일", "요청시간", "처리일", "처리시간", "진행", "처리방법", "고객요청", "입고/지연사유", "입고일자", "대체여부", "메이커", "장애장비", "자산코드", "모델명", "S/N번호", "처리내용", "설치수량", "PC S/W", "PC H/W", "모니터", "프린터/스케너", "통신장비", "서버/워크", "아답터", "기타")
	
	from_date = request("from_date")
	to_date = request("to_date")
	company = request("company")
	date_sw = request("date_sw")
	process_sw = request("process_sw")
	field_check = request("field_check")
	field_view = request("field_view")
	savefilename = from_date + to_date + ".xls"


'Response.write from_date & "<br>" & to_date & "<br>" & company & "<br>" & date_sw & "<br>" & process_sw & "<br>" & field_check & "<br>" & field_view & "<br>" & savefilename  & "<br>"
'Response.end

	Response.Buffer = True
	Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect

	if c_grade = "7" then
		k = 0
		Sql="SELECT * FROM etc_code WHERE etc_type = '51' AND used_sw = 'Y' AND group_name = '"+user_name+"' ORDER BY etc_name ASC"
		rs_etc.Open Sql, Dbconn, 1
		while not rs_etc.eof
			k = k + 1
			company_tab(k) = rs_etc("etc_name")
			rs_etc.movenext()
		Wend
		rs_etc.close()						
	end if

	' 2018-03-06 as_acpt.mg_ce_id 소속 정보 표시 from emp_master
	base_sql = "SELECT  A.acpt_no " & _
	           "      , A.acpt_date " & _
	           "      , A.acpt_man " & _
	           "      , A.acpt_user " & _
	           "      , CASE WHEN ifnull(A.cowork_yn, 'N') = 'N' THEN 'NO' WHEN ifnull(A.cowork_yn, 'N') = 'Y' THEN 'YES' END AS cowork "&_
	           "      , concat(A.tel_ddd,'-',A.tel_no1,'-',A.tel_no2) " & _
	           "      , concat(A.hp_ddd,'-',A.hp_no1,'-',A.hp_no2) " & _
	           "      , A.company " & _
	           "      , A.dept " & _
	           "      , concat(A.sido,' ',A.gugun,' ',A.dong,' ',A.addr) " & _
	           "      , A.mg_ce,A.mg_ce_id " & _
             "      , (SELECT emp_org_name FROM emp_master WHERE emp_no = A.mg_ce_id ) AS emp_org_name " & _
	           "      , A.as_memo " & _
	           "      , A.request_date " & _
	           "      , A.request_time " & _
	           "      , A.visit_date,A.visit_time " & _
	           "      , A.as_process        " & _ 
	           "      , A.as_type           " & _ 
	           "      , A.visit_request_yn  " & _ 
	           "      , A.into_reason       " & _ 
	           "      , A.in_date           " & _ 
	           "      , A.in_replace        " & _ 
	           "      , A.maker             " & _ 
	           "      , A.as_device         " & _ 
	           "      , A.asets_no          " & _ 
	           "      , A.model_no          " & _ 
	           "      , A.serial_no         " & _ 
	           "      , A.as_history        " & _ 
	           "      , A.dev_inst_cnt      " & _ 
	           "      , A.err_pc_sw         " & _ 
	           "      , A.err_pc_hw         " & _ 
	           "      , A.err_monitor       " & _ 
	           "      , A.err_printer       " & _ 
	           "      , A.err_network       " & _ 
	           "      , A.err_server        " & _ 
	           "      , A.err_adapter       " & _ 
	           "      , A.err_etc           " & _ 
	           "  FROM  as_acpt A           "

	'base_sql = base_sql + "   A.as_process        " & _
	'                      " , A.as_type           " & _
	'                      " , A.visit_request_yn  " & _
	'                      " , A.into_reason       " & _
	'                      " , A.in_date           " & _
	'                      " , A.in_replace        " & _
	'                      " , A.maker             " & _
	'                      " , A.as_device         " & _
	'                      " , A.asets_no          " & _
	'                      " , A.model_no          " & _
	'                      " , A.serial_no         " & _
	'                      " , A.as_history        " & _
	'                      " , A.dev_inst_cnt      " & _
	'                      " , A.err_pc_sw         " & _
	'                      " , A.err_pc_hw         " & _
	'                      " , A.err_monitor       " & _
	'                      " , A.err_printer       " & _
	'                      " , A.err_network       " & _
	'                      " , A.err_server        " & _
	'                      " , A.err_adapter       " & _
	'                      " , A.err_etc           " & _
	'                      " FROM  as_acpt A "
	                  'INNER JOIN emp_master B on A.mg_ce_id = B.emp_no "

'Response.write "date_sw: " & date_sw & "<br> process_sw: " & process_sw & "<br> field_check: " & field_check & "<br> c_grade: " & c_grade & "<br> company: " & company & "<br>"
'date_sw: acpt
'process_sw: A
'field_check: total
'c_grade: 4
'company: 전체


	if date_sw = "acpt" then
'		if c_grade = "7" or c_grade = "8" then
			date_sql = "where (cast(A.acpt_date as date) >= '" + from_date  + "' and cast(A.acpt_date as date) <= '" + to_date  + "')"
'		  else
'			date_sql = "where (cast(acpt_date as date) >= '" + from_date  + "' and cast(acpt_date as date) <= '" + to_date  + "') and (mg_group ='" + mg_group + "')"
'		end if
	else
'		if c_grade = "7" or c_grade = "8" then
			date_sql = "where (A.visit_date >= '" + from_date  + "' and A.visit_date <= '" + to_date  + "')"
'		  else
'			date_sql = "where (visit_date >= '" + from_date  + "' and visit_date <= '" + to_date  + "') and (mg_group ='" + mg_group + "')"
'		end if
	end if
	
	if process_sw = "A" then
		process_sql = " and ( A.as_process = '완료' or A.as_process = '대체' or A.as_process = '취소' or A.as_process = '접수' or A.as_process = '연기' or A.as_process = '입고' or as_process = '대체입고' ) "
	elseif process_sw = "Y" then
		process_sql = " and ( A.as_process = '완료' or A.as_process = '대체' or A.as_process = '취소') "
	else
		process_sql = " and ( A.as_process = '접수' or A.as_process = '연기' or A.as_process = '입고' or as_process = '대체입고') "
	end if
	
	if field_check <> "total" then
		field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) ORDER BY A.acpt_date DESC"
	else
		field_sql = " ORDER BY A.acpt_date DESC"
	end if
	sql = base_sql + date_sql + process_sql + field_sql

	if c_grade = "7" then
		com_sql = "A.company = '" + company_tab(1) + "'"	
		for kk = 2 to k
			com_sql = com_sql + " or A.company = '" + company_tab(kk) + "'"
		next
		sql = base_sql + date_sql + " and (" + com_sql + ") " + process_sql + field_sql
	end if

	if c_grade = "8" then
		com_sql = " and (company = '" + user_name + "') "
		sql = base_sql + date_sql + com_sql + process_sql + field_sql
	end if

	if company = "전체" then
		sql = sql
	else
		com_sql = " and (A.company = '" + company + "') "
		sql = base_sql + date_sql + com_sql + process_sql + field_sql		
	end if

	Rs.Open Sql, Dbconn, 1

'Response.write Sql
'Response.end
	
	if rs.eof then
		response.write"<script language=javascript>"
		response.write"alert('다운 할 자료가 없습니다 ....');"
		response.write"history.go(-1);"
		response.write"</script>"	
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
		<title></title>
		</head>
		<body>
			<table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>
				<tr>
					<%=chr(13)&chr(10)%>
					<%	
						i = 0
						for each whatever in rs.fields
'						Response.write i&": " & i
							if i < 38 then
					%>
					<td><b><%=title_name(i)%></b></TD>
					<%=chr(13)&chr(10)%>
					<%	
							end if
							i = i + 1
						next
					%>
				</tr>
				<%=chr(13)&chr(10)%>
				<%
					alldata=rs.getrows
					numcols=ubound(alldata,1)
					numrows=ubound(alldata,2)
					
					jj = 0
					FOR j= 0 TO numrows 
						jj = jj + 1
						if	jj > 1500 then
							jj = 0
							response.flush
						end if
				%>
				<tr>
					<%=chr(13)&chr(10)%>
					<%  FOR i=0 to numcols
								if i > 37 then
									exit for
								end if
								thisfield = alldata(i,j)
								'Response.write thisfield &"<BR>"
								
								if i = 20 then
									if thisfield = "Y" then
										thisfield = "방문요청"
									else
										thisfield = ""
									end if
								end if

					      if isnull(thisfield) then
					         thisfield=""
					      end if
					      if trim(thisfield)="" then
					         thisfield=""
					      end if
					      
					      err_memo = ""
					      
					      if i > 29  and i < 38 then
					      	if thisfield <> "" then
					      		for k = 1 to 100 step 6
					      			chkfield = mid(thisfield,k,4)
					      			if chkfield = "" or chkfield= null then
					      				exit for
					      			end if
					      			
					      			sql_etc = "SELECT * FROM etc_code WHERE etc_code = '" + chkfield +"'"
					      			Set Rs_etc=dbconn.execute(Sql_etc)
					      			
					      			if rs_etc.eof or rs_etc.bof then
					      				etc_name = ""
					      			else
					      				etc_name = rs_etc("etc_name")
					      				
					      				if err_memo = "" then
					      					err_memo = etc_name
					      				else
					      					err_memo = err_memo + "," +etc_name
					      				end if
					      	end if
					      	rs_etc.close()
					    next
					  end if		
						thisfield = err_memo
				end if
			%>
			<%	if i = 1 then %>
			<td valign=top>
			<%=thisfield%>
		</td><%=chr(13)&chr(10)%>
<%		else	%>
		<td style="mso-number-format:'\@'" valign=top>
			<%=thisfield%>
		</td><%=chr(13)&chr(10)%>
<%	end if 		%>
<%  NEXT	%>
	</tr><%=chr(13)&chr(10)%>
<%NEXT%>
</table>

</body>
</html>
