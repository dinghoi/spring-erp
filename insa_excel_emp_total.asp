<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim sch_tab(20,10)
dim car_tab(20,10)
dim qul_tab(20,10)
dim fam_tab(20,10)
dim edu_tab(20,10)
dim lan_tab(20,10)
	 
view_condi=Request("view_condi")

curr_date = datevalue(mid(cstr(now()),1,10))

if view_condi = "" then
	view_condi = "전체"
end if

title_line = "직원현황(" + view_condi + ")" + cstr(curr_date)

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_sch = Server.CreateObject("ADODB.Recordset")
Set rs_car = Server.CreateObject("ADODB.Recordset")
Set rs_qul = Server.CreateObject("ADODB.Recordset")
Set RsschCnt = Server.CreateObject("ADODB.Recordset")
Set RscarCnt = Server.CreateObject("ADODB.Recordset")
Set RsqulCnt = Server.CreateObject("ADODB.Recordset")

Set Rs_fam = Server.CreateObject("ADODB.Recordset")
Set rs_app = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
Set rs_lan = Server.CreateObject("ADODB.Recordset")
Set rs_stay = Server.CreateObject("ADODB.Recordset")
Set RsfamCnt = Server.CreateObject("ADODB.Recordset")
Set RsappCnt = Server.CreateObject("ADODB.Recordset")
Set RseduCnt = Server.CreateObject("ADODB.Recordset")
Set RslanCnt = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'if view_condi = "전체" then
       Sql = "SELECT * FROM emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY emp_in_date,emp_no,emp_name ASC" 
'   else	   
'	   Sql = "SELECT * FROM emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_company = '"&view_condi&"') and (emp_no < '900000') ORDER BY emp_in_date,emp_no,emp_name ASC" 
'end if

Rs.Open Sql, Dbconn, 1
	

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th colspan="20" scope="col">기본정보</th>
                                <th colspan="7" scope="col">학력사항</th>
                                <th colspan="5" scope="col">경력사항</th>
                                <th colspan="5" scope="col">자격증 현황</th>
							</tr>
                            <tr>
								<th class="first" scope="col">사번</th>
                                <th scope="col">성명</th>
                                <th scope="col">생년월일</th>
								<th scope="col">주민번호</th>
                                <th scope="col">직급</th>
                                <th scope="col">직위</th>
                                <th scope="col">직책</th>
                                <th scope="col">회사</th>
                                <th scope="col">본부</th>
                                <th scope="col">사업부</th>
                                <th scope="col">팀</th>
                                <th scope="col">소속</th>
                                <th scope="col">상주처</th>
                                <th scope="col">상주처회사</th>
                                <th scope="col">최초입사일</th>
                                <th scope="col">입사일</th>
                                <th scope="col">최종학력</th>
                                <th scope="col">현주소</th>
                                <th scope="col">핸드폰</th>
                                <th scope="col">e메일</th>
                                <th scope="col">병역사항</th>
                                
                                <th scope="col">기간</th>
                                <th scope="col">학교명</th>
                                <th scope="col">학과</th>
                                <th scope="col">전공</th>
                                <th scope="col">부전공</th>
                                <th scope="col">학위</th>
                                <th scope="col">졸업</th>
                                
                                <th scope="col">재직기간</th>
                                <th scope="col">회사명</th>
                                <th scope="col">부서</th>
                                <th scope="col">직위</th>
                                <th scope="col">담당업무</th>
                                
                                <th scope="col">자격종목</th>
                                <th scope="col">등급</th>
                                <th scope="col">합격일자</th>
                                <th scope="col">발급기관</th>
                                <th scope="col">자격증번호</th>

							</tr>
						</thead>
						<tbody>
			<%
						do until rs.eof
						
						   emp_no = rs("emp_no")

'학력사항 db
for i = 0 to 20
	for j = 0 to 10
		sch_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_school where sch_empno = '"&emp_no&"' order by sch_empno, sch_seq asc"
	Rs_sch.Open Sql, Dbconn, 1	
	while not rs_sch.eof
		k = k + 1
		sch_tab(k,1) = rs_sch("sch_start_date")
		sch_tab(k,2) = rs_sch("sch_end_date")
		sch_tab(k,3) = rs_sch("sch_school_name")
		sch_tab(k,4) = rs_sch("sch_dept")
		sch_tab(k,5) = rs_sch("sch_major")
		sch_tab(k,6) = rs_sch("sch_sub_major")
		sch_tab(k,7) = rs_sch("sch_degree")
		sch_tab(k,8) = rs_sch("sch_finish")
		rs_sch.movenext()
	Wend
    rs_sch.close()		
	k_sch = k				


'경력사항 db
for i = 0 to 20
	for j = 0 to 10
		car_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_career where career_empno = '"&emp_no&"' order by career_empno, career_seq asc"
	Rs_car.Open Sql, Dbconn, 1	
	while not rs_car.eof
		k = k + 1
		car_tab(k,1) = rs_car("career_join_date")
		car_tab(k,2) = rs_car("career_end_date")
		car_tab(k,3) = rs_car("career_office")
		car_tab(k,4) = rs_car("career_dept")
		car_tab(k,5) = rs_car("career_position")
		car_tab(k,6) = rs_car("career_task")
		rs_car.movenext()
	Wend
    rs_car.close()	
    k_car = k		

'자격사항 db
for i = 0 to 20
	for j = 0 to 10
		qul_tab(i,j) = ""
	next
next

	k = 0
    Sql="select * from emp_qual where qual_empno = '"&emp_no&"' order by qual_empno, qual_seq asc"
	rs_qul.Open Sql, Dbconn, 1	
	while not rs_qul.eof
		k = k + 1
		qul_tab(k,1) = rs_qul("qual_type")
		qul_tab(k,2) = rs_qul("qual_grade")
		qul_tab(k,3) = rs_qul("qual_pass_date")
		qul_tab(k,4) = rs_qul("qual_org")
		qul_tab(k,5) = rs_qul("qual_no")
		rs_qul.movenext()
	Wend
    rs_qul.close()	
	k_qul = k	
	
	if rs("emp_birthday") = "1900-01-01" then
		   emp_birthday = ""
	   else 
		   emp_birthday = rs("emp_birthday")
	end if
	
	emp_email = rs("emp_email") + "@k-won.co.kr"					   
						    

						   for jj = 1 to 20

							   if jj = 1 then
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_no")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=emp_birthday%></td>

									<td class="left" bgcolor="#EEFFFF"><%=rs("emp_person1")%>-<%=rs("emp_person2")%></td>

                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_grade")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_job")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_position")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_company")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_bonbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_saupbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_team")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_org_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_reside_place")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_reside_company")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_first_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_in_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_last_edu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_family_sido")%>&nbsp;<%=rs("emp_family_gugun")%>&nbsp;<%=rs("emp_family_dong")%>&nbsp;<%=rs("emp_family_addr")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=emp_email%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("emp_military_id")%></td>
                                    
								    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,1)%>&nbsp;~&nbsp;<%=sch_tab(jj,2)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,3)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,4)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,5)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,6)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,7)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=sch_tab(jj,8)%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,1)%>&nbsp;~&nbsp;<%=car_tab(jj,2)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,3)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,4)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,5)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=car_tab(jj,6)%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,1)%>&nbsp;</td>
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,2)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,3)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,4)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=qul_tab(jj,5)%></td>

						         </tr>
            <%
			                    else
								   if sch_tab(jj,1) <> "" or car_tab(jj,1) <> "" or qul_tab(jj,1) <> "" then
		    %>		
                                 <tr>
								    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    
								    <td class="left" ><%=sch_tab(jj,1)%>&nbsp;~&nbsp;<%=sch_tab(jj,2)%></td>
                                    <td class="left" ><%=sch_tab(jj,3)%></td>
                                    <td class="left" ><%=sch_tab(jj,4)%></td>
                                    <td class="left" ><%=sch_tab(jj,5)%></td>
                                    <td class="left" ><%=sch_tab(jj,6)%></td>
                                    <td class="left" ><%=sch_tab(jj,7)%></td>
                                    <td class="left" ><%=sch_tab(jj,8)%></td>
                                    
                                    <td class="left" ><%=car_tab(jj,1)%>&nbsp;~&nbsp;<%=car_tab(jj,2)%></td>
                                    <td class="left" ><%=car_tab(jj,3)%></td>
                                    <td class="left" ><%=car_tab(jj,4)%></td>
                                    <td class="left" ><%=car_tab(jj,5)%></td>
                                    <td class="left" ><%=car_tab(jj,6)%></td>
                                    
                                    <td class="left" ><%=qul_tab(jj,1)%>&nbsp;</td>
                                    <td class="left" ><%=qul_tab(jj,2)%></td>
                                    <td class="left" ><%=qul_tab(jj,3)%></td>
                                    <td class="left" ><%=qul_tab(jj,4)%></td>
                                    <td class="left" ><%=qul_tab(jj,5)%></td>
						         </tr>            
            <%            							
							       end if
							 end if
	                       next
							  
						   rs.movenext()
						loop
						rs.close()
		    %>						
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>
