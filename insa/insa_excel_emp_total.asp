<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim sch_tab(20,10)
Dim car_tab(20,10)
Dim qul_tab(20,10)
'dim fam_tab(20,10)
'dim edu_tab(20,10)
'dim lan_tab(20,10)

Dim view_condi, curr_date, title_line, savefilename, rsReport
Dim emp_birthday, emp_email, jj

view_condi = Request("view_condi")

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

'Sql = "SELECT * FROM emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY emp_in_date,emp_no,emp_name ASC"
objBuilder.Append "SELECT emtt.emp_no, emtt.emp_birthday, emtt.emp_name, emtt.emp_email, emtt.emp_person1, emtt.emp_person2, "
objBuilder.Append "	emtt.emp_grade, emtt.emp_job, emtt.emp_position, emtt.emp_company, emtt.emp_bonbu, "
objBuilder.Append "	emtt.emp_saupbu, emtt.emp_team, emtt.emp_org_name, emtt.emp_reside_place, "
objBuilder.Append "	emtt.emp_reside_company, emtt.emp_first_date, emtt.emp_in_date, emtt.emp_last_edu, "
objBuilder.Append "	emtt.emp_family_sido, emtt.emp_family_gugun, emtt.emp_family_dong, emtt.emp_family_addr, "
objBuilder.Append "	emtt.emp_hp_ddd, emtt.emp_hp_no1, emtt.emp_hp_no2, emtt.emp_military_id, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
objBuilder.Append "	eomt.org_reside_place, org_reside_company "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "
objBuilder.Append "	AND emtt.emp_no < '900000' "
objBuilder.Append "ORDER BY emtt.emp_in_date, emtt.emp_no, emtt.emp_name ASC"

Set rsReport = Server.CreateObject("ADODB.RecordSet")
rsReport.Open objBuilder.ToString(), Dbconn, 1
objBuilder.Clear()
%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">-->
<!DOCTYPE HTML>
<html lang="ko">
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
						Dim i, j, k
						Dim rs_sch, k_sch, rs_car, k_car
						Dim rs_qul, k_qul

						Set rs_sch = Server.CreateObject("ADODB.RecordSet")
						Set rs_car = Server.CreateObject("ADODB.RecordSet")
						Set rs_qul = Server.CreateObject("ADODB.RecordSet")

						do until rsReport.eof
						   emp_no = rsReport("emp_no")

							'학력사항 db
							for i = 0 to 20
								for j = 0 to 10
									sch_tab(i,j) = ""
								next
							next

							k = 0

							'Sql="select * from emp_school where sch_empno = '"&emp_no&"' order by sch_empno, sch_seq asc"
							objBuilder.Append "SELECT sch_start_date, sch_end_date, sch_school_name, sch_dept, "
							objBuilder.Append "	sch_major, sch_sub_major, sch_degree, sch_finish "
							objBuilder.Append "FROM emp_school WHERE sch_empno = '"&emp_no&"' ORDER BY sch_empno, sch_seq ASC "

							rs_sch.Open objBuilder.ToString(), Dbconn, 1
							objBuilder.Clear()

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
							'Sql="select * from emp_career where career_empno = '"&emp_no&"' order by career_empno, career_seq asc"
							objBuilder.Append "SELECT career_join_date, career_end_date, career_office, "
							objBuilder.Append "	career_dept, career_position, career_task "
							objBuilder.Append "FROM emp_career WHERE career_empno = '"&emp_no&"' ORDER BY career_empno, career_seq ASC "

							rs_car.Open objBuilder.ToString(), Dbconn, 1
							objBuilder.Clear()

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
							'Sql="select * from emp_qual where qual_empno = '"&emp_no&"' order by qual_empno, qual_seq asc"
							objBuilder.Append "SELECT qual_type, qual_grade, qual_pass_date, qual_org, qual_no "
							objBuilder.Append "FROM emp_qual WHERE qual_empno = '"&emp_no&"' ORDER BY qual_empno, qual_seq ASC "

							rs_qul.Open objBuilder.ToString(), Dbconn, 1
							objBuilder.Clear()

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

							if rsReport("emp_birthday") = "1900-01-01" then
								   emp_birthday = ""
							   else
								   emp_birthday = rsReport("emp_birthday")
							end if

							emp_email = rsReport("emp_email") + "@k-won.co.kr"

						   for jj = 1 to 20

							   if jj = 1 then
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_no")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=emp_birthday%></td>

									<td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_person1")%>-<%=rsReport("emp_person2")%></td>

                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_grade")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_job")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_position")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("org_company")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("org_bonbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("org_saupbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("org_team")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("org_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("org_reside_place")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("org_reside_company")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_first_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_in_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_last_edu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_family_sido")%>&nbsp;<%=rsReport("emp_family_gugun")%>&nbsp;<%=rsReport("emp_family_dong")%>&nbsp;<%=rsReport("emp_family_addr")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_hp_ddd")%>-<%=rsReport("emp_hp_no1")%>-<%=rsReport("emp_hp_no2")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=emp_email%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rsReport("emp_military_id")%></td>

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

						   rsReport.movenext()
						loop
						rsReport.close() : Set rsReport = Nothing
						DBConn.Close() : Set DBConn = Nothing
		    %>
						</tbody>
					</table>
				</div>
		</div>
	</div>
	</body>
</html>
