<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_school_list.asp"

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
end if

if view_condi = "" then
	view_condi = "전체"
	condi_sql = " "
	condi = ""
end if

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_qual = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "상주처회사" then

           Sql= "select count(*) " & _
	               "    from emp_school " &_ 
				   "    INNER JOIN emp_master " & _
	               "    ON emp_school.sch_empno = emp_master.emp_no WHERE (isNull(emp_master.emp_end_date) or emp_master.emp_end_date = '1900-01-01') and (emp_master.emp_reside_company like '%" + condi + "%')"
		   		   
           Set RsCount = Dbconn.Execute (sql)
		   tottal_record = cint(RsCount(0))
           IF tottal_record mod pgsize = 0 THEN
	                 total_page = int(tottal_record / pgsize) 'Result.PageCount
                 ELSE
	                 total_page = int((tottal_record / pgsize) + 1)
           END IF

           Sql= "select * " & _
	               "    from emp_school a, emp_master b " & _
	               "    where a.sch_empno = b.emp_no AND (isNull(b.emp_end_date) or b.emp_end_date = '1900-01-01') and (b.emp_reside_company like '%" + condi + "%') " & _
				   "    ORDER BY sch_empno ASC limit "& stpage & "," &pgsize  
		   Rs.Open Sql, Dbconn, 1
      else 
           if view_condi = "성명" then
				  Sql= "select count(*) " & _
	               "    from emp_school " &_ 
				   "    INNER JOIN emp_master " & _
	               "    ON emp_school.sch_empno = emp_master.emp_no WHERE (isNull(emp_master.emp_end_date) or emp_master.emp_end_date = '1900-01-01') and (emp_master.emp_name like '%" + condi + "%')"
		   		   
                  Set RsCount = Dbconn.Execute (sql)
		          tottal_record = cint(RsCount(0))
                  IF tottal_record mod pgsize = 0 THEN
	                        total_page = int(tottal_record / pgsize) 'Result.PageCount
                        ELSE
	                        total_page = int((tottal_record / pgsize) + 1)
                  END IF

                  Sql= "select * " & _
	                      "    from emp_school a, emp_master b " & _
	                      "    where a.sch_empno = b.emp_no AND  (isNull(b.emp_end_date) or b.emp_end_date = '1900-01-01') and (b.emp_name like '%" + condi + "%') " & _
				          "    ORDER BY sch_empno ASC limit "& stpage & "," &pgsize  
		          Rs.Open Sql, Dbconn, 1
		      else
		         if view_condi = "전체" then
    	                  condi_sql = ""
                    else
                          condi_sql = " and emp_school."+view_condi+" like '%" + condi + "%'"
                 end if		
		
		         Sql= "select count(*) " & _
	               "    from emp_school " &_ 
				   "    INNER JOIN emp_master " & _
	               "    ON emp_school.sch_empno = emp_master.emp_no WHERE (isNull(emp_master.emp_end_date) or emp_master.emp_end_date = '1900-01-01')" + condi_sql
		
'	             Sql = "SELECT count(*) FROM emp_school "+condi_sql+""
                 Set RsCount = Dbconn.Execute (sql)

                 tottal_record = cint(RsCount(0)) 'Result.RecordCount

                 IF tottal_record mod pgsize = 0 THEN
	                    total_page = int(tottal_record / pgsize) 'Result.PageCount
                    ELSE
	                    total_page = int((tottal_record / pgsize) + 1)
                 END IF
                 
				 Sql= "select * " & _
	               "    from emp_school " &_ 
				   "    INNER JOIN emp_master " & _
	               "    ON emp_school.sch_empno = emp_master.emp_no WHERE (isNull(emp_master.emp_end_date) or emp_master.emp_end_date = '1900-01-01')" +condi_sql+" ORDER BY sch_empno ASC limit "& stpage & "," &pgsize 
				 
'                 Sql = "SELECT * FROM emp_school "+condi_sql+" ORDER BY sch_empno ASC limit "& stpage & "," &pgsize 
                 Rs.Open Sql, Dbconn, 1
           end if			 
end if

title_line = " 직원 학력 현황 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "5 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_school_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="성명" <%If view_condi = "성명" then %>selected<% end if %>>성명</option>
                                  <option value="sch_dept" <%If view_condi = "sch_dept" then %>selected<% end if %>>학과</option>
                                  <option value="sch_major" <%If view_condi = "sch_major" then %>selected<% end if %>>전공</option>
                                  <option value="sch_school_name" <%If view_condi = "sch_school_name" then %>selected<% end if %>>학교</option>
                                  <option value="상주처회사" <%If view_condi = "상주처회사" then %>selected<% end if %>>상주처회사</option>
                                </select>
								<strong>조건 : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left" >
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
                <form name="frm_del" method="post" action="org_del_ok.asp?page=<%=page%>&ck_sw=<%="n"%>&view_condi=<%=view_condi%>&condi=<%=condi%>">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="6%" >
							<col width="7%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="*" >
                            <col width="14%" >
                            <col width="12%" >
                            <col width="12%" >
                            <col width="8%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
                                <th scope="col">성명</th>
                                <th scope="col">직위</th>
								<th scope="col">회사</th>
								<th scope="col">소속</th>
                                <th scope="col">학교명</th>
								<th scope="col">기간</th>
								<th scope="col">학과</th>
								<th scope="col">전공</th>
								<th scope="col">부전공</th>
                                <th scope="col">학위</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

                         sch_empno = rs("sch_empno")
                         if sch_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&sch_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_name = Rs_emp("emp_name")
							  emp_grade = Rs_emp("emp_grade")
							  emp_job = Rs_emp("emp_job")
		                      emp_position = Rs_emp("emp_position")
							  emp_org_code = Rs_emp("emp_org_code")
							  emp_org_name = Rs_emp("emp_org_name")
							  emp_company = Rs_emp("emp_company")
		                   end if
	                       Rs_emp.Close()
	                	  end if	
						  
	           			%>
							<tr>
								<td><%=rs("sch_empno")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("sch_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=emp_name%></a>
								</td>
                                <td><%=emp_job%>&nbsp;</td>
                                <td><%=emp_company%>&nbsp;</td>
                                <td><%=emp_org_name%>&nbsp;</td>
                                <td class="first" style=" border-left:1px solid #e3e3e3;"><%=rs("sch_school_name")%>&nbsp;</td>
                                <td><%=rs("sch_start_date")%>∼<%=rs("sch_end_date")%>&nbsp;</td>
                                <td><%=rs("sch_dept")%>&nbsp;</td>
                                <td><%=rs("sch_major")%>&nbsp;</td>
                                <td><%=rs("sch_sub_major")%>&nbsp;</td>
                                <td><%=rs("sch_degree")%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1
                
                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_schoollist.asp?view_condi=<%=view_condi%>&condi=<%=condi%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_school_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_school_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_school_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_school_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_school_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

