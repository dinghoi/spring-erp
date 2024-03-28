<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_qual_list.asp"

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_company = request.form("view_company")
	view_condi = request.form("view_condi")
	condi = request.form("condi")
  else
	view_company=Request("view_company")
	view_condi = request("view_condi")
	condi = request("condi")  
end if

if view_condi = "" then
	view_company = "케이원정보통신"
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
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_qual = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "상주처회사" then

            Sql= "select count(*) " & _
	               "    from emp_qual " &_ 
				   "    INNER JOIN emp_master " & _
	               "    ON emp_qual.qual_empno = emp_master.emp_no WHERE (isNull(emp_master.emp_end_date) or emp_master.emp_end_date = '1900-01-01') and (emp_master.emp_company = '"&view_company&"') and (emp_master.emp_reside_company like '%" + condi + "%')"

           'Sql= "select count(*) " & _
	       '        "    from emp_qual a, emp_master b " & _
	       '        "    where a.qual_empno = b.emp_no AND b.emp_reside_company=like '%" + condi + "%'" 
		   		   
           Set RsCount = Dbconn.Execute (sql)
		   tottal_record = cint(RsCount(0))
           IF tottal_record mod pgsize = 0 THEN
	                 total_page = int(tottal_record / pgsize) 'Result.PageCount
                 ELSE
	                 total_page = int((tottal_record / pgsize) + 1)
           END IF

           Sql= "select * " & _
	               "    from emp_qual a, emp_master b " & _
	               "    where a.qual_empno = b.emp_no AND (isNull(b.emp_end_date) or b.emp_end_date = '1900-01-01') and (b.emp_company = '"&view_company&"') and (b.emp_reside_company like '%" + condi + "%') " & _
				   "    ORDER BY qual_empno ASC limit "& stpage & "," &pgsize  
		   Rs.Open Sql, Dbconn, 1
end if

if view_condi = "자격증명" then
'	condi_sql = " where qual_type like '%" + condi + "%'"
'	Sql = "SELECT count(*) FROM emp_qual "+condi_sql+""
	
	Sql= "select count(*) " & _
	               "    from emp_qual " &_ 
				   "    INNER JOIN emp_master " & _
	               "    ON emp_qual.qual_empno = emp_master.emp_no WHERE (isNull(emp_master.emp_end_date) or emp_master.emp_end_date = '1900-01-01') and (emp_master.emp_company = '"&view_company&"') and (emp_qual.qual_type like '%" + condi + "%')"
	
    Set RsCount = Dbconn.Execute (sql)

    tottal_record = cint(RsCount(0)) 'Result.RecordCount

    IF tottal_record mod pgsize = 0 THEN
	      total_page = int(tottal_record / pgsize) 'Result.PageCount
       ELSE
	      total_page = int((tottal_record / pgsize) + 1)
    END IF

'    Sql = "SELECT * FROM emp_qual "+condi_sql+" ORDER BY qual_empno ASC limit "& stpage & "," &pgsize 
	
	Sql= "select * " & _
	               "    from emp_qual a, emp_master b " & _
	               "    where a.qual_empno = b.emp_no AND (isNull(b.emp_end_date) or b.emp_end_date = '1900-01-01') and (b.emp_company = '"&view_company&"') and (a.qual_type like '%" + condi + "%') " & _
				   "    ORDER BY qual_empno ASC limit "& stpage & "," &pgsize  
	
    Rs.Open Sql, Dbconn, 1
end if

if view_condi = "전체" then
'	condi_sql = ""
'	Sql = "SELECT count(*) FROM emp_qual "+condi_sql+""
	
	Sql= "select count(*) " & _
	               "    from emp_qual " &_ 
				   "    INNER JOIN emp_master " & _
	               "    ON emp_qual.qual_empno = emp_master.emp_no WHERE (isNull(emp_master.emp_end_date) or emp_master.emp_end_date = '1900-01-01') and (emp_master.emp_company = '"&view_company&"')"
	
    Set RsCount = Dbconn.Execute (sql)

    tottal_record = cint(RsCount(0)) 'Result.RecordCount

    IF tottal_record mod pgsize = 0 THEN
	       total_page = int(tottal_record / pgsize) 'Result.PageCount
       ELSE
	       total_page = int((tottal_record / pgsize) + 1)
    END IF

'    Sql = "SELECT * FROM emp_qual "+condi_sql+" ORDER BY qual_empno ASC limit "& stpage & "," &pgsize 
	
	Sql= "select * " & _
	               "    from emp_qual a, emp_master b " & _
	               "    where a.qual_empno = b.emp_no AND (isNull(b.emp_end_date) or b.emp_end_date = '1900-01-01') and (b.emp_company = '"&view_company&"') " & _
				   "    ORDER BY qual_empno ASC limit "& stpage & "," &pgsize  
	
    Rs.Open Sql, Dbconn, 1
end if

title_line = " 자격증 보유 현황 "
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_qual_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_company" id="view_company" type="text" style="width:150px">

                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <strong>조건 : </strong>
                                <label>                                
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="자격증명" <%If view_condi = "자격증명" then %>selected<% end if %>>자격증명</option>
                                  <option value="상주처회사" <%If view_condi = "상주처회사" then %>selected<% end if %>>상주처회사</option>
                                </select>
                                </label>
								<strong>검색명 : </strong>
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
							<col width="14%" >
							<col width="6%" >
							<col width="*" >
							<col width="12%" >
							<col width="8%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="10%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">자격종목</th>
								<th scope="col">등급</th>
								<th scope="col">발급기관</th>
								<th scope="col">자격등록번호</th>
								<th scope="col">취득일</th>
								<th scope="col">사번</th>
                                <th scope="col">성명</th>
                                <th scope="col">직위</th>
								<th scope="col">회사</th>
								<th scope="col">소속</th>
								<th scope="col">상세</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

                         qual_empno = rs("qual_empno")
                         if qual_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&qual_empno&"'"
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
								<td class="first"><%=rs("qual_type")%>&nbsp;</td>
                                <td><%=rs("qual_grade")%>&nbsp;</td>
                                <td><%=rs("qual_org")%>&nbsp;</td>
                                <td><%=rs("qual_no")%>&nbsp;</td>
                                <td><%=rs("qual_pass_date")%>&nbsp;</td>
                                <td><%=rs("qual_empno")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("qual_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=emp_name%></a>
								</td>
                                <td><%=emp_job%>&nbsp;</td>
                                <td><%=emp_company%>&nbsp;</td>
                                <td><%=emp_org_name%>&nbsp;</td>
                                <td><a href="#" onClick="pop_Window('insa_qual_view.asp?emp_no=<%=rs("qual_empno")%>&emp_name=<%=emp_name%>','qualview','scrollbars=yes,width=800,height=400')">조회</a>&nbsp;</td>
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
                    <a href="insa_excel_quallist.asp?view_condi=<%=view_condi%>&condi=<%=condi%>&view_company=<%=view_company%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_qual_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&view_company=<%=view_company%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_qual_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&view_company=<%=view_company%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_qual_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&view_company=<%=view_company%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_qual_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&view_company=<%=view_company%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_qual_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&view_company=<%=view_company%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
                    <% if end_view = "Y" then %>
					<a href="#" onClick="pop_Window('insa_family_add.asp?family_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_family_add_pop','scrollbars=yes,width=750,height=400')" class="btnType04">자격조회</a>
					<a href="payment_slip_end.asp?be_pg=<%=be_pg%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>" class="btnType04">자격증 조회</a>
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

