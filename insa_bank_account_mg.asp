<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_bank_account_mg.asp"

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")
owner_view=request("owner_view")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	owner_view=Request.form("owner_view")
	condi = request.form("condi")
  else
	view_condi = request("view_condi")
	owner_view=request("owner_view")
	condi = request("condi")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	condi = ""
	owner_view = "C"
	ck_sw = "n"
end if

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_bnk = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if condi = "" then
      Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"')  and (emp_no < '900000')"
   else  
      if owner_view = "C" then 
            Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"') and (emp_name like '%"+condi+"%')"
         else
            Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"') and (emp_no = '"+condi+"')"
	  end if
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if condi = "" then
      Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"')  and (emp_no < '900000') ORDER BY emp_no,emp_name ASC limit "& stpage & "," &pgsize 
   else  
      if owner_view = "C" then 
            Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"') and (emp_name like '%"+condi+"%') ORDER BY emp_no,emp_name ASC limit "& stpage & "," &pgsize 
         else
            Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"+view_condi+"') and (emp_no = '"+condi+"') ORDER BY emp_no,emp_name ASC limit "& stpage & "," &pgsize 
	  end if
end if

Rs.Open Sql, Dbconn, 1

title_line = "직원 은행계좌 현황 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "6 1";
			}
			function goAction () {
			   window.close () ;
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
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_code_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_bank_account_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = '회사' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">사번
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">성명
                                </label>
							<strong>조건 : </strong>
								<label>
        						<input name="condi" type="text" id="condi" value="<%=condi%>" style="width:100px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="9%" >
                            <col width="6%" >
							<col width="12%" >
                            <col width="9%" >
							<col width="*" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
								<th scope="col">입사일</th>
                                <th scope="col">소속</th>
                                <th scope="col">거래은행</th>
								<th scope="col">계좌번호</th>
                                <th scope="col">예금주</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                                <th colspan="2" scope="col">은행계좌</th>
							</tr>
						</thead>
						<tbody>
						<%
						if  view_condi <> "" then 
						do until rs.eof
                              bank_name = ""
							  account_no = ""
							  account_holder = ""
							  emp_no = rs("emp_no")
							  emp_person1 = rs("emp_person1")
							  emp_person2 = rs("emp_person2")
	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                        <%
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
                          %>
                                <td><%=bank_name%>&nbsp;</td>
                                <td><%=account_no%>&nbsp;</td>
                                <td><%=account_holder%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
                                <td><a href="#" onClick="pop_Window('insa_bank_account_add.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&emp_person1=<%=rs("emp_person1")%>&emp_person2=<%=rs("emp_person2")%>&u_type=<%="U"%>','insa_bank_add_pop','scrollbars=yes,width=750,height=300')">수정</a></td>
                                <td><a href="#" onClick="pop_Window('insa_bank_account_add.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&emp_person1=<%=rs("emp_person1")%>&emp_person2=<%=rs("emp_person2")%>&u_type=<%=""%>','insa_bank_add_pop','scrollbars=yes,width=750,height=300')">등록</a></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						end if
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
                    <a href="insa_excel_banklist.asp?view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_bank_account_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_bank_account_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_bank_account_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_bank_account_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_bank_account_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&in_&view_condi=<%=view_condi%>&condi=<%=condi%>&owner_view=<%=owner_view%>&ck_sw=<%="y"%>">[마지막]</a>
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

