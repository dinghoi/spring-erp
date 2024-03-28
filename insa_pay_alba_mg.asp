<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim field_check
Dim field_view

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

view_condi = request("view_condi")
condi = request("condi")
owner_view=request("owner_view")

be_pg = "insa_pay_alba_mg.asp"

curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
	owner_view=request("owner_view")
	condi = request("condi")
  else
	view_condi=Request.form("view_condi")
	owner_view=Request.form("owner_view")
	condi = request.form("condi")
End if

If view_condi = "" Then
	view_condi = "케이원정보통신"
	condi = ""
	owner_view = "C"
	ck_sw = "n"
End If

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_into = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY company,org_name,draft_no ASC"

if condi = "" then
       where_sql = " WHERE (end_yn = '' or isnull(end_yn)) and (company = '"&view_condi&"')"
   else  
      if owner_view = "C" then 	   
	         where_sql = " WHERE (end_yn = '' or isnull(end_yn)) and (company = '"&view_condi&"') and (draft_man like '%"+condi+"%')"
		 else
		     where_sql = " WHERE (end_yn = '' or isnull(end_yn)) and (company = '"&view_condi&"') and (draft_no = '"+condi+"')"
	  end if
end if

Sql = "SELECT count(*) FROM emp_alba_mst " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_alba_mst " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "사업소득자 현황"
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
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_alba_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_alba_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>회사 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">

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
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="10%" >
							
							<col width="10%" >
                            <col width="8%" >
							<col width="*" >
                            <col width="6%" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">등록번호</th>
								<th scope="col">소득자명</th>
								<th scope="col">업무<br>등록일</th>
								<th scope="col">주민번호</th>
								<th scope="col">거주<br>구분</th>
								<th scope="col">소득구분</th>
								<th scope="col">회사</th>
								<th scope="col">부서</th>
								
								<th scope="col">핸드폰</th>
                                <th scope="col">은행</th>
								<th scope="col">계좌번호</th>
                                <th scope="col">예금주</th>
                                <th scope="col">수정</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
						%>
							<tr>
								<td class="first"><%=rs("draft_no")%>&nbsp;</td>
								<td><%=rs("draft_man")%></td>
								<td><%=rs("draft_date")%></td>
								<td><%=rs("person_no1")%>-<%=rs("person_no2")%></td>
								<td><%=rs("draft_live_name")%>&nbsp;</td>
								<td><%=rs("draft_tax_id")%>&nbsp;</td>
								<td><%=rs("company")%></td>
								<td><%=rs("org_name")%>&nbsp;</td>
                                
                                <td><%=rs("hp_ddd")%>-<%=rs("hp_no1")%>-<%=rs("hp_no2")%>&nbsp;</td>
                                <td><%=rs("bank_name")%>&nbsp;</td>
                                <td><%=rs("account_no")%>&nbsp;</td>
                                <td><%=rs("account_name")%>&nbsp;</td>
								<td>
                                <a href="#" onClick="pop_Window('insa_pay_alba_add.asp?draft_no=<%=rs("draft_no")%>&u_type=<%="U"%>','car_info_add_popup','scrollbars=yes,width=750,height=450')">변경</a>
                                </td>
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
				    <td width="20%">
					<div class="btnCenter">
                    <a href="insa_excel_alba_list.asp?view_condi=<%=view_condi%>"class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_pay_alba_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_alba_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_alba_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_pay_alba_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_alba_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('insa_pay_alba_add.asp?view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>','pay_alba_add_popup','scrollbars=yes,width=750,height=450')" class="btnType04">사업소득자등록</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

