<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_pay_expense_mg.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

in_empno=Request.form("in_empno")
in_name=Request.form("in_name")

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "y" then
	view_condi = request("view_condi")
	ex_deduct_id = request("ex_deduct_id") 
	from_date=request("from_date")
    to_date=request("to_date") 
  else
	view_condi = request.form("view_condi")
	ex_deduct_id = Request.Form("ex_deduct_id") 
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
end if

if view_condi = "" then
	view_condi = "(주)케이원정보통신"
	ex_deduct_id = "G"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
end if

rever_yyyymm = mid(cstr(from_date),1,7) '귀속년월
ex_pay_date = to_date '지급일

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'Sql = "SELECT * FROM k1_memb where "+condi_sql+"mg_group = '"+mg_group+"' ORDER BY user_name ASC"
'where_sql = " WHERE isNull(emp_end_date) or emp_end_date = '1900-01-01'"

Sql = "select count(*) from pay_expense where (rever_yymm = '"+rever_yyyymm+"') and (ex_company = '"+view_condi+"') and (ex_deduct_id = '"+ex_deduct_id+"')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "select * from pay_expense where (rever_yymm = '"+rever_yyyymm+"') and (ex_company = '"+view_condi+"') and (ex_deduct_id = '"+ex_deduct_id+"') ORDER BY rever_yymm,ex_date,ex_code,ex_emp_no ASC limit "& stpage & "," &pgsize 

Rs.Open Sql, Dbconn, 1

title_line = "  지급/공제자료 입력 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
		    $(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  

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
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_expense_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
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
								<strong>귀속년월 : </strong>
                                <input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>지급일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<label>
                                <strong>구분 : </strong>
        						    <select name="ex_deduct_id" id="ex_deduct_id" style="width:100px">
                                      <option value="" <% if ex_deduct_id = "" then %>selected<% end if %>>선택</option>
                                      <option value="G" <% if ex_deduct_id = "G" then %>selected<% end if %>>지급항목</option>
                                      <option value="D" <% if ex_deduct_id = "D" then %>selected<% end if %>>공제항목</option>
                                    </select>
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
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="6%" >
							<col width="*" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">귀속년월</th>
								<th scope="col">발생일자</th>
								<th scope="col">항목</th>
								<th scope="col">성명</th>
								<th scope="col">소속</th>
                                <th scope="col">금액</th>
                                <th scope="col">과세여부</th>
								<th scope="col">근무일수</th>
                                <th scope="col">비고</th>
                                <th scope="col">변경</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  emp_no = rs("emp_no")
							  ex_tax_name = "해당없음"
	           			%>
							<tr>
								<td class="first"><%=rs("rever_yymm")%>&nbsp;<td>
                                <td><%=rs("ex_date")%>&nbsp;</td>
                                <td><%=rs("ex_code_name")%>&nbsp;</td>
                                <td><%=rs("ex_emp_name")%>(<%=rs("ex_emp_no")%>)&nbsp;</td>
                                <td><%=rs("ex_org_name")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(ex_amount,0)%>&nbsp;</td>
                                <% if rs("ex_tax_id") =  "1" then ex_tax_name = "과세" end if %>
                                <% if rs("ex_tax_id") =  "2" then ex_tax_name = "비과세" end if %>
                                <% if rs("ex_tax_id") =  "3" then ex_tax_name = "감면세액" end if %>
                                <td><%=ex_tax_name%>&nbsp;</td>
                                <td><%=rs("ex_work_cnt")%>&nbsp;</td>
                                <td><%=rs("ex_comment")%>&nbsp;</td>
                                <td><a href="#" onClick="pop_Window('insa_pay_expense_add.asp?ex_emp_no=<%=rs("ex_emp_no")%>&ex_emp_name=<%=rs("ex_emp_name")%>&ex_deduct_id=<%=rs("ex_deduct_id")%>&ex_date=<%=rs("ex_date")%>&ex_code=<%=rs("ex_code")%>&u_type=<%="U"%>','insa_pay_give_add_pop','scrollbars=yes,width=750,height=400')">수정</a></td>
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
                    <a href="insa_excel_pay_expense.asp?view_condi=<%=view_condi%>&in_empno=<%=in_empno%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_expense_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&in_empno=<%=in_empno%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_expense_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&in_empno=<%=in_empno%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_expense_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&in_empno=<%=in_empno%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_expense_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&in_empno=<%=in_empno%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_expense_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&in_empno=<%=in_empno%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_month_give_add.asp?emp_no=<%=in_empno%>&emp_name=<%=in_name%>&rever_yyyymm=<%=rever_yyyymm%>&give_date=<%=give_date%>&view_condi=<%=view_condi%>&u_type=<%=""%>','insa_pay_give_add_pop','scrollbars=yes,width=750,height=700')" class="btnType04">야특근 Upload</a>
					<a href="#" onClick="pop_Window('insa_pay_month_deduct_add.asp?emp_no=<%=in_empno%>&emp_name=<%=in_name%>&rever_yyyymm=<%=rever_yyyymm%>&give_date=<%=give_date%>&u_type=<%=""%>','insa_pay_deduct_add_pop','scrollbars=yes,width=750,height=400')" class="btnType04">기타공제Upload</a>
                    <a href="#" onClick="pop_Window('insa_pay_expense_add.asp?ex_deduct_id=<%=ex_deduct_id%>&rever_yyyymm=<%=rever_yyyymm%>&ex_pay_date=<%=ex_pay_date%>&view_condi=<%=view_condi%>&u_type=<%=""%>','insa_pay_give_add_pop','scrollbars=yes,width=750,height=400')" class="btnType04">지급/공제등록</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

