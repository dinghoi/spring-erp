<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim field_check
Dim field_view

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	owner_view=Request("owner_view")
	field_check=Request("field_check")
	field_view=Request("field_view")
  else
	owner_view=Request.form("owner_view")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if

If owner_view = "" Then
	owner_view = "T"
	field_check = "total"
End If

If field_check = "total" Then
	field_view = ""
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
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'base_sql = "select * FROM car_info INNER JOIN memb ON car_info.owner_emp_no = memb.emp_no "
base_sql = "select * FROM car_info "

if owner_view = "C" then
	owner_sql = " where car_owner = '회사' "
  elseif owner_view = "P" then
	owner_sql = " where car_owner = '개인' "
  else  
  	owner_sql = " where (car_owner = '개인' or car_owner = '회사') "
end if

if field_check <> "total" then
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if

sql = "select count(*) FROM car_info " + owner_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

order_sql = " ORDER BY car_owner DESC, car_no DESC"

sql = base_sql + owner_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'Response.write Sql
title_line = "차량 관리"
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
				return "7 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.field_check.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건검색</dt>
                        <dd>
                            <p>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">총괄
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">회사
                                <input name="owner_view" type="radio" value="P" <% if owner_view = "P" then %>checked<% end if %> style="width:25px">개인
                                </label>
                                <label>
								<strong>필드조건</strong>
                                <select name="field_check" id="field_check" style="width:100px">
                                  <option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                  <option value="buy_gubun" <% if field_check = "buy_gubun" then %>selected<% end if %>>구매구분</option>
                                  <option value="owner_emp_name" <% if field_check = "owner_emp_name" then %>selected<% end if %>>운행자</option>
                                  <option value="oil_kind" <% if field_check = "oil_kind" then %>selected<% end if %>>유종</option>
                                  <option value="car_no" <% if field_check = "car_no" then %>selected<% end if %>>차량번호</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:100px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="*" >
							<col width="5%" >
							<col width="4%" >
							<col width="6%" >
							<col width="6%" >
              <col width="6%" >
              <col width="6%" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="4%" >
              <col width="4%" >
              <col width="5%" >
							<col width="5%" >
              <col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">차량번호</th>
								<th scope="col">차종/연식</th>
								<th scope="col">유종</th>
								<th scope="col">소유</th>
								<th scope="col">구매<br>구분</th>
								<th scope="col">차량등록일</th>
                <th scope="col">보험료</th>
                <th scope="col">보험기간</th>
								<th scope="col">운행자</th>
								<th scope="col">최종KM</th>
								<th scope="col">최종검사일</th>
								<th scope="col">변경</th>
								<th scope="col">보험</th>
								<th scope="col">운행자</th>
								<th scope="col">AS등록</th>
								<th scope="col">과태료</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
						     owner_emp_name = ""
							 owner_emp_no = rs("owner_emp_no")
						     if rs("owner_emp_name") = "" or isnull(rs("owner_emp_name")) then
							     Sql="select * from emp_master where emp_no = '"&owner_emp_no&"'"
	                             Set rs_emp=DbConn.Execute(Sql)
								 owner_emp_name = rs_emp("emp_name")
							   else 
							     owner_emp_name = rs("owner_emp_name")
							 end if
							if rs("last_check_date") = "1900-01-01"  then
	                               last_check_date = ""
							   else 
							       last_check_date = rs("last_check_date")
	                        end if
	                        if rs("end_date") = "1900-01-01" then
	                               end_date = ""
							   else 
							       end_date = rs("end_date")
	                        end if
							if rs("car_year") = "1900-01-01" then
	                               car_year = ""
							   else 
							       car_year = rs("car_year")
	                        end if
						%>
							<tr>
								<td class="first">
                                <a href="#" onClick="pop_Window('insa_car_info_view.asp?car_no=<%=rs("car_no")%>&car_name=<%=rs("car_name")%>&car_year=<%=rs("car_year")%>&car_reg_date=<%=rs("car_reg_date")%>&oil_kind=<%=rs("oil_kind")%>','insa_car_info_pop','scrollbars=yes,width=800,height=600')"><%=rs("car_no")%></a>&nbsp;
                                </td>
								<td class="left"><%=rs("car_name")%>(<%=car_year%>)</td>
								<td><%=rs("oil_kind")%></td>
								<td><%=rs("car_owner")%></td>
								<td><%=rs("buy_gubun")%>&nbsp;<%=rs("rental_company")%></td>
								<td><%=rs("car_reg_date")%>&nbsp;</td>
                                <td class="right"><a href="#" onClick="pop_Window('insa_car_ins_view.asp?car_no=<%=rs("car_no")%>&car_name=<%=rs("car_name")%>&car_year=<%=rs("car_year")%>&car_reg_date=<%=rs("car_reg_date")%>&u_type=<%="U"%>','insa_car_ins_pop','scrollbars=yes,width=1200,height=600')"><%=formatnumber(rs("insurance_amt"),0)%></a>&nbsp;</td>
                                <td><%=rs("insurance_date")%>&nbsp;</td>
                                <td><a href="#" onClick="pop_Window('insa_car_drvuser_view.asp?car_no=<%=rs("car_no")%>&car_name=<%=rs("car_name")%>&car_year=<%=rs("car_year")%>&car_reg_date=<%=rs("car_reg_date")%>&u_type=<%="U"%>','insa_car_user_pop','scrollbars=yes,width=750,height=600')"><%=owner_emp_name%>(<%=rs("owner_emp_no")%>)</a>&nbsp;</td>
                                <td class="right"><a href="#" onClick="pop_Window('insa_car_drv_view.asp?car_no=<%=rs("car_no")%>&car_name=<%=rs("car_name")%>&car_year=<%=rs("car_year")%>&car_reg_date=<%=rs("car_reg_date")%>&u_type=<%="U"%>','insa_car_drv_pop','scrollbars=yes,width=1250,height=600')"><%=formatnumber(rs("last_km"),0)%></a>&nbsp;</td>
								<td><%=last_check_date%>&nbsp;</td>
								<td>
									<a href="#" onClick="pop_Window('insa_car_info_add.asp?car_no=<%=rs("car_no")%>&u_type=<%="U"%>','car_info_add_popup','scrollbars=yes,width=750,height=450')">변경</a>
								</td>
								<td>
									<a href="#" onClick="pop_Window('insa_car_insurance_add.asp?car_no=<%=rs("car_no")%>&car_name=<%=rs("car_name")%>&car_year=<%=rs("car_year")%>&car_reg_date=<%=rs("car_reg_date")%>&u_type=<%=""%>','car_insurance_add_popup','scrollbars=yes,width=750,height=410')">등록</a>
                </td>
                <td>
                	<a href="#" onClick="pop_Window('insa_car_drvuser_add.asp?car_no=<%=rs("car_no")%>&car_name=<%=rs("car_name")%>&car_year=<%=rs("car_year")%>&car_reg_date=<%=rs("car_reg_date")%>&owner_emp_name=<%=owner_emp_name%>&owner_emp_no=<%=rs("owner_emp_no")%>&u_type=<%=""%>','car_drvuser_add_popup','scrollbars=yes,width=750,height=300')">변경</a>
                </td>
								<td>
                	<a href="#" onClick="pop_Window('insa_car_as_add.asp?car_no=<%=rs("car_no")%>&car_name=<%=rs("car_name")%>&car_year=<%=rs("car_year")%>&car_reg_date=<%=rs("car_reg_date")%>&owner_emp_name=<%=owner_emp_name%>&owner_emp_no=<%=rs("owner_emp_no")%>&car_use_dept=<%=rs("car_use_dept")%>&oil_kind=<%=rs("oil_kind")%>&car_owner=<%=rs("car_owner")%>&u_type=<%=""%>','car_as_add_popup','scrollbars=yes,width=750,height=350')">등록</a>
                </td>
                <td>
                	<a href="#" onClick="pop_Window('insa_car_penalty_add.asp?car_no=<%=rs("car_no")%>&car_name=<%=rs("car_name")%>&car_year=<%=rs("car_year")%>&car_reg_date=<%=rs("car_reg_date")%>&owner_emp_name=<%=owner_emp_name%>&owner_emp_no=<%=rs("owner_emp_no")%>&car_use_dept=<%=rs("car_use_dept")%>&oil_kind=<%=rs("oil_kind")%>&car_owner=<%=rs("car_owner")%>&u_type=<%=""%>','car_as_add_popup','scrollbars=yes,width=750,height=410')">등록</a>
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
                    <a href="insa_excel_car_info.asp?owner_view=<%=owner_view%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_car_mg.asp?page=<%=first_page%>&owner_view=<%=owner_view%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_car_mg.asp?page=<%=intstart -1%>&owner_view=<%=owner_view%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_car_mg.asp?page=<%=i%>&owner_view=<%=owner_view%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_car_mg.asp?page=<%=intend+1%>&owner_view=<%=owner_view%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_car_mg.asp?page=<%=total_page%>&owner_view=<%=owner_view%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('insa_car_info_add.asp','car_info_add_popup','scrollbars=yes,width=750,height=500')" class="btnType04">신규차량등록</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

