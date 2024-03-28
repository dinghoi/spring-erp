<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_car_insurance_list.asp"

from_date=Request.form("from_date")
to_date=Request.form("to_date")

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	owner_view=Request.form("owner_view")
	view_condi = request.form("view_condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
  else
	owner_view=Request("owner_view")
	view_condi = request("view_condi")
	from_date=request("from_date")
    to_date=request("to_date")
end if

if view_condi = "" then
	view_condi = "전체"
	owner_view = "T"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
end if

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_car = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if owner_view = "C" then
	owner_sql = " ins_last_date >= '"+from_date+"' and ins_last_date <= '"+to_date+"' "
  else
	owner_sql = " ins_date >= '"+from_date+"' and ins_date <= '"+to_date+"' "
end if

if view_condi = "전체" then
   Sql = "select count(*) from car_insurance where " + owner_sql
   else  
   Sql = "select count(*) from car_insurance where ins_car_no='"+view_condi+"' and " + owner_sql
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

order_sql = " ORDER BY ins_car_no,ins_date DESC"

if view_condi = "전체" then
   Sql = "select * from car_insurance where " + owner_sql + order_sql + " limit "& stpage & "," &pgsize 
   else  
   Sql = "select * from car_insurance where ins_car_no = '"+view_condi+"' and " + owner_sql + order_sql + " limit "& stpage & "," &pgsize 
end if
Rs.Open Sql, Dbconn, 1

if owner_view = "C" then
	title_line = ""+ view_condi +" - 차량 보험만료 예상현황 "
  else
	title_line = ""+ view_condi +" - 차량 보험가입현황 "
end if

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
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_insurance_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>차량번호 : </strong>
                              <%
								Sql="select * from car_info where (end_date = '1900-01-01' or isNull(end_date)) ORDER BY car_no ASC"
	                            rs_car.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                			  <% 
								do until rs_car.eof 
			  				  %>
                					<option value='<%=rs_car("car_no")%>' <%If view_condi = rs_car("car_no") then %>selected<% end if %>><%=rs_car("car_no")%></option>
                			  <%
									rs_car.movenext()  
								loop 
								rs_car.Close()
							  %>
            					</select>
                                </label>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">가입일
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">만기일
                                </label>
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
                            <col width="6%" >
							<col width="10%" >
                            <col width="6%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="7%" >
                            <col width="6%" >
                            <col width="4%" >
                            <col width="*" >
                            <col width="3%" >
                		</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">차량번호</th>
                                <th scope="col">가입일</th>
                                <th scope="col">보험사</th>
                                <th scope="col">보험기간</th>
                                <th scope="col">보험료</th>
                                <th scope="col">대인1</th>
                                <th scope="col">대인2</th>
                                <th scope="col">대물</th>
                                <th scope="col">자기보험</th>
                                <th scope="col">무상해</th>
                                <th scope="col">자차</th>
                                <th scope="col">연령</th>
                                <th scope="col">긴급<br>출동</th>
                                <th scope="col">계약내용</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

                              car_no = rs("ins_car_no")
							  
							  Sql = "SELECT * FROM car_info where car_no = '"&car_no&"'"
                              Set rs_car = DbConn.Execute(SQL)
							  if not rs_car.eof then
									car_name = rs_car("car_name")
									car_year = rs_car("car_year")
									car_reg_date = rs_car("car_reg_date")
	                             else
								    car_name = ""
									car_year = ""
									car_reg_date = ""
                              end if
                              rs_car.close()
	           			%>
							<tr>
								<td><%=rs("ins_car_no")%>&nbsp;</td>
                                <td><%=rs("ins_date")%>&nbsp;</td>
								<td><%=rs("ins_company")%>&nbsp;</td>
                                <td><%=rs("ins_last_date")%>&nbsp;</td>
                                <td><%=formatnumber(rs("ins_amount"),0)%>&nbsp;</td>
                                <td><%=rs("ins_man1")%>&nbsp;</td>
                                <td><%=rs("ins_man2")%>&nbsp;</td>
                                <td><%=rs("ins_object")%>&nbsp;</td>
                                <td><%=rs("ins_self")%>&nbsp;</td>
                                <td><%=rs("ins_injury")%>&nbsp;</td>
                                <td><%=rs("ins_self_car")%>&nbsp;</td>
                                <td><%=rs("ins_age")%>&nbsp;</td>
                                <td><%=rs("ins_scramble")%>&nbsp;</td>
                         <% if rs("ins_contract_yn") = "Y" then %>
                                <td>계약내용포함&nbsp;</td>
                         <%    else %>
                                <td>계약내용미포함(<%=rs("ins_comment")%>)&nbsp;</td>
                         <% end if %>
                                <td>
                                 <a href="#" onClick="pop_Window('insa_car_insurance_add.asp?car_no=<%=rs("ins_car_no")%>&ins_date=<%=rs("ins_date")%>&car_name=<%=car_name%>&car_year=<%=car_year%>&car_reg_date=<%=car_reg_date%>&u_type=<%="U"%>','car_insurance_add_popup','scrollbars=yes,width=750,height=410')">수정</a>
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
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_car_insurance.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&owner_view=<%=owner_view%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_car_insurance_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_car_insurance_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_car_insurance_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_car_insurance_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_car_insurance_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
            <% if view_condi <> "전체" then    %>           
                    <a href="#" onClick="pop_Window('insa_car_insurance_add.asp?car_no=<%=view_condi%>&u_type=<%=""%>','car_insurance_add_popup','scrollbars=yes,width=750,height=410')" class="btnType04">차량보험료 등록</a>
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

