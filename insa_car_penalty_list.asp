<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_car_penalty_list.asp"

from_date=Request.form("from_date")
to_date=Request.form("to_date")

Page=Request("page")
'view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
  else
	field_check=Request("field_check")
	field_view=Request("field_view")
	from_date=request("from_date")
    to_date=request("to_date")
end if

if field_check = "" then
	field_check = "total"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
end if

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
Set Rs_car = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

owner_sql = " where pe_date >= '"+from_date+"' and pe_date <= '"+to_date+"' "
order_sql = " ORDER BY pe_car_no,pe_date,pe_seq DESC"

if field_check <> "total" then
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if

sql = "select count(*) FROM car_penalty " + owner_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

tot_amount = 0
tot_in_amt = 0

sql = "select * from car_penalty " + owner_sql + field_sql + order_sql
Rs.Open Sql, Dbconn, 1
do until rs.eof
       tot_amount = tot_amount + int(rs("pe_amount"))
	   tot_in_amt = tot_in_amt + int(rs("pe_in_amt"))
	rs.movenext()
loop
rs.close()	

jan_amount = tot_amount - tot_in_amt

sql = "select * from car_penalty " + owner_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = ""+ view_condi +" - 차량 과태료 현황 "
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
			
			function car_penalty_del(val, val2, val3, val4) {

            if (!confirm("정말 삭제하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.pe_car_no.value = val;
			document.frm.pe_date.value = val2;
			document.frm.pe_seq.value = val3;
			document.frm.car_name.value = val4;
		
            document.frm.action = "insa_car_penalty_del.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_penalty_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <label>
								<strong>필드조건</strong>
                                <select name="field_check" id="field_check" style="width:100px">
                                  <option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                  <option value="pe_car_no" <% if field_check = "pe_car_no" then %>selected<% end if %>>차량번호</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:100px; text-align:left" >
								</label>
								<label>
								<strong>발생일(From) : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일(To) : </strong>
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
							<col width="8%" >
							<col width="8%" >
                            <col width="6%" >
							<col width="6%" >
                            <col width="6%" >
							<col width="*" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="8%" >
                            <col width="6%" >
                            <col width="8%" >
                            <col width="3%" >
                            <col width="3%" >
                		</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">차량번호</th>
                                <th scope="col">차종</th>
								<th scope="col">운행자</th>
								<th scope="col">부서</th>
                                <th scope="col">위반일자</th>
								<th scope="col">위반내용</th>
								<th scope="col">과태료</th>
								<th scope="col">위반장소</th>
                                <th scope="col">납입일자</th>
                                <th scope="col">통보일자</th>
                                <th scope="col">통보방법</th>
                                <th scope="col">미납</th>
                                <th scope="col">비고</th>
                                <th scope="col">수정</th>
                                <th scope="col">체크</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

                           car_no = rs("pe_car_no")
						  if rs("pe_in_date") = "1900-01-01"  then
	                               pe_in_date = ""
							   else 
							       pe_in_date = rs("pe_in_date")
	                       end if
	                       if rs("pe_notice_date") = "1900-01-01" then
	                               pe_notice_date = ""
							   else 
							       pe_notice_date = rs("pe_notice_date")
	                       end if
							  
		                   Sql = "SELECT * FROM car_info where car_no = '"&car_no&"'"
                           Set rs_car = DbConn.Execute(SQL)
		                   if not rs_car.eof then
		                        	car_name = rs_car("car_name")
		                    		car_year = rs_car("car_year")
			                    	car_reg_date = rs_car("car_reg_date")
		                    		car_use_dept = rs_car("car_use_dept")
	                    			car_company = rs_car("car_company")
	                     			car_use = rs_car("car_use")
									car_owner = rs_car("car_owner")
	                    			owner_emp_name = rs_car("owner_emp_name")
	                    			owner_emp_no = rs_car("owner_emp_no")
	                     			oil_kind = rs_car("oil_kind")
	                          else
	                     		    car_name = ""
	                    			car_year = ""
			                      	car_reg_date = ""
			                    	car_use_dept = ""
		                    		car_company = ""
		                    		car_use = ""
									car_owner = ""
	                    			owner_emp_name = ""
		                    		owner_emp_no = ""
	                    			oil_kind = ""
                           end if
                           rs_car.close()
						   
						    task_place = replace(rs("pe_place"),chr(34),chr(39))
							view_place = task_place
							if len(task_place) > 10 then
								view_place = mid(task_place,1,10) + ".."
							end if
							
							task_notice = replace(rs("pe_notice"),chr(34),chr(39))
							view_notice = task_notice
							if len(task_notice) > 6 then
								view_notice = mid(task_notice,1,6) + ".."
							end if
							
							if isnull(rs("pe_default")) then
							          pe_default = ""
								else
								      pe_default = rs("pe_default")
							end if
							task_default = replace(pe_default,chr(34),chr(39))
							view_default = task_default
							if len(task_default) > 4 then
								view_default = mid(task_default,1,4) + ".."
							end if
							
							task_bigo = replace(rs("pe_bigo"),chr(34),chr(39))
							view_bigo = task_bigo
							if len(task_bigo) > 6 then
								view_bigo = mid(task_bigo,1,6) + ".."
							end if

	           			%>
							<tr>
								<td class="first"><%=rs("pe_car_no")%>&nbsp;</td>
                                <td><%=car_name%>&nbsp;</td>
                                <td><%=owner_emp_name%>(<%=owner_emp_no%>)&nbsp;</td>
                                <td><%=car_use_dept%>&nbsp;</td>
                                <td><%=rs("pe_date")%>&nbsp;</td>
								<td><%=rs("pe_comment")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pe_amount"),0)%>&nbsp;</td>
                                <td class="left"><p style="cursor:pointer"><span title="<%=task_place%>"><%=view_place%></span></p></td>
                                <td><%=pe_in_date%>&nbsp;</td>
                                <td><%=pe_notice_date%>&nbsp;</td>
                                <td class="left"><p style="cursor:pointer"><span title="<%=task_notice%>"><%=view_notice%></span></p></td>
                                <td class="left"><p style="cursor:pointer"><span title="<%=task_default%>"><%=view_default%></span></p></td>
                                <td class="left"><p style="cursor:pointer"><span title="<%=task_bigo%>"><%=view_bigo%></span></p></td>
                                <td>
                                 <a href="#" onClick="pop_Window('insa_car_penalty_add.asp?car_no=<%=rs("pe_car_no")%>&car_name=<%=car_name%>&car_year=<%=car_year%>&car_reg_date=<%=car_reg_date%>&owner_emp_name=<%=owner_emp_name%>&owner_emp_no=<%=owner_emp_no%>&car_use_dept=<%=car_use_dept%>&oil_kind=<%=oil_kind%>&car_owner=<%=car_owner%>&pe_date=<%=rs("pe_date")%>&pe_seq=<%=rs("pe_seq")%>&u_type=<%="U"%>','car_as_add_popup','scrollbars=yes,width=750,height=410')">수정</a>
                                </td>
                                <td>
                                 <a href="#" onClick="car_penalty_del('<%=rs("pe_car_no")%>', '<%=rs("pe_date")%>', '<%=rs("pe_seq")%>', '<%=car_name%>');return false;">삭제</a></td>
                                </td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
                            <tr>
								<td colspan="4" class="first" style="background:#ffe8e8;">총계</td>
                                <td style="background:#ffe8e8;">과태료 계</td>
                                <td colspan="2" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_amount,0)%>&nbsp;</td>
                                <td style="background:#ffe8e8;">납입 계</td>
                                <td colspan="2" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(tot_in_amt,0)%>&nbsp;</td>
                                <td style="background:#ffe8e8;">미납 계</td>
                                <td colspan="2" class="right" style="font-size:11px; background:#ffe8e8;"><%=formatnumber(jan_amount,0)%>&nbsp;</td>
                                <td colspan="2" style="background:#ffe8e8;">&nbsp;</td>
							</tr>
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
                    <a href="insa_excel_car_penalty.asp?field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_car_penalty_list.asp?page=<%=first_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_car_penalty_list.asp?page=<%=intstart -1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_car_penalty_list.asp?page=<%=i%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_car_penalty_list.asp?page=<%=intend+1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_car_penalty_list.asp?page=<%=total_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
            <% if field_check <> "total" then    %>           
                    <a href="#" onClick="pop_Window('insa_car_penalty_add.asp?car_no=<%=car_no%>&u_type=<%=""%>','car_penalty_add_popup','scrollbars=yes,width=750,height=410')" class="btnType04">차량과태료 등록</a>
            <% end if %>        
					</div>                  
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="pe_car_no" value="<%=pe_car_no%>" ID="Hidden1">
                  <input type="hidden" name="pe_date" value="<%=pe_date%>" ID="Hidden1">
                  <input type="hidden" name="pe_seq" value="<%=pe_seq%>" ID="Hidden1">
                  <input type="hidden" name="car_name" value="<%=car_name%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

