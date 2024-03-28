<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt
Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "insa_sawo_ask_list.asp"
curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	from_date=request("from_date")
    to_date=request("to_date") 
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
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
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_mem = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

'sawo_mem에 있는지 체크해야함
'if give_ask_process = "2" then
'    Sql="select * from emp_sawo_mem where sawo_empno = '"&give_empno&"'"
'    Rs_mem.Open Sql, Dbconn, 1
 
'    sawo_give_cnt = 0
'	sawo_give_pay = 0
'    if not Rs_mem.eof then
'       sawo_give_cnt = Rs_mem("sawo_give_count")
'       sawo_give_pay = Rs_mem("sawo_give_pay")
'    end if
'    Rs_mem.Close()

'    sawo_give_cnt = sawo_give_cnt + 1
'    sawo_give_pay = sawo_give_pay + give_pay
'end if

	Sql= "select count(*) " & _
	          "    from emp_sawo_ask a, emp_master b " & _
	          "    where (a.ask_empno = b.emp_no) AND (b.emp_sawo_id = 'Y') and (a.ask_process = '0') and (a.ask_company = '"+view_condi+"') and (Cast(a.ask_reg_date as date) >= '" + from_date + "' AND Cast(a.ask_reg_date as date) <= '"+to_date+"')" & _
		      "    ORDER BY ask_company,ask_date,ask_empno DESC" 

'order_Sql = " ORDER BY ask_company,ask_date,ask_empno DESC"
'where_sql = " WHERE (ask_process = '0') and (ask_company = '"+view_condi+"') and (Cast(ask_reg_date as date) >= '" + from_date + "' AND Cast(ask_reg_date as date) <= '"+to_date+"')"

'Sql = "SELECT count(*) FROM emp_sawo_ask " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

	Sql= "select * " & _
	          "    from emp_sawo_ask a, emp_master b " & _
	          "    where (a.ask_empno = b.emp_no) AND (b.emp_sawo_id = 'Y') and (a.ask_process = '0') and (a.ask_company = '"+view_condi+"') and (Cast(a.ask_reg_date as date) >= '" + from_date + "' AND Cast(a.ask_reg_date as date) <= '"+to_date+"')" & _
		      "    ORDER BY ask_company,ask_date,ask_empno DESC limit "& stpage & "," &pgsize

'sql = "select * from emp_sawo_ask " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " 경조금 신청 현황 "
give_ask_process = "2"

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
				return "8 1";
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sawo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_sawo_ask_list.asp" method="post" name="frm">
                <fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where  org_level = '회사' ORDER BY org_code ASC"
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
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="9%" >
                            <col width="9%" >
							<col width="6%" >
							<col width="4%" >
							<col width="4%" >
							<col width="12%" >
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
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
								<th scope="col">경조일시</th>
								<th scope="col">경조<br>구분</th>
								<th scope="col">경조<br>유형</th>
                                <th scope="col">경조장소</th>
                                <th scope="col">기타내역</th>
                                <th scope="col">첨부</th>
								<th scope="col">경조금</th>
							</tr>
						</thead>
					<tbody>
						<%
						
						do until rs.eof
						 
		                  ask_empno = rs("ask_empno")
		                  ask_emp_name = rs("ask_emp_name")
		
                         if ask_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&ask_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_grade = Rs_emp("emp_grade")
		                      emp_position = Rs_emp("emp_position")
		                   end if
	                       Rs_emp.Close()
	                	 end if		
						%>
							<tr>
								<td class="first"><%=rs("ask_empno")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("ask_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("ask_emp_name")%></a>
								</td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=rs("ask_company")%>&nbsp;</td>
                                <td><%=rs("ask_org_name")%>&nbsp;</td>
                                <td><%=rs("ask_date")%>&nbsp;</td>
                                <td><%=rs("ask_id")%>&nbsp;</td>
                                <td><%=rs("ask_type")%>&nbsp;</td>
                                <td><%=rs("ask_sawo_place")%>&nbsp;</td>
                                <td><%=rs("ask_sawo_comm")%>&nbsp;</td>
                                <td>
								<% 
                                If rs("ask_att_file") <> "" Then 
                                    path = "/emp_sawo/" 
                                %>
                                  <a href="att_file_download.asp?path=<%=path%>&att_file=<%=rs("ask_att_file")%>"><img src="image/att_file.gif" border="0"></a>
                                <% Else %>
				                    &nbsp;
                                <% 
								End If %>
                                </td>
                                <td><a href="#" onClick="pop_Window('insa_sawo_giveadd.asp?sawo_empno=<%=rs("ask_empno")%>&emp_name=<%=rs("ask_emp_name")%>&ask_seq=<%=rs("ask_seq")%>&ask_date=<%=rs("ask_date")%>&give_ask_process=<%=give_ask_process%>&u_type=<%=""%>','insa_sawo_giveadd_pop','scrollbars=yes,width=750,height=450')">지급</a>&nbsp;</td>
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
                    <a href="insa_excel_sawo_ask.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_sawo_ask_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_sawo_ask_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_sawo_ask_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_sawo_ask_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_sawo_ask_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
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
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>

