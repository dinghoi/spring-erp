<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_report_appoint.asp"

from_date=Request.form("from_date")
to_date=Request.form("to_date")

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	app_id = request.form("app_id")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	app_id = request("app_id")
	from_date=request("from_date")
    to_date=request("to_date")
end if

if view_condi = "" then
	view_condi = "전체"
	app_id = "전체"
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
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "전체" then
       Sql = "SELECT count(*) from emp_appoint where app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000')"
   else  
       Sql = "select count(*) from emp_appoint where app_to_company='"+view_condi+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000')"
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if view_condi = "전체" then
   if app_id = "전체" then
           Sql = "select * from emp_appoint where app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC limit "& stpage & "," &pgsize 
	  else 
		   Sql = "select * from emp_appoint where app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC limit "& stpage & "," &pgsize 
   end if	   
   else  
      if app_id = "전체" then
	          Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC limit "& stpage & "," &pgsize 
		 else	  
			  Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC limit "& stpage & "," &pgsize 
	  end if
end if
Rs.Open Sql, Dbconn, 1

'Response.write Sql

title_line = view_condi +" - 인사발령 현황(" + from_date + " ∼ " + to_date + ")"
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
				return "2 1";
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
			<!--#include virtual = "/include/insa_appoint_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_report_appoint.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
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
                                <strong>발령구분</strong>
                            <%
								Sql="select * from emp_etc_code where emp_etc_type = '10' order by emp_etc_code asc"
								Rs_etc.Open Sql, Dbconn, 1
							%>
								<select name="app_id" id="select" type="text" style="width:150px">
                                    <option value="전체" <%If app_id = "전체" then %>selected<% end if %>>전체</option>
                			<% 
								do until rs_etc.eof 
			  				%>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If app_id = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%>&nbsp;</option>
                			<%
									rs_etc.movenext()
								loop 
								rs_etc.Close()
							%>
            					</select>
								</label>
								<label>
								<strong>발령일(From) : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong> ∼ To : </strong>
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
							<col width="5%" >
                            <col width="5%" >
                            <col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="9%" >
							<col width="10%" >
							<col width="9%" >
							<col width="9%" >
							<col width="10%" >
                            <col width="9%" >
                            <col width="*" >
						</colgroup>
						<thead>
                            <tr>
				                <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">사번</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">성명</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령일</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령구분</th>
                                <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령유형</th>
                                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령전</th>
				                <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">발령후</th>
			                </tr>
                            <tr>
                                <th class="first"scope="col" style=" border-left:1px solid #e3e3e3;">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">직급/책</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">직급/책</th>
                                <th scope="col">발령내용</th>
                            </tr>
						</thead>
						<tbody>
						<%
					  	   do until rs.eof

	           			%>
							<tr>
								<td><%=rs("app_empno")%>&nbsp;</td>
                                <td><%=rs("app_emp_name")%>&nbsp;</td>
                                <td><%=rs("app_date")%>&nbsp;</td>
								<td><%=rs("app_id")%>&nbsp;</td>
                                <td><%=rs("app_id_type")%>&nbsp;</td>
                                <td><%=rs("app_to_company")%>&nbsp;</td>
                                <td><%=rs("app_to_org")%>(<%=rs("app_to_orgcode")%>)&nbsp;</td>
                                <td><%=rs("app_to_grade")%>-<%=rs("app_to_position")%>&nbsp;</td>
                                <td><%=rs("app_be_company")%>&nbsp;</td>
                                <td><%=rs("app_be_org")%>(<%=rs("app_be_orgcode")%>)&nbsp;</td>
                                <td><%=rs("app_be_grade")%>-<%=rs("app_be_position")%>&nbsp;</td>
                                <td class="left"><%=rs("app_start_date")%>&nbsp;-&nbsp;<%=rs("app_finish_date")%>&nbsp;<%=rs("app_be_enddate")%>&nbsp;<%=rs("app_reward")%>&nbsp;:&nbsp;<%=rs("app_comment")%>&nbsp;</td>
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
                    <a href="insa_excel_appoint.asp?view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_report_appoint.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_report_appoint.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_report_appoint.asp?page=<%=i%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_report_appoint.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_report_appoint.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
				    <td width="15%">
					<div class="btnCenter">
            <a href="#" onClick="pop_Window('insa_appoint_print.asp?view_condi=<%=view_condi%>&app_id=<%=app_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>','pop_report','scrollbars=yes,width=1250,height=600')" class="btnType04">출력</a>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

