<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_car_as_list.asp"

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

owner_sql = " where as_date >= '"+from_date+"' and as_date <= '"+to_date+"' "
order_sql = " ORDER BY as_car_no,as_date,as_seq DESC"

if field_check <> "total" then
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if

sql = "select count(*) FROM car_as " + owner_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from car_as " + owner_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = ""+ view_condi +" - 차량 정비현황 "
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
				<form action="insa_car_as_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
                                  <option value="as_car_no" <% if field_check = "as_car_no" then %>selected<% end if %>>차량번호</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:100px; text-align:left" >
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
							<col width="8%" >
                            <col width="10%" >
							<col width="10%" >
							<col width="12%" >
                            <col width="7%" >
							<col width="20%" >
							<col width="*" >
							<col width="6%" >
                		</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">차량번호</th>
                                <th scope="col">차종</th>
								<th scope="col">운행자</th>
								<th scope="col">부서</th>
                                <th scope="col">AS일자</th>
								<th scope="col">AS증상</th>
								<th scope="col">수리내용</th>
								<th scope="col">수리비용</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
                           as_car_no = rs("as_car_no") 
						   
						    task_memo = replace(rs("as_solution"),chr(34),chr(39))
							view_memo = task_memo
							if len(task_memo) > 30 then
								view_memo = mid(task_memo,1,30) + ".."
							end if
	           			%>
							<tr>
								<td class="first"><%=rs("as_car_no")%></td>
                                <td><%=rs("as_car_name")%></td>
                                <td><%=rs("as_owner_emp_name")%>(<%=rs("as_owner_emp_no")%>)</td>
                                <td><%=rs("as_use_org_name")%>&nbsp;</td>
                                <td><%=rs("as_date")%></td>
								<td class="left"><%=rs("as_cause")%></td>
                                <td class="left"><p style="cursor:pointer"><span title="<%=task_memo%>"><%=view_memo%></span></p></td>
                                <td class="right"><%=formatnumber(rs("as_amount"),0)%></td>
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
                    <a href="insa_excel_car_as.asp?field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_car_as_list.asp?page=<%=first_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_car_as_list.asp?page=<%=intstart -1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_car_as_list.asp?page=<%=i%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_car_as_list.asp?page=<%=intend+1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_car_as_list.asp?page=<%=total_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
            <% if field_check <> "total" then    %>           
                    <a href="#" onClick="pop_Window('insa_car_as_add.asp?car_no=<%=as_car_no%>&u_type=<%=""%>','car_as_add_popup','scrollbars=yes,width=750,height=350')" class="btnType04">차량정비내역 등록</a>
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

