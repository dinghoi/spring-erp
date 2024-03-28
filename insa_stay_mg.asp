<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim company_tab(50)
dim page_cnt
dim pg_cnt

in_stay_code =""
in_stay_name = ""

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "insa_org.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

if page_cnt > 0 then 
	pg_cnt = page_cnt
end if
if pg_cnt > 0 then
	page_cnt = pg_cnt
end if

if page_cnt < 10 or page_cnt > 20 then
	page_cnt = 10
end if

pgsize = page_cnt ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

view_sort = request("view_sort")

if view_sort = "" then
	view_sort = "DESC"
end if

order_Sql = " ORDER BY stay_name " + view_sort

'where_sql = " WHERE isNull(org_end_date)"

Sql = "SELECT count(*) FROM emp_stay "
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_stay " + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1


title_line = "[ 실근무지 현황 ]"

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
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.condi.value == "") {
					alert ("실근무지을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_stay_mg.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="4%" >
				      <col width="10%" >
				      <col width="9%" >
				      <col width="10%" >
				      <col width="12%" >
				      <col width="12%" >
				      <col width="10%" >
				      <col width="10%" >
                      <col width="10%" >
                      <col width="10%" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
                      <tr>
				        <th class="first" scope="col">코드</th>
				        <th scope="col">실근무지명</th>
				        <th colspan="4" scope="col">주&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;소</th>
				        <th scope="col">전화번호</th>
				        <th scope="col">팩스번호</th>
                        <th scope="col">상주처명</th>
                        <th scope="col">상주처회사</th>
                        <th scope="col">수정</th>
                      </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof
					  %>
				      <tr>
				        <td class="first"><%=rs("stay_code")%>&nbsp;</td>
				        <td><%=rs("stay_name")%>&nbsp;</td>
                        <td><%=rs("stay_sido")%>&nbsp;</td>
                        <td><%=rs("stay_gugun")%>&nbsp;</td>
                        <td><%=rs("stay_dong")%>&nbsp;</td>
                        <td><%=rs("stay_addr")%>&nbsp;</td>
                        <td><%=rs("stay_tel_ddd")%>-<%=rs("stay_tel_no1")%>-<%=rs("stay_tel_no2")%>&nbsp;</td>
                        <td><%=rs("stay_fax_ddd")%>-<%=rs("stay_fax_no1")%>-<%=rs("stay_fax_no2")%>&nbsp;</td>
                        <td><%=rs("stay_org_name")%>&nbsp;</td>
                        <td><%=rs("stay_reside_company")%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('insa_stay_add.asp?stay_code=<%=rs("stay_code")%>&u_type=<%="U"%>','insa_stay_add_pop','scrollbars=yes,width=750,height=600')">수정</a>&nbsp;</td>
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
				    <td>
                    <div id="paging">
                        <a href="insa_stay_mg.asp?page=<%=first_page%>&view_sort=<%=view_sort%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_stay_mg.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_stay_mg.asp?page=<%=i%>&view_sort=<%=view_sort%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="insa_stay_mg.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>">[다음]</a> <a href="insa_stay_mg.asp?page=<%=total_page%>&view_sort=<%=view_sort%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%">
					<div class="btnCenter">
					<a href="#" onClick="pop_Window('insa_stay_add.asp?stay_code=<%=in_stay_code%>&stay_name=<%=in_stay_name%>','insa_stay_add_pop','scrollbars=yes,width=750,height=400')" class="btnType04">실근무지 등록</a>
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

