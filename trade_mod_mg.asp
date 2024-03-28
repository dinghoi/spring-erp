<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

Page=Request("page")
use_sw = request("use_sw")  
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
	use_sw = request.form("use_sw")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
	use_sw = request("use_sw")  
end if

if use_sw = "" then
	view_condi = "전체"
	use_sw = "T"
	condi_sql = ""
	condi = ""
	use_sql = ""
end if

if view_condi = "전체" then
	condi = ""
end if

if view_condi = "전체" and use_sw = "T" then
	where_sql = " "
  else
  	where_sql = " where "
end if

if view_condi = "전체" then
	condi_sql = " "
  else
	if condi = "" then
		condi_sql = view_condi + " = '" + condi + "'"
	  else
		condi_sql = view_condi + " like '%" + condi + "%'"
	end if
end if

if use_sw = "T" then
	use_sql = " "
  else
	if condi_sql = " " then
		use_sql = " use_sw = '" + use_sw + "'"
	  else
 		use_sql = " and use_sw = '" + use_sw + "'"
	end if
end if

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "SELECT count(*) FROM trade "&where_sql&condi_sql&use_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM trade "&where_sql&condi_sql&use_sql&" ORDER BY trade_name ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "거래처 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리 회계 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "3 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/sales_code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="trade_mod_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <label>
								<strong>사용구분</strong>
                                <input name="use_sw" type="radio" value="T"  <% if use_sw = "T" then %>checked<% end if %> style="width:25px">총괄
                                <input name="use_sw" type="radio" value="Y"  <% if use_sw = "Y" then %>checked<% end if %> style="width:25px">사용
                                <input name="use_sw" type="radio" value="N"  <% if use_sw = "N" then %>checked<% end if %> style="width:25px">미사용
								</label>
                                <label>
								<strong>조회조건</strong>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="trade_name" <%If view_condi = "trade_name" then %>selected<% end if %>>거래처명</option>
                                  <option value="trade_id" <%If view_condi = "trade_id" then %>selected<% end if %>>거래처유형</option>
                                </select>
								</label>
                                <label>
								<strong>조건 : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="14%" >
							<col width="*" >
							<col width="10%" >
							<col width="8%" >
							<col width="8%" >
							<col width="13%" >
							<col width="13%" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">거래처(회사명)</th>
								<th scope="col">계산서발행회사</th>
								<th scope="col">그룹</th>
								<th scope="col">사업자번호</th>
								<th scope="col">대표자</th>
								<th scope="col">업태</th>
								<th scope="col">업종</th>
								<th scope="col">담당자<br>조회</th>
								<th scope="col">담당자<br>등록</th>
								<th scope="col">변경</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							i = i + 1
							trade_no = mid(rs("trade_no"),1,3) + "-" + mid(rs("trade_no"),4,2) + "-" + mid(rs("trade_no"),6) 
							sql_type="select * from type_code where etc_type='91' and etc_seq ='"+rs("mg_group")+"'"
							set rs_type=dbconn.execute(sql_type)
							if rs_type.eof or rs_type.bof then
								mg_group = "일반그룹"
							  else
								mg_group = rs_type("type_name")
							end if
							rs_type.Close()		
							if rs("use_sw") = "Y" then
								view_use = "사용"
							  else
							  	view_use = "미사용"
							end if
	           			%>
							<tr>
								<td class="first"><%=rs("trade_name")%></td>
								<td><%=rs("bill_trade_name")%>&nbsp;</td>
								<td><%=rs("group_name")%>&nbsp;</td>
								<td><%=trade_no%></td>
								<td><%=rs("trade_owner")%>&nbsp;</td>
								<td><%=rs("trade_uptae")%>&nbsp;</td>
								<td><%=rs("trade_upjong")%>&nbsp;</td>
								<td>조회</td>
								<td><a href="#" onClick="pop_Window('trade_person_mg.asp?trade_code=<%=rs("trade_code")%>','trade_person_mg_pop','scrollbars=yes,width=1000,height=400')">조회</a></td>
								<td><a href="#" onClick="pop_Window('trade_add.asp?trade_code=<%=rs("trade_code")%>&u_type=<%="U"%>','trade_add_pop','scrollbars=yes,width=750,height=400')">변경</a></td>
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
				    <td width="25%">
					<div class="btnCenter">
					</div>                  
                    </td>
				    <td>
                  <div id="paging">
                        <a href = "trade_mod_mg.asp?page=<%=first_page%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="trade_mod_mg.asp?page=<%=intstart -1%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="trade_mod_mg.asp?page=<%=i%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="trade_mod_mg.asp?page=<%=intend+1%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="trade_mod_mg.asp?page=<%=total_page%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="25%">
					<div class="btnRight">
					</div>                  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

