<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
end if

if view_condi = "" then
	view_condi = "전체"
	condi_sql = ""
	condi = ""
	use_sql = ""
end if

if view_condi = "전체" then
	condi = ""
end if

if view_condi = "전체" then
	condi_sql = " "
  else
	if condi = "" then
		condi_sql = " where " + view_condi + " = '" + condi + "'"
	  else
		condi_sql = " where " + view_condi + " like '%" + condi + "%'"
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

Sql = "SELECT count(*) FROM trade "&condi_sql&use_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM trade "&condi_sql&use_sql&" ORDER BY trade_name ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'Response.write Sql

title_line = "비용관련 거래처 변경 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
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
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="trade_cost_mod_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <label>
								<strong>조회조건</strong>
                                <select name="view_condi" id="select3" style="width:150px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                                  <option value="trade_name" <%If view_condi = "trade_name" then %>selected<% end if %>>거래처명</option>
                                  <option value="trade_id" <%If view_condi = "trade_id" then %>selected<% end if %>>거래처유형</option>
                                  <option value="emp_name" <%If view_condi = "emp_name" then %>selected<% end if %>>담당자</option>
                                  <option value="saupbu" <%If view_condi = "saupbu" then %>selected<% end if %>>담당사업부</option>
                                  <option value="group_name" <%If view_condi = "group_name" then %>selected<% end if %>>그룹명</option>
                                  <option value="saupbu" <%If view_condi = "saupbu" then %>selected<% end if %>>담당사업부</option>
                                  <option value="bill_trade_name" <%If view_condi = "bill_trade_name" then %>selected<% end if %>>계산서발행회사</option>
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
							<col width="8%" >
							<col width="12%" >
							<col width="*" >
							<col width="10%" >
							<col width="6%" >
							<col width="8%" >
							<col width="12%" >
							<col width="5%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">거래처(회사명)</th>
								<th scope="col">사업자번호</th>
								<th scope="col">대표자</th>
								<th scope="col">계산서 발행 회사명</th>
								<th scope="col">그룹</th>
								<th scope="col">유형</th>
								<th scope="col">담당자</th>
								<th scope="col">담당사업부</th>
								<th scope="col">사용</th>
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
							if rs("trade_id") = "매출" then
								view_trade_id = "AS"
							  else
							  	view_trade_id = rs("trade_id")
							end if
	           			%>
							<tr>
								<td class="first"><%=rs("trade_name")%></td>
								<td><%=trade_no%></td>
								<td><%=rs("trade_owner")%>&nbsp;</td>
								<td><%=rs("bill_trade_name")%>&nbsp;</td>
								<td><%=rs("group_name")%>&nbsp;</td>
								<td><%=view_trade_id%></td>
								<td><%=rs("emp_no")%>&nbsp;<%=rs("emp_name")%></td>
								<td><%=rs("saupbu")%>&nbsp;</td>
								<td><%=view_use%></td>
								<td><a href="#" onClick="pop_Window('trade_cost_mod.asp?trade_code=<%=rs("trade_code")%>','trade_cost_mod_pop','scrollbars=yes,width=800,height=250')">변경</a></td>
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
                    <a href="trade_excel.asp?use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                    </td>
				    <td>
                    <div id="paging">
                        <a href = "trade_cost_mod_mg.asp?page=<%=first_page%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	  <% if intstart > 1 then %>
                        <a href="trade_cost_mod_mg.asp?page=<%=intstart -1%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           				<% if i = int(page) then %>
							<b>[<%=i%>]</b>
						<% else %>
                        <a href="trade_cost_mod_mg.asp?page=<%=i%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      	<% end if %>
                      <% next %>
           				<% if 	intend < total_page then %>
                        	<a href="trade_cost_mod_mg.asp?page=<%=intend+1%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="trade_cost_mod_mg.asp?page=<%=total_page%>&use_sw=<%=use_sw%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
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

