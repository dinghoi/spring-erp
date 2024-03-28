<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

stock_level = request("stock_level")
owner_view=request("owner_view")
condi = request("condi")

be_pg = "met_stock_code_mg.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	stock_level=Request("stock_level")
	owner_view=request("owner_view")
	condi = request("condi")
  else
	stock_level=Request.form("stock_level")
	owner_view=Request.form("owner_view")
	condi = request.form("condi")
End if

If stock_level = "" Then
	stock_level = "팀"
	owner_view = "C"
	ck_sw = "n"
	condi = ""
End If

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
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY stock_level,stock_code ASC"
if condi = "" then
      where_sql = " WHERE (stock_level = '"&stock_level&"')" 
   else  
      if owner_view = "C" then 
             where_sql = " WHERE (stock_level = '"&stock_level&"') and (stock_name like '%"+condi+"%')"
         else
		     where_sql = " WHERE (stock_level = '"&stock_level&"') and (stock_code = '"+condi+"')"
	   end if
end if   
Sql = "SELECT count(*) FROM met_stock_code " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_stock_code " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = stock_level + " 창고 현황 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>상품자재관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "6 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.stock_level.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/meterials_control_header01.asp" -->
            <!--#include virtual = "/include/meterials_basic_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_stock_code_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt> 창고 검색</dt>
                        <dd>
                            <p>
                                <strong>창고유형 : </strong>
                              <%
								Sql="select * from met_etc_code where etc_type = '20' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="stock_level" id="stock_level" type="text" style="width:90px">
                			  <% 
								do until Rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("etc_name")%>' <%If stock_level = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			  <%
									Rs_etc.movenext()  
								loop 
								Rs_etc.Close()
							  %>
            					</select>
                                </label>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">창고코드
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">창고명
                                </label>
							<strong>조건 : </strong>
								<label>
        						<input name="condi" type="text" id="condi" value="<%=condi%>" style="width:100px; text-align:left">
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
				      <col width="10%" >
                      <col width="6%" >
				      <col width="10%" >
				      <col width="10%" >
                      <col width="6%" >
				      <col width="6%" >
                      <col width="6%" >
                      <col width="6%" >
				      <col width="*" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">창고코드</th>
				        <th scope="col">창고명</th>
                        <th scope="col">창고유형</th>
                        <th scope="col">창고장</th>
                        <th scope="col">회사</th>
                        <th scope="col">생성일</th>
                        <th scope="col">폐쇄일</th>
                        <th scope="col">출고담당</th>
                        <th scope="col">입고담당</th>
                        <th scope="col">소속조직</th>
                        <th scope="col">비고</th>
			          </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof 
						   stock_end_date = rs("stock_end_date")
						   if stock_end_date = "1900-01-01" then
	                            stock_end_date = ""
	                       end if
					  %>
				      <tr>
				        <td class="first"><%=rs("stock_code")%>&nbsp;</td>
                        <td><%=rs("stock_name")%>&nbsp;</td>
                        <td><%=rs("stock_level")%>&nbsp;</td>
                        <td><%=rs("stock_manager_name")%>(<%=rs("stock_manager_code")%>)&nbsp;</td>
                        <td><%=rs("stock_company")%>&nbsp;</td>
                        <td><%=rs("stock_open_date")%>&nbsp;</td>
                        <td><%=stock_end_date%>&nbsp;</td>
                        <td><%=rs("stock_go_name")%>&nbsp;</td>
                        <td><%=rs("stock_in_name")%>&nbsp;</td>
                        <td class="left"><%=rs("stock_bonbu")%>-<%=rs("stock_saupbu")%>-<%=rs("stock_team")%>&nbsp;</td>
                    <% if stock_level <> "개인" then %>
                        <td><a href="#" onClick="pop_Window('met_stock_code_add.asp?stock_code=<%=rs("stock_code")%>&stock_name=<%=rs("stock_name")%>&stock_level=<%=rs("stock_level")%>&u_type=<%="U"%>','met_stock_code_pop','scrollbars=yes,width=750,height=300')">수정</a>&nbsp;</td>
                    <% else %>
                        <td>&nbsp;</td>
                    <% end if %>
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
                    <a href="met_stock_code_excel.asp?stock_level=<%=stock_level%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_stock_code_mg.asp?page=<%=first_page%>&stock_level=<%=stock_level%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_stock_code_mg.asp?page=<%=intstart -1%>&stock_level=<%=stock_level%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_stock_code_mg.asp?page=<%=i%>&stock_level=<%=stock_level%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_stock_code_mg.asp?page=<%=intend+1%>&stock_level=<%=stock_level%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_stock_code_mg.asp?page=<%=total_page%>&stock_level=<%=stock_level%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                 <% if stock_level <> "개인" then %>
                    <a href="#" onClick="pop_Window('met_stock_code_add.asp?stock_level=<%=stock_level%>&owner_view=<%=owner_view%>&condi=<%=condi%>&u_type=<%=""%>','met_stock_code_pop','scrollbars=yes,width=750,height=300')" class="btnType04">신규 창고 등록</a>
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

