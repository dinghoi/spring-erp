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

view_condi = request("view_condi")
owner_view=request("owner_view")
condi = request("condi")

be_pg = "met_organiz_jaego_mg.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
	goods_type=Request("goods_type")
	field_check=Request("field_check")
	field_view=Request("field_view")
  else
	view_condi=Request.form("view_condi")
	goods_type=Request.form("goods_type")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if

If view_condi = "" Then
	view_condi = "케이원정보통신"
'	goods_type = "상품"
	goods_type = "전체"
	field_check = "total"
	ck_sw = "n"
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
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

'if field_check <> "total" then
'	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
'  else
'  	field_sql = " "
'end if

if field_check = "total" then
	   field_sql = " "
  elseif field_check = "본사" then  
	         field_sql = " and ( stock_company like '%" + field_view + "%' ) "
	     elseif field_check = "본부" then  
	                field_sql = " and ( stock_bonbu like '%" + field_view + "%' ) "
			    elseif field_check = "사업부" then  
	                       field_sql = " and ( stock_saupbu like '%" + field_view + "%' ) "
					   elseif field_check = "팀" then  
	                              field_sql = " and ( stock_team like '%" + field_view + "%' ) "
end if

'order_Sql = " ORDER BY stock_company,stock_goods_grade,stock_level,stock_goods_type,stock_goods_code,stock_code ASC"
order_Sql = " ORDER BY stock_goods_grade,stock_goods_name,stock_name ASC"

if goods_type = "전체" then
        Sql = "SELECT count(*) FROM met_stock_gmaster where (stock_company = '"&view_condi&"') " + field_sql
   else
        Sql = "SELECT count(*) FROM met_stock_gmaster where (stock_company = '"&view_condi&"') and (stock_goods_type = '"&goods_type&"') " + field_sql
end if

Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if goods_type = "전체" then
        sql = "select * from met_stock_gmaster where (stock_company = '"&view_condi&"') " + field_sql + order_sql + " limit "& stpage & "," &pgsize 
   else		
		sql = "select * from met_stock_gmaster where (stock_company = '"&view_condi&"') and (stock_goods_type = '"&goods_type&"') " + field_sql + order_sql + " limit "& stpage & "," &pgsize 
end if

Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 조직별 재고현황 "

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
				return "4 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
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
            <!--#include virtual = "/include/meterials_stock_jaego_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_organiz_jaego_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>검색조건</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:120px">
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
                                <strong>용도구분 : </strong>
                              <%
								Sql="select * from met_etc_code where etc_type = '01' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="goods_type" id="goods_type" type="text" style="width:90px">
                                    <option value="전체" <%If goods_type = "전체" then %>selected<% end if %>>전체</option>
                			  <% 
								do until Rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("etc_name")%>' <%If goods_type = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			  <%
									Rs_etc.movenext()  
								loop 
								Rs_etc.Close()
							  %>
            					</select>
                                </label>                                
                                <label>
								<strong>조직 조건</strong>
                                <select name="field_check" id="field_check" style="width:100px">
                                  <option value="total" <% if field_check = "total" then %>selected<% end if %>>전체</option>
                                  <option value="본사" <% if field_check = "본사" then %>selected<% end if %>>본사</option>
                                  <option value="본부" <% if field_check = "본부" then %>selected<% end if %>>본부</option>
                                  <option value="사업부" <% if field_check = "사업부" then %>selected<% end if %>>사업부</option>
                                  <option value="팀" <% if field_check = "팀" then %>selected<% end if %>>팀</option>
                                  <option value="개인" <% if field_check = "개인" then %>selected<% end if %>>개인</option>
                                </select>
								<strong>조직 명</strong>
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
				      <col width="7%" >
                      <col width="7%" >
                      <col width="12%" >
                      <col width="12%" >
                      <col width="3%" >
                      <col width="*" >
                      <col width="5%" >
				      <col width="7%" >
				      <col width="5%" >
                      <col width="7%" >
				      <col width="5%" >
				      <col width="7%" >
                      <col width="5%" >
				      <col width="7%" >
			        </colgroup>
				    <thead>
                      <tr>
				        <th rowspan="2" class="first" scope="col">코드</th>
				        <th rowspan="2" scope="col">품목구분</th>
                        <th rowspan="2" scope="col">품목명</th>
                        <th rowspan="2" scope="col">규격</th>
                        <th rowspan="2" scope="col">상태</th>
                        <th rowspan="2" scope="col">창고</th>
                        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">전년이월</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">입고</th>
                        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">출고</th>
                        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">현재고</th>
			          </tr>
                      <tr>
				        <th scope="col" style=" border-left:1px solid #e3e3e3;">수량</th>
                        <th scope="col">금액</th>
                        <th scope="col">수량</th>
                        <th scope="col">금액</th>
                        <th scope="col">수량</th>
                        <th scope="col">금액</th>
                        <th scope="col">수량</th>
                        <th scope="col">금액</th>
                      </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof

					  %>
				      <tr>
				        <td class="first"><%=rs("stock_goods_code")%>&nbsp;</td>
                        <td><%=rs("stock_goods_gubun")%>&nbsp;</td>
                        <td><%=rs("stock_goods_name")%>&nbsp;</td>
                        <td><%=rs("stock_goods_standard")%>&nbsp;</td>
                        <td><%=rs("stock_goods_grade")%>&nbsp;</td>
                        <td><%=rs("stock_name")%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_last_qty"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_last_amt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_in_amt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_go_amt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_jj_qty"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_jj_amt"),0)%>&nbsp;</td>
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
                    <a href="met_organiz_jaego_excel.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_organiz_jaego_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_organiz_jaego_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_organiz_jaego_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_organiz_jaego_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_organiz_jaego_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('met_organiz_jaego_print.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&u_type=<%=""%>','met_goods_jaego_popup','scrollbars=yes,width=1250,height=600')" class="btnType04">재고현황 출력</a>
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

