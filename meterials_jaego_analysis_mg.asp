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

be_pg = "meterials_jaego_analysis_mg.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
	goods_type=Request("goods_type")
	owner_view=request("owner_view")
	condi = request("condi")
  else
	view_condi=Request.form("view_condi")
	goods_type=Request.form("goods_type")
	owner_view=Request.form("owner_view")
	condi = request.form("condi")
End if

If view_condi = "" Then
	view_condi = "케이원정보통신"
	goods_type = "상품"
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
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY stock_company,stock_level,stock_goods_type,stock_goods_code,stock_code ASC"
if condi = "" then
      where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"')" 
   else  
      if owner_view = "C" then 
             where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_name like '%"+condi+"%')"
         else
		     where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_code = '"+condi+"')"
	   end if
end if   
Sql = "SELECT count(*) FROM met_stock_gmaster " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_stock_gmaster " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 재고변동 추이 "

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
				return "5 1";
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
			<!--#include virtual = "/include/meterials_control_header.asp" -->
            <!--#include virtual = "/include/meterials_jaego_analysis_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="meterials_jaego_analysis_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>회사 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:120px">
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
                                <strong>용도구분 : </strong>
                              <%
								Sql="select * from met_etc_code where etc_type = '01' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="goods_type" id="goods_type" type="text" style="width:90px">
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
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">품목코드
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">품목명
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
				      <col width="10%" >
				      <col width="12%" >
                      <col width="*" >
                      <col width="5%" >
				      <col width="5%" >
				      <col width="5%" >
                      <col width="5%" >
				      <col width="5%" >
				      <col width="5%" >
				      <col width="5%" >
                      <col width="5%" >
				      <col width="5%" >
                      <col width="5%" >
				      <col width="5%" >
                      <col width="5%" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th rowspan="2" class="first" scope="col">회사</th>
				        <th rowspan="2" scope="col">품목명</th>
                        <th rowspan="2" scope="col">창고</th>
                        <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">전년이월</th>
				        <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">입고</th>
                        <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">출고</th>
                        <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">현재고</th>
                        <th rowspan="2" scope="col">비고</th>
			          </tr>
                      <tr>
				        <th scope="col" style=" border-left:1px solid #e3e3e3;">A급</th>
                        <th scope="col">B급</th>
                        <th scope="col">C급</th>
                        <th scope="col">A급</th>
                        <th scope="col">B급</th>
                        <th scope="col">C급</th>
                        <th scope="col">A급</th>
                        <th scope="col">B급</th>
                        <th scope="col">C급</th>
                        <th scope="col">A급</th>
                        <th scope="col">B급</th>
                        <th scope="col">C급</th>
                      </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof
						   a_cnt = int(rs("stock_last_a_cnt")) + int(rs("stock_in_a_cnt")) - int(rs("stock_go_a_cnt"))
						   b_cnt = int(rs("stock_last_b_cnt")) + int(rs("stock_in_b_cnt")) - int(rs("stock_go_b_cnt"))
						   c_cnt = int(rs("stock_last_c_cnt")) + int(rs("stock_in_c_cnt")) - int(rs("stock_go_c_cnt"))
					  %>
				      <tr>
				        <td class="first"><%=rs("stock_company")%>&nbsp;</td>
                        <td><%=rs("stock_goods_name")%>(<%=rs("stock_goods_code")%>)&nbsp;</td>
                        <td><%=rs("stock_name")%>(<%=rs("stock_code")%>)&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_last_a_cnt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_last_b_cnt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_last_c_cnt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_in_a_cnt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_in_b_cnt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_in_c_cnt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_go_a_cnt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_go_b_cnt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_go_c_cnt"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(a_cnt,0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(b_cnt,0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(c_cnt,0)%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('met_stock_gmaster_modify.asp?stock_code=<%=rs("stock_code")%>&stock_goods_code=<%=rs("stock_goods_code")%>&u_type=<%="U"%>','met_gmaster_modi_pop','scrollbars=yes,width=1250,height=400')">수정</a>&nbsp;</td>
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
                    <a href="met_excel_gmaster_modify.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="meterials_jaego_analysis_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="meterials_jaego_analysis_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="meterials_jaego_analysis_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="meterials_jaego_analysis_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[다음]</a> <a href="meterials_jaego_analysis_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <% if owner_view = "T" then 
                              goods_code = condi
							  Sql = "SELECT * FROM met_goods_code where goods_code = '"&goods_code&"'"
                              Set Rs_good = DbConn.Execute(SQL)
							  if not Rs_good.eof then
                                   goods_type = Rs_good("goods_type")
								   goods_name = Rs_good("goods_name")
							  end if
							  Rs_good.close()
				    %>
                    <a href="#" onClick="pop_Window('met_stock_gmaster_add.goods_code=<%=goods_code%>&goods_name=<%=goods_name%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&u_type=<%=""%>','met_stock_gmaster_popup','scrollbars=yes,width=1250,height=400')" class="btnType04">재고조정 등록</a>
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

