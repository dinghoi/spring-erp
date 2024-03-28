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
org_name = request.cookies("nkpmg_user")("coo_org_name")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")
emp_company = request.cookies("nkpmg_user")("coo_emp_company")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

view_condi = request("view_condi")
owner_view=request("owner_view")
condi = request("condi")

be_pg = "met_stock_out_sale_mg.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
	stock = request("stock")
	goods_type=Request("goods_type")
	owner_view=request("owner_view")
	condi = request("condi")
  else
	view_condi=Request.form("view_condi")
	stock = request.form("stock")
	goods_type=Request.form("goods_type")
	owner_view=Request.form("owner_view")
	condi = request.form("condi")
End if

If view_condi = "" Then
	view_condi = emp_company
	stock = user_name
	goods_type = "상품"
	owner_view = "C"
	ck_sw = "n"
	condi = ""
End If

chulgo_id = "영업출고"

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

order_Sql = " ORDER BY stock_company,stock_goods_grade,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_code ASC"

if goods_type = "전체" then 
   if condi = "" then
         where_sql = " WHERE (stock_company = '"&view_condi&"')" 
      else  
         if owner_view = "C" then 
                where_sql = " WHERE (stock_company = '"&view_condi&"') and (stock_goods_name like '%"+condi+"%')"
            else
		        where_sql = " WHERE (stock_company = '"&view_condi&"') and (stock_goods_code like '%"+condi+"%')"
   	      end if
   end if
  else
   if condi = "" then
         where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"')" 
      else  
         if owner_view = "C" then 
                where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_name like '%"+condi+"%')"
            else
		        where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_code like '%"+condi+"%')"
	      end if
   end if   
end if

'if stock = "" then
'       stock_sql = ""
'   else
'       stock_sql = " and (stock_code = '"&stock&"') "
'end if

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (stock_name like '%"&stock&"%') "
end if

' 입출고 및 재고수량이 없는것
'cnt_sql = " and (stock_in_qty > 0 or stock_go_qty > 0 or stock_JJ_qty > 0) "
cnt_sql = " and (stock_JJ_qty > 0) "

Sql = "SELECT count(*) FROM met_stock_gmaster " + where_sql + stock_sql + cnt_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_stock_gmaster " + where_sql + stock_sql + cnt_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 상품 고객사 출고 "

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
				return "7 1";
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
            <!--#include virtual = "/include/meterials_stock_saleout_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_stock_out_sale_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>검색 조건</dt>
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
                                <label>
                                <strong>창고(영업) 명 : </strong>
                                   <input name="stock" type="text" id="stock" value="<%=stock%>" style="width:100px; text-align:left; ime-mode:active">
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
				      <col width="7%" >
                      <col width="8%" >
                      <col width="17%" >
                      <col width="14%" >
                      <col width="14%" >
                      <col width="4%" >
                      <col width="*" >
				      <col width="6%" >
                      <col width="5%" >
				      <col width="5%" >
				      <col width="6%" >
                      <col width="5%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">코드</th>
				        <th scope="col">품목구분</th>
                        <th scope="col">품목명</th>
                        <th scope="col">규격</th>
                        <th scope="col">모델</th>
                        <th scope="col">상태</th>
                        <th scope="col">창고</th>
                        <th scope="col">전년이월</th>
				        <th scope="col">입고</th>
                        <th scope="col">출고</th>
                        <th scope="col">현재고</th>
                        <th scope="col">출고<br>등록</th>
			          </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof
                           stock_level = rs("stock_level")
						   stock_code = rs("stock_code")
						   stock_goods_type = rs("stock_goods_type")
						   stock_goods_code = rs("stock_goods_code")
						   stock_jj_qty = rs("stock_jj_qty")
						   
						   sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
                           Set Rs_good = DbConn.Execute(SQL)
                           if not Rs_good.eof then
    	                         goods_model = Rs_good("goods_model")
		                         goods_serial_no = Rs_good("goods_serial_no")
                              else
		                         goods_model = ""
		                         goods_serial_no = ""
                           end if
                           Rs_good.close()
					  %>
				      <tr>
				        <td class="first">
                        
                        <a href="#" onClick="pop_Window('met_pummok_inout_detail.asp?stock_goods_code=<%=rs("stock_goods_code")%>&stock_code=<%=rs("stock_code")%>&stock_company=<%=rs("stock_company")%>&stock_name=<%=rs("stock_name")%>&stock_goods_type=<%=rs("stock_goods_type")%>&u_type=<%=""%>','met_pummok_inout_detail_pop','scrollbars=yes,width=1130,height=650')"><%=rs("stock_goods_code")%></a>
                        </td>
                        
                        <td><%=rs("stock_goods_gubun")%>&nbsp;</td>
                        <td><%=rs("stock_goods_name")%>&nbsp;</td>
                        <td><%=rs("stock_goods_standard")%>&nbsp;</td>
                        <td><%=goods_model%>&nbsp;</td>
                        <td><%=rs("stock_goods_grade")%>&nbsp;</td>
                        <td><%=rs("stock_name")%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_last_qty"),0)%>&nbsp;</td>
                        <td class="right">
                  <% if  rs("stock_in_qty") > 0 then    %>                        
                        <a href="#" onClick="pop_Window('met_pummok_stin_sale_detail.asp?stock_goods_code=<%=rs("stock_goods_code")%>&stock_goods_type=<%=rs("stock_goods_type")%>&stock_code=<%=rs("stock_code")%>&stock_name=<%=rs("stock_name")%>','met_pummok_stin_detail_pop','scrollbars=yes,width=930,height=650')"><%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;</a>
                  <%    else   %>
                        <%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;
                  <% end if    %>                        
                        </td>
                        <td class="right">
                  <% if  rs("stock_go_qty") > 0 then    %>
                        <a href="#" onClick="pop_Window('met_stock_sale_chulgo_detail.asp?stock_goods_code=<%=rs("stock_goods_code")%>&stock_goods_type=<%=rs("stock_goods_type")%>&stock_code=<%=rs("stock_code")%>&stock_name=<%=rs("stock_name")%>','met_pummok_chulgo_detail_pop','scrollbars=yes,width=930,height=650')"><%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;</a>
                  <%    else   %>
                        <%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;
                  <% end if    %>
                        </td>
                        <td class="right"><%=formatnumber(rs("stock_jj_qty"),0)%>&nbsp;</td>
                        <td>
                 <% if stock_jj_qty > 0 and user_id = "900002" then    %>  
                        <a href="#" onClick="pop_Window('met_chulgo_sale_add02.asp?view_condi=<%=view_condi%>&stock_code=<%=stock_code%>&stock_goods_type=<%=stock_goods_type%>&stock_goods_code=<%=stock_goods_code%>&u_type=<%=""%>','met_chulgo_sale_add02_pop','scrollbars=yes,width=1230,height=650')">영업</a>
                 <%    else     %>
                        -
                 <%  end if     %>  
                        </td>
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
                    <a href="met_stock_jaego_excel.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&stock=<%=stock%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_stock_out_sale_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&stock=<%=stock%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_stock_out_sale_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&stock=<%=stock%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_stock_out_sale_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&stock=<%=stock%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_stock_out_sale_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&stock=<%=stock%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_stock_out_sale_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&owner_view=<%=owner_view%>&condi=<%=condi%>&stock=<%=stock%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                 <% if user_id = "900002" then    %>  
                    <a href="#" onClick="pop_Window('met_chulgo_sale_add01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&chulgo_id=<%=chulgo_id%>&u_type=<%=""%>','met_chulgo_sale_add01_pop','scrollbars=yes,width=1230,height=650')" class="btnType04">영업출고등록(다품목)</a>
                 <%  end if     %>        
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

