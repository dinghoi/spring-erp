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
from_date=Request.form("from_date")
to_date=Request.form("to_date")

be_pg = "met_stock_in_mg.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
	goods_type=Request("goods_type")
	from_date=request("from_date")
    to_date=request("to_date")
  else
	view_condi=Request.form("view_condi")
	goods_type=Request.form("goods_type")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
End if

If view_condi = "" Then
'	view_condi = "케이원정보통신"
'	goods_type = "상품"
	view_condi = "전체"
	goods_type = "전체"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	ck_sw = "n"
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
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_ing = "3" '발주서를 출력한 경우만 입고등록을 해야 함 and order_ing = '3' sql에 추가할것

order_Sql = " ORDER BY order_company,order_date,order_no,order_seq DESC"
if view_condi = "전체" then
   if goods_type = "전체" then
      where_sql = " WHERE (order_date >= '"+from_date+"' and order_date <= '"+to_date+"')" 
	  else
	  where_sql = " WHERE (order_goods_type = '"&goods_type&"') and (order_date >= '"+from_date+"' and order_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "전체" then
      where_sql = " WHERE (order_company = '"&view_condi&"') and (order_date >= '"+from_date+"' and order_date <= '"+to_date+"')"
	  else
	  where_sql = " WHERE (order_goods_type = '"&goods_type&"') and (order_company = '"&view_condi&"') and (order_date >= '"+from_date+"' and order_date <= '"+to_date+"')"
   end if
end if   
  
Sql = "SELECT count(*) FROM met_order" + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_order " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 발주 -> 입고등록 "

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
				return "1 1";
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
            <!--#include virtual = "/include/meterials_stock_in_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_stock_in_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
                                <label>
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
								<strong>발주일자(From) : </strong>
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
				      <col width="3%" >
                      <col width="6%" >
                      <col width="6%" >
                      <col width="6%" >
				      <col width="10%" >
                      <col width="7%" >
				      <col width="6%" >
                      <col width="5%" >
                      <col width="11%" >
                      <col width="8%" >
				      <col width="6%" >
                      <col width="6%" >
                      <col width="*" >
                      <col width="5%" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">순번</th>
                        <th scope="col">구매번호</th>
                        <th scope="col">용도구분</th>
				        <th scope="col">구매품의일</th>
                        <th scope="col">부서</th>
                        <th scope="col">발주번호</th>
                        <th scope="col">발주일자</th>
                        <th scope="col">발주담당</th>
                        <th scope="col">발주거래처</th>
                        <th scope="col">발주품목</th>
                        <th scope="col">발주금액</th>
                        <th scope="col">입고예정일</th>
                        <th scope="col">입고창고</th>
                        <th scope="col">진행상태</th>
                        <th scope="col">입고</th>
			          </tr>
			        </thead>
				    <tbody>
               <%
						seq = tottal_record - ( page - 1 ) * pgsize
						do until rs.eof
                           order_id = rs("order_id")
						   
						   order_no = rs("order_no")
						   order_seq = rs("order_seq")
						   order_date = rs("order_date")
						   
						   buy_no = rs("order_buy_no")
						   buy_seq = rs("order_buy_seq")
						   buy_date = rs("order_buy_date")
						   
						   order_ing = rs("order_ing")
						   buy_ing = rs("order_ing")
						   buy_ing_gubun = ""
						    
						   if order_id = "1"  then
						          order_id_name = "대기전표"
						      elseif order_id = "2" then
							            order_id_name = "수주전표"
						   end if
						     
						   if buy_ing = "0" then
						         buy_ing_gubun = "구매품의"
						      elseif buy_ing = "1" then
							            buy_ing_gubun = "부분발주"
									 elseif buy_ing = "2" then
							                   buy_ing_gubun = "전체발주"
										    elseif buy_ing = "3" then
							                          buy_ing_gubun = "발주완료"
												   elseif buy_ing = "4" then
							                                 buy_ing_gubun = "입고"
						   end if
						   
						   sql = "select * from met_order_goods where (og_order_no = '"&order_no&"') and (og_order_seq = '"&order_seq&"') and (og_order_date = '"&order_date&"')  ORDER BY og_seq,og_goods_code ASC"
						   Set Rs_buy=DbConn.Execute(Sql)
						   if Rs_buy.eof or Rs_buy.bof then
								bg_goods_name = ""
							  else
							  	bg_goods_name = Rs_buy("og_goods_name")
						   end if
						   Rs_buy.close()
				%>
				      <tr>
				        <td class="first"><%=seq%></td>
                        <td>
        <%  if order_id = "0" then  %>
                        <a href="#" onClick="pop_Window('met_buy_detail_list.asp?buy_no=<%=buy_no%>&buy_date=<%=buy_date%>&buy_seq=<%=buy_seq%>&u_type=<%=""%>','met_buy_detail_pop','scrollbars=yes,width=930,height=650')"><%=buy_no%></a>
        <%       else  %>   
                        <%=order_id_name%>&nbsp;      
        <%  end if %>                                                  
                        </td>
                        <td><%=rs("order_goods_type")%>&nbsp;</td>
                        <td><%=rs("order_buy_date")%>&nbsp;</td>
                        <td><%=rs("order_org_name")%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('met_buy_order_detail.asp?order_no=<%=rs("order_no")%>&order_date=<%=rs("order_date")%>&order_seq=<%=rs("order_seq")%>&u_type=<%=""%>','met_buy_order_detail_pop','scrollbars=yes,width=930,height=650')"><%=rs("order_no")%>&nbsp;<%=rs("order_seq")%></a>
                        </td>
                        <td><%=rs("order_date")%>&nbsp;</td>
                        <td><%=rs("order_emp_name")%>&nbsp;</td>
                        <td><%=rs("order_trade_name")%>&nbsp;</td>
                        <td><%=bg_goods_name%>&nbsp;외</td>
                        <td class="right"><%=formatnumber(rs("order_cost"),0)%></td>
                        <td><%=rs("order_in_date")%>&nbsp;</td>
                        <td><%=rs("order_stock_name")%>&nbsp;</td>
                        <td><%=buy_ing_gubun%>&nbsp;</td>
        <%  if order_ing = "3" or order_ing = "1" or order_ing = "2" then  %>
                        <td>
                        <a href="#" onClick="pop_Window('met_stock_in_add.asp?order_no=<%=rs("order_no")%>&order_seq=<%=rs("order_seq")%>&order_date=<%=rs("order_date")%>&u_type=<%=""%>','met_stock_in_pop','scrollbars=yes,width=1230,height=650')">등록</a>
                        </td>
        <%     else  %>
                        <td>-</td>
        <%  end if %>                                 
			          </tr>
	    <%
							rs.movenext()
							seq = seq -1
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
                        <a href="met_stock_in_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_stock_in_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_stock_in_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_stock_in_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_stock_in_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
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

