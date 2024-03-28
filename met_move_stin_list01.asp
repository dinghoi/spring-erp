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
from_date=Request.form("from_date")
to_date=Request.form("to_date")

be_pg = "met_move_stin_list01.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
	stock = request("stock")
	goods_type=Request("goods_type")
	from_date=request("from_date")
    to_date=request("to_date")
  else
	view_condi=Request.form("view_condi")
	stock = request.form("stock")
	goods_type=Request.form("goods_type")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
End if

If view_condi = "" Then
'	view_condi = "케이원정보통신"
	view_condi = emp_company
	stock = user_name
'	goods_type = "상품"
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
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

'mvin_id = "본사출고"

order_Sql = " ORDER BY mvin_in_date DESC"

if view_condi = "전체" then
   if goods_type = "전체" then
          where_sql = " WHERE (mvin_in_date >= '"&from_date&"' and mvin_in_date <= '"&to_date&"')" 
	  else
	      where_sql = " WHERE (mvin_goods_type = '"&goods_type&"') and (mvin_in_date >= '"&from_date&"' and mvin_in_date <= '"&to_date&"')" 
   end if
 else  
   if goods_type = "전체" then
          where_sql = " WHERE (mvin_stock_company = '"&view_condi&"') and (mvin_in_date >= '"&from_date&"' and mvin_in_date <= '"&to_date&"')"
	  else
	      where_sql = " WHERE (mvin_goods_type = '"&goods_type&"') and (mvin_stock_company = '"&view_condi&"') and (mvin_in_date >= '"&from_date&"' and mvin_in_date <= '"&to_date&"')"
   end if
end if   

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (mvin_stock_name like '%"&stock&"%') "
end if
 
Sql = "SELECT count(*) FROM met_mv_in " + where_sql + stock_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_mv_in " + where_sql + stock_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 창고(CE 등) 입고 현황 "

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
				return "3 1";
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
			<!--#include virtual = "/include/meterials_control_header01.asp" -->
            <!--#include virtual = "/include/meterials_stock_ceout_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_move_stin_list01.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
                                <strong>입고창고 명 : </strong>
                                   <input name="stock" type="text" id="stock" value="<%=stock%>" style="width:100px; text-align:left">
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
								<strong>입고일자(From) : </strong>
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
                      <col width="8%" >
                      <col width="8%" >
				      <col width="8%" >
                      <col width="8%" >
                      <col width="12%" >
                      <col width="8%" >
                      <col width="8%" >
				      <col width="10%" >
                      <col width="12%" >
				      <col width="*" >
			        </colgroup>
				    <thead>
                      <tr>
				        <th class="first" scope="col">순번</th>
                        <th scope="col">용도구분</th>
                        <th scope="col">출고일자</th>
                        <th scope="col">출고번호</th>
                        <th scope="col">출고구분</th>
                        <th scope="col">출고창고</th>
                        <th scope="col">입고일자</th>
                        <th scope="col">입고번호</th>
                        <th scope="col">입고창고</th>
                        <th scope="col">품목..</th>
                        <th scope="col">적요</th>
                      </tr>
			        </thead>
				    <tbody>
                  <%
						seq = tottal_record - ( page - 1 ) * pgsize
						do until rs.eof
						   mvin_in_date = rs("mvin_in_date")
						   mvin_in_stock = rs("mvin_in_stock")
						   mvin_in_seq = rs("mvin_in_seq")
						   
						   chulgo_date = rs("chulgo_date")
						   chulgo_stock = rs("chulgo_stock")
						   chulgo_seq = rs("chulgo_seq")
						   
						   sql = "select * from met_mv_in_goods where (mvin_in_date = '"&mvin_in_date&"') and (mvin_in_stock = '"&mvin_in_stock&"') and (mvin_in_seq = '"&mvin_in_seq&"')  ORDER BY in_goods_seq,in_goods_code ASC"
						   Set Rs_good=DbConn.Execute(Sql)
						   if Rs_good.eof or Rs_good.bof then
								bg_goods_name = ""
							  else
							  	bg_goods_name = Rs_good("in_goods_name")
						   end if
						   Rs_good.close()
						   
						   mvin_no = mid(cstr(rs("mvin_in_date")),3,2) + mid(cstr(rs("mvin_in_date")),6,2) + mid(cstr(rs("mvin_in_date")),9,2) 
						   chulgo_no = mid(cstr(rs("chulgo_date")),3,2) + mid(cstr(rs("chulgo_date")),6,2) + mid(cstr(rs("chulgo_date")),9,2) 
						   rele_no = mid(cstr(rs("rele_date")),3,2) + mid(cstr(rs("rele_date")),6,2) + mid(cstr(rs("rele_date")),9,2) 
						   
				  %>
				      <tr>
				        <td class="first"><%=seq%></td>
                        <td><%=rs("mvin_goods_type")%>&nbsp;</td>
                        <td><%=chulgo_date%>&nbsp;</td>
                        <td>
						<a href="#" onClick="pop_Window('met_move_chulgo_detail01.asp?chulgo_date=<%=rs("chulgo_date")%>&chulgo_stock=<%=rs("chulgo_stock")%>&chulgo_seq=<%=rs("chulgo_seq")%>&u_type=<%=""%>','met_move_chulgo_detail_pop','scrollbars=yes,width=930,height=650')"><%=rs("rele_no")%>&nbsp;<%=rs("rele_seq")%></a>
                        </td>
                        <td><%=rs("mvin_id")%>&nbsp;</td>
                        <td><%=rs("chulgo_stock_name")%>(<%=rs("chulgo_stock")%>)&nbsp;</td>
                        <td><%=rs("mvin_in_date")%>&nbsp;</td>
                        <td><%=mvin_no%><%=rs("mvin_in_stock")%>&nbsp;<%=rs("mvin_in_seq")%></td>
                        <td><%=rs("mvin_stock_name")%>&nbsp;</td>
                        <td><%=bg_goods_name%>&nbsp;외</td>
                        <td class="left"><%=rs("chulgo_memo")%>&nbsp;</td>
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
				    <td width="20%">
					<div class="btnCenter">
                    <a href="met_move_stin_excel.asp?view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_move_stin_list01.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_move_stin_list01.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_move_stin_list01.asp?page=<%=i%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_move_stin_list01.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_move_stin_list01.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
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

