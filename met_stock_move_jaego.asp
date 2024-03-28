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

stock_in_man = user_id

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

view_condi = request("view_condi")
from_date=Request.form("from_date")
to_date=Request.form("to_date")

be_pg = "met_stock_move_jaego.asp"
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
	view_condi = "케이원정보통신"
	stock = ""
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

chulgo_id = "창고이동"
chulgo_type = "부분출고"  '출고완료

if goods_type = "전체" then
          where_sql = " WHERE (in_stock_date = '' or in_stock_date = '0000-00-00') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')" 
	  else
	      where_sql = " WHERE (in_stock_date = '' or in_stock_date = '0000-00-00') and (chulgo_goods_type = '"&goods_type&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')" 
end if

if goods_type = "전체" then
       sql = " delete from met_not_enter where (rele_stock_company = '"&view_condi&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')"
   else
       sql = " delete from met_not_enter where (chulgo_goods_type = '"&goods_type&"') and (rele_stock_company = '"&view_condi&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')"
end if
dbconn.execute(sql)

sql = "select * from met_mv_go " + where_sql
rs.Open sql, Dbconn, 1
do until rs.eof
   rele_stock = rs("rele_stock")
   rele_seq = rs("rele_seq")
   rele_date = rs("rele_date")
   
   chulgo_date = rs("chulgo_date")
   chulgo_stock = rs("chulgo_stock")
   chulgo_seq = rs("chulgo_seq")
		
   chulgo_id = rs("chulgo_id")
   chulgo_type = rs("chulgo_type")
   chulgo_goods_type = rs("chulgo_goods_type")
   chulgo_stock_company = rs("chulgo_stock_company")
   chulgo_stock_name = rs("chulgo_stock_name")
   chulgo_emp_no = rs("chulgo_emp_no")
   chulgo_emp_name = rs("chulgo_emp_name")
   chulgo_company = rs("chulgo_company")
   chulgo_bonbu = rs("chulgo_bonbu")
   chulgo_saupbu = rs("chulgo_saupbu")
   chulgo_team = rs("chulgo_team")
   chulgo_org_name = rs("chulgo_org_name")
   chulgo_memo = rs("chulgo_memo")

  
   sql = "select * from met_mv_reg where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')"
   Set Rs_reg = DbConn.Execute(SQL)
   if not Rs_reg.eof then
    	rele_stock_company = Rs_reg("rele_stock_company")
        rele_stock_name = Rs_reg("rele_stock_name")
        rele_emp_no = Rs_reg("rele_emp_no")
        rele_emp_name = Rs_reg("rele_emp_name")
        rele_company = Rs_reg("rele_company")
        rele_bonbu = Rs_reg("rele_bonbu")
        rele_saupbu = Rs_reg("rele_saupbu")
        rele_team = Rs_reg("rele_team")
        rele_org_name = Rs_reg("rele_org_name")
        chulgo_rele_date = Rs_reg("chulgo_rele_date")
      else
		rele_stock_company = ""
        rele_stock_name = ""
        rele_emp_no = ""
        rele_emp_name = ""
        rele_company = ""
        rele_bonbu = ""
        rele_saupbu = ""
        rele_team = ""
        rele_org_name = ""
        chulgo_rele_date = ""
   end if
   Rs_reg.close()
   
   sql = "select * from met_mv_go_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')  ORDER BY cg_goods_seq,cg_goods_code ASC"
   Set Rs_good = DbConn.Execute(SQL)
   do until Rs_good.eof or Rs_good.bof
	    cg_goods_seq = Rs_good("cg_goods_seq")
		cg_goods_code = Rs_good("cg_goods_code")
		cg_goods_gubun = Rs_good("cg_goods_gubun")
		cg_standard = Rs_good("cg_standard")
		cg_goods_name = Rs_good("cg_goods_name")
		cg_goods_grade = Rs_good("cg_goods_grade")
		rl_qty = Rs_good("rl_qty")
		cg_qty = Rs_good("cg_qty")
		cg_type = Rs_good("cg_type")
		
		sql="insert into met_not_enter (chulgo_date,chulgo_stock,chulgo_seq,cg_goods_seq,cg_goods_code,chulgo_goods_type,chulgo_type,chulgo_stock_company,chulgo_stock_name,chulgo_emp_no,chulgo_emp_name,rele_date,rele_stock,rele_seq,rele_stock_company,rele_stock_name,rele_emp_no,rele_emp_name,cg_goods_gubun,cg_standard,cg_goods_name,cg_goods_grade,rl_qty,cg_qty,cg_type) values ('"&chulgo_date&"','"&chulgo_stock&"','"&chulgo_seq&"','"&cg_goods_seq&"','"&cg_goods_code&"','"&chulgo_goods_type&"','"&chulgo_type&"','"&chulgo_stock_company&"','"&chulgo_stock_name&"','"&chulgo_emp_no&"','"&chulgo_emp_name&"','"&rele_date&"','"&rele_stock&"','"&rele_seq&"','"&rele_stock_company&"','"&rele_stock_name&"','"&rele_emp_no&"','"&rele_emp_name&"','"&cg_goods_gubun&"','"&cg_standard&"','"&cg_goods_name&"','"&cg_goods_grade&"','"&rl_qty&"','"&cg_qty&"','"&cg_type&"')"
	
	    dbconn.execute(sql)

        Rs_good.movenext()
   loop
   Rs_good.close()

   rs.movenext()
loop
rs.close()	


order_Sql = " ORDER BY chulgo_date,chulgo_stock,chulgo_seq DESC"

if view_condi = "전체" then
   if goods_type = "전체" then
          where_sql = " WHERE (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')" 
	  else
	      where_sql = " WHERE (chulgo_goods_type = '"&goods_type&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "전체" then
          where_sql = " WHERE (rele_stock_company = '"&view_condi&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')"
	  else
	      where_sql = " WHERE (chulgo_goods_type = '"&goods_type&"') and (rele_stock_company = '"&view_condi&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"')"
   end if
end if   


if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (rele_stock_name like '%"&stock&"%') "
end if

 
Sql = "SELECT count(*) FROM met_not_enter " + where_sql + stock_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_not_enter " + where_sql + stock_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 창고이동중 재고 현황 "

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
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('work1').style.display = 'none';
					document.getElementById('work2').style.display = 'none';
					document.getElementById('acpt1').style.display = '';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('work1').style.display = '';
					document.getElementById('work2').style.display = '';
					document.getElementById('acpt1').style.display = 'none';
				}	
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/meterials_control_header.asp" -->
            <!--#include virtual = "/include/meterials_stock_jaego_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_stock_move_jaego.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>검색</dt>
                        <dd>
                            <p>
                               <strong>신청회사 : </strong>
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
                                <strong>신청창고 명 : </strong>
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
								<strong>신청일자(From) : </strong>
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
				      <col width="7%" >
                      <col width="6%" >
                      <col width="12%" >
				      <col width="8%" >
                      <col width="10%" >
				      <col width="7%" >
                      <col width="12%" >
                      <col width="6%" >
				      <col width="*" >
				      <col width="6%" >
			        </colgroup>
				    <thead>
                      <tr>
				        <th class="first" scope="col">순번</th>
                        <th scope="col">용도구분</th>
                        <th scope="col">출고일자</th>
                        <th scope="col">출고번호</th>
                        <th scope="col">출고상태</th>
                        <th scope="col">출고창고</th>
                        <th scope="col">출고담당</th>
                        <th scope="col">출고품목</th>
                        <th scope="col">신청일(No)</th>
                        <th scope="col">신청창고</th>
                        <th scope="col">신청담당</th>
                        <th scope="col">의뢰수량</th>
                        <th scope="col">출고수량</th>
                      </tr>
			        </thead>
				    <tbody>
                  <%
						seq = tottal_record - ( page - 1 ) * pgsize
						do until rs.eof
						   chulgo_date = rs("chulgo_date")
						   chulgo_stock = rs("chulgo_stock")
						   chulgo_seq = rs("chulgo_seq")
						   
						   rele_date = rs("rele_date")
						   rele_stock = rs("rele_stock")
						   rele_seq = rs("rele_seq")
					       
						   chulgo_no = mid(cstr(rs("chulgo_date")),3,2) + mid(cstr(rs("chulgo_date")),6,2) + mid(cstr(rs("chulgo_date")),9,2) 
						   rele_no = mid(cstr(rs("rele_date")),3,2) + mid(cstr(rs("rele_date")),6,2) + mid(cstr(rs("rele_date")),9,2) 

				  %>
				      <tr>
				        <td class="first"><%=seq%></td>
                        <td><%=rs("chulgo_goods_type")%>&nbsp;</td>
                        <td><%=rs("chulgo_date")%>&nbsp;</td>
                        <td>
						<a href="#" onClick="pop_Window('met_move_chulgo_detail.asp?chulgo_date=<%=rs("chulgo_date")%>&chulgo_stock=<%=rs("chulgo_stock")%>&chulgo_seq=<%=rs("chulgo_seq")%>&u_type=<%=""%>','met_move_chulgo_detail_pop','scrollbars=yes,width=930,height=650')"><%=chulgo_no%>&nbsp;<%=rs("chulgo_stock")%><%=rs("chulgo_seq")%></a>
                        </td>
                        <td><%=rs("chulgo_type")%>&nbsp;</td>
                        <td><%=rs("chulgo_stock_name")%>(<%=rs("chulgo_stock")%>)&nbsp;</td>
                        <td><%=rs("chulgo_emp_name")%>(<%=rs("chulgo_emp_no")%>)&nbsp;</td>
                        <td><%=rs("cg_goods_name")%>&nbsp;</td>
                        <td>
						<a href="#" onClick="pop_Window('met_move_reg_detail.asp?rele_date=<%=rs("rele_date")%>&rele_stock=<%=rs("rele_stock")%>&rele_seq=<%=rs("rele_seq")%>&u_type=<%=""%>','met_move_reg_detail_pop','scrollbars=yes,width=930,height=650')"><%=rele_no%>&nbsp;<%=rs("rele_stock")%><%=rs("rele_seq")%></a>
                        </td>
                        <td><%=rs("rele_stock_name")%>&nbsp;</td>
                        <td><%=rs("rele_emp_name")%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("rl_qty"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("cg_qty"),0)%>&nbsp;</td>
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
                    <a href="met_stock_move_jaego_excel.asp?view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_stock_move_jaego.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_stock_move_jaego.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_stock_move_jaego.asp?page=<%=i%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_stock_move_jaego.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_stock_move_jaego.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
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

