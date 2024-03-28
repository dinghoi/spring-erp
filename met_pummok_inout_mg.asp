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
met_grade = request.cookies("nkpmg_user")("coo_cost_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

view_condi = request("view_condi")

be_pg = "met_pummok_inout_mg.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

strNowWeek = WeekDay(curr_date)
Select Case (strNowWeek)
   Case 1
       week = "일요일"
   Case 2
       week = "월요일"
   Case 3
       week = "화요일"
   Case 4
       week = "수요일"
   Case 5
       week = "목요일"
   Case 6
       week = "금요일"
   Case 7
       week = "토요일"
End Select

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
	stock = "케이원정보통신"
	goods_type = "A/S자재"
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
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_jae = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

if stock = "" then
       sql = " delete from met_stock_inout_goods where (stock_company = '"&view_condi&"') and (stock_date >= '"+from_date+"' and stock_date <= '"+to_date+"') and (stock_goods_type = '"&goods_type&"')" 	
   else
       sql = " delete from met_stock_inout_goods where (stock_company = '"&view_condi&"') and (stock_name like '%"&stock&"%') and (stock_date >= '"+from_date+"' and stock_date <= '"+to_date+"') and (stock_goods_type = '"&goods_type&"')" 	
end if
dbconn.execute(sql)

jjj = 0

'구매입고
if stock = "" then
	   sql = "select * from met_stin_goods where (stin_stock_company = '"&view_condi&"') and (stin_date >= '"+from_date+"' and stin_date <= '"+to_date+"') and (stin_goods_type = '"&goods_type&"')" 
   else
       sql = "select * from met_stin_goods where  (stin_stock_company = '"&view_condi&"') and (stin_stock_name like '%"&stock&"%') and (stin_date >= '"+from_date+"' and stin_date <= '"+to_date+"') and (stin_goods_type = '"&goods_type&"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof

    stin_goods_code = rs("stin_goods_code")
	
	sql = "select * from met_goods_code where (goods_code = '"&stin_goods_code&"')"
    Set Rs_good = DbConn.Execute(SQL)
    if not Rs_good.eof then
		goods_grade = Rs_good("goods_grade")
      else
		goods_grade = ""
    end if
    Rs_good.close()
    
	jjj = jjj + 1
    inout_number = right(("00000" + cstr(jjj)),5)

    id_seq = "1"
    sql="insert into met_stock_inout_goods (stock_code,stock_goods_type,stock_goods_code,stock_date,id_seq,inout_number,stock_company,stock_name,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_grade,stock_last_qty,stock_in_qty,stock_go_qty,stock_jj_qty,stock_id,inout_no,inout_seq) values ('"&rs("stin_stock_code")&"','"&rs("stin_goods_type")&"','"&rs("stin_goods_code")&"','"&rs("stin_date")&"','"&id_seq&"','"&inout_number&"','"&rs("stin_stock_company")&"','"&rs("stin_stock_name")&"','"&rs("stin_goods_gubun")&"','"&rs("stin_goods_name")&"','"&rs("stin_standard")&"','"&goods_grade&"',0,'"&rs("stin_qty")&"',0,0,'"&rs("stin_id")&"','"&rs("stin_order_no")&"','"&rs("stin_order_seq")&"')"
	
	dbconn.execute(sql)

	rs.movenext()
loop
rs.close()		

'본사출고 / 고객사출고는제외
if stock = "" then
	   sql = "select * from met_chulgo_goods where (chulgo_stock_company = '"&view_condi&"') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"') and (cg_goods_type = '"&goods_type&"') and (cg_type <> '고객출고')" 
   else
       sql = "select * from met_chulgo_goods where  (chulgo_stock_company = '"&view_condi&"') and (chulgo_stock_name like '%"&stock&"%') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"') and (cg_goods_type = '"&goods_type&"') and (cg_type <> '고객출고')"
end if
Rs.Open Sql, Dbconn, 1
do until rs.eof

    chulgo_date = rs("chulgo_date")
	yymmdd = mid(cstr(chulgo_date),3,2) + mid(cstr(chulgo_date),6,2)  + mid(cstr(chulgo_date),9,2)
	rele_no = yymmdd + rs("chulgo_stock")
    id_seq = "3"
    
	jjj = jjj + 1
    inout_number = right(("00000" + cstr(jjj)),5)
	
    sql="insert into met_stock_inout_goods (stock_code,stock_goods_type,stock_goods_code,stock_date,id_seq,inout_number,stock_company,stock_name,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_grade,stock_last_qty,stock_in_qty,stock_go_qty,stock_jj_qty,stock_id,inout_no,inout_seq,chulgo_return,out_service_no,trade_name,trade_dept,rele_company,rele_saupbu,rele_team,rele_stock_name) values ('"&rs("chulgo_stock")&"','"&rs("cg_goods_type")&"','"&rs("cg_goods_code")&"','"&rs("chulgo_date")&"','"&id_seq&"','"&inout_number&"','"&rs("chulgo_stock_company")&"','"&rs("chulgo_stock_name")&"','"&rs("cg_goods_gubun")&"','"&rs("cg_goods_name")&"','"&rs("cg_standard")&"','"&rs("cg_goods_grade")&"',0,0,'"&rs("cg_qty")&"',0,'"&rs("cg_type")&"','"&rele_no&"','"&rs("chulgo_seq")&"','"&rs("cg_return")&"','"&rs("rl_service_no")&"','"&rs("rl_trade_name")&"','"&rs("rl_trade_dept")&"','"&rs("rl_company")&"','"&rs("rl_saupbu")&"','"&rs("rl_team")&"','"&rs("rl_stock_name")&"')"
	
	dbconn.execute(sql)

	rs.movenext()
loop
rs.close()		

order_Sql = " ORDER BY stock_code,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_goods_code,stock_date,id_seq,inout_no,inout_seq ASC"
where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_date >= '"+from_date+"' and stock_date <= '"+to_date+"')" 

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (stock_name like '%"&stock&"%') "
end if

sql = "select * from met_stock_inout_goods " + where_sql + stock_sql + order_sql 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = stock + " < " + goods_type + " > 입★출고 현황 "

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
			
			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/meterials_control_header01.asp" -->
            <!--#include virtual = "/include/meterials_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_pummok_inout_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
                                <strong>창고 명 : </strong>
                                   <input name="stock" type="text" id="stock" value="<%=stock%>" style="width:100px; text-align:left; ime-mode:active">
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
								<strong>입출고일자(From) : </strong>
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
					<table cellpadding="0" cellspacing="0">
					<tr>
                    	<td>
      					<DIV id="topLine2" style="width:1200px;overflow:hidden;">                     
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="5%" >
                      <col width="5%" >
                      <col width="10%" >
                      <col width="8%" >
                      <col width="8%" >
                      <col width="10%" >
				      <col width="4%" >
                      <col width="10%" >
                      <col width="8%" >
				      <col width="8%" >
				      <col width="6%" >
                      <col width="6%" >
                      <col width="6%" >
                      <col width="*" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">대분류</th>
                        <th scope="col">중분류</th>
                        <th scope="col">규격</th>
				        <th scope="col">품목구분</th>
                        <th scope="col">품목코드</th>
                        <th scope="col">품목명</th>
                        <th scope="col">상태</th>
                        <th scope="col">창고</th>
                        <th scope="col">입출고일자</th>
                        <th scope="col">입출고구분</th>
				        <th scope="col">입고</th>
                        <th scope="col">출고</th>
                        <th scope="col">현재고</th>
                        <th scope="col">비고</th>
			          </tr>
			        </thead>
				   </table>
                   </DIV>
				   </td>
                </tr>
				<tr>
               	  <td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">   
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="5%" >
                      <col width="5%" >
                      <col width="10%" >
                      <col width="8%" >
                      <col width="8%" >
                      <col width="10%" >
				      <col width="4%" >
                      <col width="10%" >
                      <col width="8%" >
				      <col width="8%" >
				      <col width="6%" >
                      <col width="6%" >
                      <col width="6%" >
                      <col width="*" >
			        </colgroup>                                         
				    <tbody>
                      <%
						sum_in_cnt = 0	 
						sum_out_cnt = 0
						sum_jae_cnt = 0
						
						if rs.eof or rs.bof then
							bi_goods_code = ""
							bi_stock = ""
							bi_goods_type = ""
					      else						  
							if isnull(rs("stock_goods_code")) or rs("stock_goods_code") = "" then	
								bi_goods_code = ""
								bi_stock = ""
							    bi_goods_type = ""
							  else
								bi_goods_code = rs("stock_goods_code")
								bi_stock = rs("stock_code")
							    bi_goods_type = rs("stock_goods_type")
							end if
						end if
						
						do until rs.eof

                           if isnull(rs("stock_goods_code")) or rs("stock_goods_code") = "" then
								     goods_goods_code = ""
							  else
							  	     goods_goods_code = rs("stock_goods_code")
						   end if

						   if bi_goods_code <> goods_goods_code then
							     Sql = "SELECT * FROM met_goods_code where goods_code = '"&bi_goods_code&"'"
                                 Set Rs_good = DbConn.Execute(SQL)
							     if not Rs_good.eof then
								      goods_goods_name = Rs_good("goods_name")
							     end if
							     Rs_good.close()
								 
								 sql="select * from met_stock_gmaster where stock_code='"&bi_stock&"' and stock_goods_code='"&bi_goods_code&"' and stock_goods_type='"&bi_goods_type&"'"
	                             set Rs_jae=dbconn.execute(sql)
                                 if not Rs_jae.eof then
								        jj_a_qty = Rs_jae("stock_JJ_qty")
									else
									    jj_a_qty = 0
								 end if
								 Rs_jae.close()
								 
								 sum_jjj_cnt = jj_a_qty - sum_in_cnt + sum_out_cnt
								 sum_jae_cnt = sum_jjj_cnt + sum_in_cnt - sum_out_cnt
					  %>
                                 <tr>
								    <td colspan="8" bgcolor="#EEFFFF" class="first"><%=goods_goods_name%>&nbsp;(<%=bi_goods_code%>)&nbsp;&nbsp;계</td>
							        <td bgcolor="#EEFFFF" >재고</th>
                                    <td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_jjj_cnt,0)%>&nbsp;</td>
                                    <td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_in_cnt,0)%>&nbsp;</td>
							        <td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_out_cnt,0)%>&nbsp;</td>
                                    <td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_jae_cnt,0)%>&nbsp;</td>
							        <td bgcolor="#EEFFFF" >&nbsp;</th>
						         </tr>
                      <%
							     sum_in_cnt = 0	 
						         sum_out_cnt = 0
						         sum_jae_cnt = 0
								 bi_goods_code = goods_goods_code
								 bi_stock = rs("stock_code")
							     bi_goods_type = rs("stock_goods_type")
						   end if
							  
						   stock_goods_type = rs("stock_goods_type")
						   stock_goods_code = rs("stock_goods_code")

                           sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
                           Set Rs_good = DbConn.Execute(SQL)
                           if not Rs_good.eof then
    	                          goods_level1 = Rs_good("goods_level1")
		                          goods_level2 = Rs_good("goods_level2")
                              else
		                          goods_level1 = ""
		                          goods_level2 = ""
                           end if
                           Rs_good.close()
						   
						   sum_in_cnt = sum_in_cnt + int(rs("stock_in_qty"))
						   sum_out_cnt = sum_out_cnt + int(rs("stock_go_qty"))
						   sum_jae_cnt = sum_jae_cnt + 0
						   
					  %>
				      <tr>
				        <td class="first"><%=goods_level1%>&nbsp;</td>
                        <td><%=goods_level2%>&nbsp;</td>
                        <td><%=rs("stock_goods_standard")%>&nbsp;</td>
                        <td><%=rs("stock_goods_gubun")%>&nbsp;</td>
                        <td><%=rs("stock_goods_code")%>&nbsp;</td>
                        <td><%=rs("stock_goods_name")%>&nbsp;</td>
                        <td><%=rs("stock_goods_grade")%>&nbsp;</td>
                        <td><%=rs("stock_name")%>&nbsp;</td>
                        <td><%=rs("stock_date")%>&nbsp;</td>
                        <td><%=rs("stock_id")%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;</td>
                        <td class="right">&nbsp;</td>
                        <td>&nbsp;</td>
			          </tr>
				      <%
							rs.movenext()
						loop
						rs.close()
						
						Sql = "SELECT * FROM met_goods_code where goods_code = '"&bi_goods_code&"'"
                        Set Rs_good = DbConn.Execute(SQL)
						if not Rs_good.eof then
						      goods_goods_name = Rs_good("goods_name")
						end if
				        Rs_good.close()
						
						sql="select * from met_stock_gmaster where stock_code='"&bi_stock&"' and stock_goods_code='"&bi_goods_code&"' and stock_goods_type='"&bi_goods_type&"'"
	                    set Rs_jae=dbconn.execute(sql)
                        if not Rs_jae.eof then
						        jj_a_qty = Rs_jae("stock_JJ_qty")
							else
							    jj_a_qty = 0
						end if
						Rs_jae.close()
								 
						sum_jjj_cnt = jj_a_qty - sum_in_cnt + sum_out_cnt
						sum_jae_cnt = sum_jjj_cnt + sum_in_cnt - sum_out_cnt
					  %>
                        <tr>
						   <td colspan="8" bgcolor="#EEFFFF" class="first"><%=goods_goods_name%>&nbsp;(<%=bi_goods_code%>)&nbsp;&nbsp;계</td>
						   <td bgcolor="#EEFFFF" >재고</th>
                           <td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_jjj_cnt,0)%>&nbsp;</td>
                           <td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_in_cnt,0)%>&nbsp;</td>
						   <td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_out_cnt,0)%>&nbsp;</td>
                           <td bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_jae_cnt,0)%>&nbsp;</td>
						   <td bgcolor="#EEFFFF" >&nbsp;</th>
						</tr>
			        </tbody>
				</table>
                </DIV>
				</td>
                </tr>
				</table>          
                    
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="met_pummok_inout_mg_excel.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&stock=<%=stock%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td width="50%">
                    </td>                    
				    <td width="20%">
					<div class="btnCenter">

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

