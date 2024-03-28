<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

' 창고별 품목별 기간별 입출고 현황

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
met_grade = request.cookies("nkpmg_user")("coo_cost_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

view_condi = request("view_condi")

be_pg = "met_stock_pummok_inout01.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
	stock = request("stock")
	goods_type=Request("goods_type")
	goods_gubun=Request("goods_gubun")
	from_date=request("from_date")
    to_date=request("to_date")
  else
	view_condi=Request.form("view_condi")
	stock = request.form("stock")
	goods_type=Request.form("goods_type")
	goods_gubun=Request.form("goods_gubun")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
End if

If view_condi = "" Then
	view_condi = "케이원정보통신"
	stock = ""
	goods_type = "A/S자재"
	goods_gubun = ""
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	ck_sw = "n"
End If

'st_date = cstr(mid(curr_date,1,4)) + "-" + "01" + "-" + "01"
st_date = "2015" + "-" + "01" + "-" + "01"  '자재관리 초기재고 등록 기준..추후 매년 이월을 하면 매년 1월1일 기준으로 바꿀것
be_from_date = cstr(mid(from_date,1,4)) + cstr(mid(from_date,6,2)) + cstr(mid(from_date,9,2))
be_to_date = cstr(mid(to_date,1,4)) + cstr(mid(to_date,6,2)) + cstr(mid(to_date,9,2))

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
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_out = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY stock_company,stock_goods_type,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_code ASC"

if goods_type = "전체" then 
         where_sql = " WHERE (a.stock_company = '"&view_condi&"')" 
   else
         where_sql = " WHERE (a.stock_company = '"&view_condi&"') and (a.stock_goods_type = '"&goods_type&"')" 
end if

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (a.stock_name like '%"&stock&"%') "
end if

if goods_gubun = "" then
       gubun_sql = ""
   else
       gubun_sql = " and (a.stock_goods_gubun like '%"&goods_gubun&"%') "
end if

' 입출고 및 재고수량이 없는것
cnt_sql = " and (a.stock_in_qty > 0 or a.stock_go_qty > 0 or a.stock_JJ_qty > 0) "

yes_goods_sql = " and (a.stock_goods_code = b.goods_code AND b.goods_used_sw = 'Y') "

Sql = "SELECT count(*) FROM met_stock_gmaster a, met_goods_code b " + where_sql + stock_sql + gubun_sql + cnt_sql + yes_goods_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_stock_gmaster a, met_goods_code b " + where_sql + stock_sql + gubun_sql + cnt_sql + yes_goods_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = stock + " < " + goods_type + " > 품목/기간별 입출고 현황 "

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
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/meterials_control_header01.asp" -->
            <!--#include virtual = "/include/meterials_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_stock_pummok_inout01.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<strong>회사</strong>
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
                                <strong>창고명</strong>
                                   <input name="stock" type="text" id="stock" value="<%=stock%>" style="width:100px; text-align:left; ime-mode:active">
                                </label>
                                <strong>용도구분</strong>
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
                                <strong>품목구분</strong>
                              <%
								Sql="select * from met_etc_code where etc_type = '04' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>
                                <label>
                                <select name="goods_gubun" id="goods_gubun" type="text" style="width:90px">
                                    <option value="">선택</option>
                			  <% 
								do until Rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("etc_name")%>' <%If goods_gubun = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			  <%
									Rs_etc.movenext()  
								loop 
								Rs_etc.Close()
							  %>
            					</select>
                                 </label>                                
                                <label>
								<strong>입출고일자(From)</strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>∼To</strong>
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
				      <col width="7%" >
                      <col width="8%" >
                      <col width="8%" >
                      <col width="18%" >
                      <col width="16%" >
                      <col width="10%" >
                      <col width="4%" >
                      <col width="*" >
				      <col width="6%" >
                      <col width="5%" >
				      <col width="5%" >
				      <col width="5%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">용도구분</th>
                        <th scope="col">코드</th>
				        <th scope="col">품목구분</th>
                        <th scope="col">품목명</th>
                        <th scope="col">규격</th>
                        <th scope="col">Part_No.</th>
                        <th scope="col">상태</th>
                        <th scope="col">창고</th>
                        <th scope="col">재고수량</th>
				        <th scope="col">입고</th>
                        <th scope="col">출고</th>
                        <th scope="col">현재고</th>
			          </tr>
			        </thead>
				    <tbody>
                      <%
						do until rs.eof
                           stock_level = rs("stock_level")
						   stock_company = rs("stock_company")
						   stock_code = rs("stock_code")
						   stock_goods_type = rs("stock_goods_type")
						   stock_goods_code = rs("stock_goods_code")
						   stock_last_qty = rs("stock_last_qty")
						   stock_jj_qty = rs("stock_JJ_qty")
						   
						   sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
                           Set Rs_good = DbConn.Execute(SQL)
                           if not Rs_good.eof then
    	                         part_number = Rs_good("part_number")
                              else
		                         part_number = ""
                           end if
                           Rs_good.close()

                           be_in_cnt = 0	 
						   be_out_cnt = 0
						   be_jae_cnt = 0
						   
						   sum_in_cnt = 0	 
						   sum_out_cnt = 0
						   sum_jae_cnt = 0
'입고찾기- 본사구매입고	/ 직접 영업및ce쪽으로 입고시킨것						   
						   sql = "select * from met_stin_goods where (stin_stock_company = '"&stock_company&"') and (stin_stock_code = '"&stock_code&"') and (stin_goods_type = '"&stock_goods_type&"') and (stin_goods_code = '"&stock_goods_code&"') and (stin_date >= '"+st_date+"' and stin_date <= '"+to_date+"') ORDER BY stin_date ASC"
			   
'			               response.write sql
			   
	                       Rs_in.Open Sql, Dbconn, 1
                           while not Rs_in.eof
							  
							  be_stin_date = cstr(mid(Rs_in("stin_date"),1,4)) + cstr(mid(Rs_in("stin_date"),6,2)) + cstr(mid(Rs_in("stin_date"),9,2))
							  if be_stin_date < be_from_date then
							         be_in_cnt = be_in_cnt + int(Rs_in("stin_qty"))
							  end if
							  
							  if be_stin_date >= be_from_date and be_stin_date <= be_to_date then
							         sum_in_cnt = sum_in_cnt + int(Rs_in("stin_qty"))
							  end if
							  
						   	  Rs_in.movenext()
                           Wend
                           Rs_in.close()	
						   
'입고찾기- 본사 또는 개인이 ce쪽으로 출고한것(창고이동 입고건)							   
						   sql = "select * from met_mv_in_goods where (mvin_in_stock = '"&stock_code&"') and (in_goods_type = '"&stock_goods_type&"') and (in_goods_code = '"&stock_goods_code&"') and (mvin_in_date >= '"+st_date+"' and mvin_in_date <= '"+to_date+"') ORDER BY mvin_in_date ASC"
						   
						   Rs_in.Open Sql, Dbconn, 1
                           while not Rs_in.eof
							  
							  be_stin_date = cstr(mid(Rs_in("mvin_in_date"),1,4)) + cstr(mid(Rs_in("mvin_in_date"),6,2)) + cstr(mid(Rs_in("mvin_in_date"),9,2))
							  if be_stin_date < be_from_date then
							         be_in_cnt = be_in_cnt + int(Rs_in("in_qty"))
							  end if
							  
							  if be_stin_date >= be_from_date and be_stin_date <= be_to_date then
							         sum_in_cnt = sum_in_cnt + int(Rs_in("in_qty"))
							  end if
							  
						   	  Rs_in.movenext()
                           Wend
                           Rs_in.close()	

'출고찾기						   
						   sql = "select * from met_chulgo_goods where  (chulgo_stock_company = '"&stock_company&"') and (chulgo_stock = '"&stock_code&"') and (cg_goods_type = '"&stock_goods_type&"') and (cg_goods_code = '"&stock_goods_code&"') and (chulgo_date >= '"&st_date&"' and chulgo_date <= '"&to_date&"') and (cg_type <> '고객출고') ORDER BY chulgo_date ASC"
						   
						   Rs_out.Open Sql, Dbconn, 1
                           while not Rs_out.eof
						      
							  be_chulgo_date = cstr(mid(Rs_out("chulgo_date"),1,4)) + cstr(mid(Rs_out("chulgo_date"),6,2)) + cstr(mid(Rs_out("chulgo_date"),9,2))
							  
							  if be_chulgo_date < be_from_date then
							         be_out_cnt = be_out_cnt + int(Rs_out("cg_qty"))
							  end if
							  
							  if be_chulgo_date >= be_from_date and be_chulgo_date <= be_to_date then
							         sum_out_cnt = sum_out_cnt + int(Rs_out("cg_qty"))
							  end if

						   	  Rs_out.movenext()
                           Wend
                           Rs_out.close()
						   
						   be_jae_cnt = stock_last_qty + be_in_cnt - be_out_cnt
						   sum_jae_cnt = be_jae_cnt + sum_in_cnt - sum_out_cnt

					  %>
				      <tr>
				        <td class="first"><%=rs("stock_goods_type")%>&nbsp;</td>
                        <td><%=rs("stock_goods_code")%>&nbsp;</td>
                        <td><%=rs("stock_goods_gubun")%>&nbsp;</td>
                        <td><%=rs("stock_goods_name")%>&nbsp;</td>
                        <td><%=rs("stock_goods_standard")%>&nbsp;</td>
                        <td><%=part_number%>&nbsp;</td>
                        <td><%=rs("stock_goods_grade")%>&nbsp;</td>
                        <td><%=rs("stock_name")%>&nbsp;</td>
                        <td class="right"><%=formatnumber(be_jae_cnt,0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(sum_in_cnt,0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(sum_out_cnt,0)%>&nbsp;</td>
                        <td class="right"><%=formatnumber(sum_jae_cnt,0)%>&nbsp;</td>
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
                    <a href="met_stock_pummok_inout01_excel.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&goods_gubun=<%=goods_gubun%>&from_date=<%=from_date%>&to_date=<%=to_date%>&stock=<%=stock%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_stock_pummok_inout01.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&goods_gubun=<%=goods_gubun%>&from_date=<%=from_date%>&to_date=<%=to_date%>&stock=<%=stock%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_stock_pummok_inout01.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&goods_gubun=<%=goods_gubun%>&from_date=<%=from_date%>&to_date=<%=to_date%>&stock=<%=stock%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_stock_pummok_inout01.asp?page=<%=i%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&goods_gubun=<%=goods_gubun%>&from_date=<%=from_date%>&to_date=<%=to_date%>&stock=<%=stock%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_stock_pummok_inout01.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&goods_gubun=<%=goods_gubun%>&from_date=<%=from_date%>&to_date=<%=to_date%>&stock=<%=stock%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_stock_pummok_inout01.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&goods_gubun=<%=goods_gubun%>&from_date=<%=from_date%>&to_date=<%=to_date%>&stock=<%=stock%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="met_stock_pummok_inout01_excel2.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&goods_gubun=<%=goods_gubun%>&from_date=<%=from_date%>&to_date=<%=to_date%>&stock=<%=stock%>" class="btnType04">입출고일자별 내역</a>
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

