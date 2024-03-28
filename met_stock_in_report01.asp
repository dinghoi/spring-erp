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
met_grade = request.cookies("nkpmg_user")("coo_met_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

view_condi = request("view_condi")
from_date=Request.form("from_date")
to_date=Request.form("to_date")

be_pg = "/met_stock_in_report01.asp"
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
	view_condi = "전체"
	stock = ""
	goods_type = "전체"
	goods_name = ""
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
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

' 권한 체크 본사 마스터 권한이 없는 경우 입고현황을 볼수 없게끔
if met_grade <> "0"  then
   view_condi = ""
   goods_type = ""
   from_date = ""
   to_date = ""
end if

order_Sql = " ORDER BY stin_in_date DESC"

if view_condi = "전체" then
   if goods_type = "전체" then
      where_sql = " WHERE (stin_id = '구매') and (stin_in_date >= '"+from_date+"' and stin_in_date <= '"+to_date+"')"
	  else
	  where_sql = " WHERE (stin_id = '구매') and (stin_goods_type = '"&goods_type&"') and (stin_in_date >= '"+from_date+"' and stin_in_date <= '"+to_date+"')"
   end if
 else
   if goods_type = "전체" then
      where_sql = " WHERE (stin_id = '구매') and (stin_stock_company = '"&view_condi&"') and (stin_in_date >= '"+from_date+"' and stin_in_date <= '"+to_date+"')"
	  else
	  where_sql = " WHERE (stin_id = '구매') and (stin_goods_type = '"&goods_type&"') and (stin_stock_company = '"&view_condi&"') and (stin_in_date >= '"+from_date+"' and stin_in_date <= '"+to_date+"')"
   end if
end if

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (stin_stock_name like '%"&stock&"%') "
end if

if goods_name = "" then
       goods_name_sql = ""
   else
       goods_name_sql = " and (stin_goods_name like '%"&goods_name&"%') "
end if

Sql = "SELECT count(*) FROM met_stin" + where_sql + stock_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_stin " + where_sql + stock_sql + order_sql + " limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 입고 현황 "
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
				return "0 1";
			}

			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});

			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if(document.frm.view_condi.value == ""){
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}
				return true;
			}

			//엑셀 다운로드
			function stockIn_Excel(url){
				var grade = $('#grade').val();
				var view_condi = $('#view_condi').val();
				var stock = $('#stock').val();
				var goods_name = $('#goods_name').val();
				var goods_type = $('#goods_type').val();
				var from_date = $('#from_date').val();
				var to_date = $('#to_date').val();

				//사용 권한 체크
				if(grade !== '0'){
					non_grade();
				}else{
					move(url&"?view_condi="+view_condi+"&stock="+stock+"&goods_name="+goods_name+"&goods_type="+goods_type+"&from_date="+from_date+"&to_date="+to_date);
				}
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/meterials_control_header01.asp" -->
            <!--#include virtual = "/include/meterials_stock_in_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/met_stock_in_report01.asp?ck_sw=<%="n"%>" method="post" name="frm">

				<input type="hidden" name="grade" id="grade" value="<%=met_grade%>" />
				<input type="hidden" name="view_condi" id="view_condi" value="<%=view_condi%>" />
				<input type="hidden" name="stock" id="stock" value="<%=stock%>" />
				<input type="hidden" name="goods_name" id="goods_name" value="<%=goods_name%>" />
				<input type="hidden" name="goods_type" id="goods_type" value="<%=goods_type%>" />
				<input type="hidden" name="from_date" id="from_date" value="<%=from_date%>" />
				<input type="hidden" name="to_date" id="to_date" value="<%=to_date%>" />

				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>검색조건</dt>
                        <dd>
                            <p>
                               <strong>회사&nbsp;</strong>
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
                                <strong>입고창고&nbsp;</strong>
                                   <input name="stock" type="text" id="stock" value="<%=stock%>" style="width:100px; text-align:left; ime-mode:active">
                                </label>
                                <label>
                                <strong>용도구분&nbsp;</strong>
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
								<strong>입고일자&nbsp;</strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								<strong>∼ </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>

				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="3%" >
                      <col width="6%" >
                      <col width="7%" >
                      <col width="7%" >
                      <col width="6%" >
				      <col width="10%" >
                      <col width="10%" >
                      <col width="*" >
                      <col width="12%" >
                      <col width="7%" >
                      <col width="12%" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
				      <tr>
				        <th class="first" scope="col">순번</th>
                        <th scope="col">용도구분</th>
                        <th scope="col">입고일자</th>
                        <th scope="col">입고번호</th>
                        <th scope="col">입고구분</th>
                        <th scope="col">그룹사</th>
                        <th scope="col">사업부</th>
                        <th scope="col">입고창고</th>
                        <th scope="col">입고품목</th>
                        <th scope="col">입고금액</th>
                        <th scope="col">구매거래처</th>
                        <th scope="col">비고</th>
			          </tr>
			        </thead>
				    <tbody>
               <%
					seq = tottal_record - ( page - 1 ) * pgsize
					do until rs.eof
                           stin_in_date = rs("stin_in_date")
						   stin_order_no = rs("stin_order_no")
						   stin_order_seq = rs("stin_order_seq")

						   sql = "select * from met_stin_goods where (stin_date = '"&stin_in_date&"') and (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"')  ORDER BY stin_goods_seq,stin_goods_code ASC"
						   Set Rs_buy=DbConn.Execute(Sql)
						   if Rs_buy.eof or Rs_buy.bof then
								bg_goods_name = ""
							  else
							  	bg_goods_name = Rs_buy("stin_goods_name")
						   end if
						   Rs_buy.close()

			   %>
				      <tr>
				        <td class="first"><%=seq%></td>
                        <td><%=rs("stin_goods_type")%>&nbsp;</td>
                        <td><%=rs("stin_in_date")%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('met_stock_in_detail01.asp?stin_in_date=<%=stin_in_date%>&stin_order_no=<%=stin_order_no%>&stin_order_seq=<%=stin_order_seq%>&u_type=<%=""%>','met_stin_detail_pop','scrollbars=yes,width=930,height=650')"><%=rs("stin_order_no")%>&nbsp;<%=rs("stin_order_seq")%></a>
                        </td>
                        <td><%=rs("stin_id")%>&nbsp;</td>
                        <td><%=rs("stin_buy_company")%>&nbsp;</td>
                        <td><%=rs("stin_buy_saupbu")%>&nbsp;</td>
                        <td><%=rs("stin_stock_name")%>&nbsp;(<%=rs("stin_stock_company")%>)</td>
                        <td><%=bg_goods_name%>&nbsp;외</td>
                        <td class="right"><%=formatnumber(rs("stin_cost"),0)%></td>
                        <td><%=rs("stin_trade_name")%>&nbsp;</td>
                        <td>
                        <a href="#" onClick="pop_Window('met_stock_in_modify01.asp?stin_in_date=<%=stin_in_date%>&stin_order_no=<%=stin_order_no%>&stin_order_seq=<%=stin_order_seq%>&u_type=<%="U"%>','met_stock_in_modify_pop','scrollbars=yes,width=1230,height=650')">수정</a>
                        </td>
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
					<div class="btnleft">
                <%' if met_grade = "0" then	%>
                    <!--<a href="met_stock_in_excel01.asp?view_condi=<%=view_condi%>&stock=<%=stock%>&goods_name=<%=goods_name%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>-->
					<a href="#" onclick="stockIn_Excel('/met_stock_in_excel01.asp');" class="btnType04">엑셀다운로드</a>
                <%' end if	%>
					</div>
                  	</td>
                    <td>
                    <div id="paging">
                        <a href="<%=be_pg%>?page=<%=first_page%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_name=<%=goods_name%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="<%=be_pg%>?page=<%=intstart -1%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_name=<%=goods_name%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   						<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
						<% else %>
                        <a href="<%=be_pg%>?page=<%=i%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_name=<%=goods_name%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
						 <% end if %>
                      <% next %>
   						<% if intend < total_page then %>
                        <a href="<%=be_pg%>?page=<%=intend+1%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_name=<%=goods_name%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="<%=be_pg%>?page=<%=total_page%>&view_condi=<%=view_condi%>&stock=<%=stock%>&goods_name=<%=goods_name%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
						<% end if %>
                    </div>
                    </td>
                    <td width="20%">
					<div class="btnCenter">
                <% if met_grade = "0" then	%>
                        <a href="#" onClick="pop_Window('met_stock_in_add01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&u_type=<%=""%>','met_buy_popup','scrollbars=yes,width=1230,height=650')" class="btnType04">입고 등록</a>

                        <a href="#" onClick="pop_Window('met_stock_in_print01.asp?view_condi=<%=view_condi%>&goods_type=<%=goods_type%>&from_date=<%=from_date%>&to_date=<%=to_date%>','met_stock_in_print_pop','scrollbars=yes,width=1250,height=600')" class="btnType04">입고현황 출력</a>
                <% end if	%>
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

