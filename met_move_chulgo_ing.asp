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

stock_go_man = user_id

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

condi = request("condi")
from_date=Request.form("from_date")
to_date=Request.form("to_date")

be_pg = "met_move_chulgo_ing.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	condi=Request("condi")
	from_date=request("from_date")
    to_date=request("to_date")
  else
	condi=Request.form("condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
End if

If condi = "" Then
	condi = "전체"
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
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

rele_id = "창고이동"
chulgo_ing = "출고의뢰"

'and (rele_sign_yn = 'Y') 테스트후 적용할것

order_Sql = " ORDER BY rele_date,rele_stock,rele_seq DESC"

if condi = "전체" then  
       where_sql = " WHERE (rele_id = '"&rele_id&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"') "
   else 
	   where_sql = " WHERE (chulgo_stock_name = '"&condi&"') and (rele_date >= '"+from_date+"' and rele_date <= '"+to_date+"') "
end if

Sql = "SELECT count(*) FROM met_mv_reg " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from met_mv_reg " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " 창고이동 출고의뢰건 출고등록 "

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
				if (document.frm.condi.value == "") {
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
            <!--#include virtual = "/include/meterials_stock_move_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_move_chulgo_ing.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>회사 검색</dt>
                        <dd>
                            <p>
                               <strong>출고창고 : </strong>
                                    <input name="condi" type="text" id="condi" style="width:120px" value="<%=condi%>"> 
                                    <a href="#" class="btnType03" onClick="pop_Window('met_stock_search.asp?gubun=<%="mving"%>&stock_go_man=<%=stock_go_man%>','stock_search_pop','scrollbars=yes,width=600,height=400')">찾기</a>
                                    <input type="hidden" name="stock_code" value="<%=stock_code%>" ID="Hidden1">
                                    <input type="hidden" name="stock_manager_code" value="<%=stock_manager_code%>" ID="Hidden1">
                                    <input type="hidden" name="stock_manager_name" value="<%=stock_manager_name%>" ID="Hidden1">
                               <label>
								<strong>출고의뢰일(From) : </strong>
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
                <h3 class="stit" style="font-size:12px;">※ 출고창고 찾기를 하시고 검색을 하십시요!</h3>
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableList">
				    <colgroup>
				      <col width="3%" >
                      <col width="6%" >
                      <col width="6%" >
                      <col width="7%" >
				      <col width="8%" >
                      <col width="6%" >
				      <col width="10%" >
				      <col width="8%" >
                      <col width="10%" >
				      <col width="8%" >
				      <col width="6%" >
				      <col width="6%" >
                      <col width="*" >
                      <col width="3%" >
			        </colgroup>
				    <thead>
                      <tr>
				        <th class="first" scope="col">순번</th>
                        <th scope="col">결재요청</th>
                        <th scope="col">의뢰일자</th>
                        <th scope="col">의뢰번호</th>
                        <th scope="col">용도구분</th>
                        <th scope="col">신청자</th>
                        <th scope="col">소속</th>
                        <th scope="col">의뢰창고</th>
                        <th scope="col">의뢰품목</th>
                        <th scope="col">출고요청<br>창고</th>
                        <th scope="col">출고요청<br>일자</th>
                        <th scope="col">출고상태</th>
                        <th scope="col">적요</th>
                        <th scope="col">출고</th>
                      </tr>
			        </thead>
				    <tbody>
                      <%
						seq = tottal_record - ( page - 1 ) * pgsize
						do until rs.eof
						   rele_date = rs("rele_date")
						   rele_stock = rs("rele_stock")
						   rele_seq = rs("rele_seq")
					       
						   if rs("rele_sign_yn") = "Y" then
								sign_view = "결재완료"
							  elseif rs("rele_sign_yn") = "N" then 
								sign_view = "미결재"
							  else
								sign_view = "결재중"
						   end if
						   
						   sql = "select * from met_mv_reg_goods where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')  ORDER BY rl_goods_seq,rl_goods_code ASC"
						   Set Rs_good=DbConn.Execute(Sql)
						   if Rs_good.eof or Rs_good.bof then
								bg_goods_name = ""
							  else
							  	bg_goods_name = Rs_good("rl_goods_name")
						   end if
						   Rs_good.close()
						   
						   rele_no = mid(cstr(rs("rele_date")),3,2) + mid(cstr(rs("rele_date")),6,2) + mid(cstr(rs("rele_date")),9,2) 

					  %>
				      <tr>
				        <td class="first"><%=seq%></td>
                        <td><%=sign_view%></td>
                        <td><%=rs("rele_date")%>&nbsp;</td>
                        <td>
						<a href="#" onClick="pop_Window('met_move_reg_detail.asp?rele_date=<%=rs("rele_date")%>&rele_stock=<%=rs("rele_stock")%>&rele_seq=<%=rs("rele_seq")%>&u_type=<%=""%>','met_move_reg_detail_pop','scrollbars=yes,width=930,height=650')"><%=rele_no%>&nbsp;<%=rs("rele_stock")%><%=rs("rele_seq")%></a>
                        </td>
						<td><%=rs("rele_goods_type")%>&nbsp;</td>
                        <td><%=rs("rele_emp_name")%>&nbsp;</td>
                        <td><%=rs("rele_org_name")%>&nbsp;</td>
                        <td><%=rs("rele_stock_name")%>&nbsp;</td>
                        <td><%=bg_goods_name%>&nbsp;외</td>
                        <td><%=rs("chulgo_stock_name")%>&nbsp;</td>
                        <td><%=rs("chulgo_rele_date")%>&nbsp;</td>
                        <td><%=rs("chulgo_ing")%>&nbsp;</td>
                        <td><%=rs("rele_memo")%>&nbsp;</td>      
                        <td>
        <% if rs("chulgo_ing") = "출고의뢰" or rs("chulgo_ing") = "부분출고" then	%>
                        <a href="#" onClick="pop_Window('met_move_chulgo_add.asp?rele_date=<%=rs("rele_date")%>&rele_stock=<%=rs("rele_stock")%>&rele_seq=<%=rs("rele_seq")%>&stock_go_man=<%=stock_go_man%>&u_type=<%=""%>','met_move_chulgo_add_pop','scrollbars=yes,width=1230,height=650')">등록</a>
		<%   else	%>
								-
		<% end if	%>
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
					<div class="btnCenter">

					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_move_chulgo_ing.asp?page=<%=first_page%>&condi=<%=condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_move_chulgo_ing.asp?page=<%=intstart -1%>&condi=<%=condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
   	        <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_move_chulgo_ing.asp?page=<%=i%>&condi=<%=condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
   	        <% if 	intend < total_page then %>
                        <a href="met_move_chulgo_ing.asp?page=<%=intend+1%>&condi=<%=condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="met_move_chulgo_ing.asp?page=<%=total_page%>&condi=<%=condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
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

