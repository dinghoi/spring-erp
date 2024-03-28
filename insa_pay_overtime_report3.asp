<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

be_pg = "insa_pay_overtime_report2.asp"

from_date=Request.form("from_date")
to_date=Request.form("to_date")

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
	pmg_yymm=Request.form("pmg_yymm")
  else
	view_condi = request("view_condi")
	from_date=request("from_date")
    to_date=request("to_date")
	pmg_yymm=request("pmg_yymm")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	'pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
	
	overtime_count = 0	
	sum_overtime_pay = 0	
end if

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next


pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from overtime where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N' ORDER BY emp_company,team,org_name,work_date,mg_ce_id ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    overtime_count = overtime_count + 1
    sum_overtime_pay = sum_overtime_pay + int(rs("overtime_amt"))
	rs.movenext()
loop
rs.close()

Sql = "select * from overtime where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N' ORDER BY emp_company,team,org_name,mg_ce_id,work_date ASC"

Rs.Open Sql, Dbconn, 1


sql = " delete from pay_overtime_cost " 	
dbconn.execute(sql)

i = 0

do until rs.eof
    
	i = i + 1
	max_seq = "000000" + cstr(i)
	ot_seq = right(max_seq,6)
	
    sql="insert into pay_overtime_cost (work_date,mg_ce_id,ot_seq,emp_name,cost_detail,emp_company,bonbu,saupbu,team,org_name,reside_place,company,dept,from_time,to_time,overtime_amt,sign_no,acpt_no,work_gubun,work_memo) values ('"&rs("work_date")&"','"&rs("mg_ce_id")&"','"&ot_seq&"','"&rs("user_name")&"','"&rs("cost_detail")&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','"&rs("reside_place")&"','"&rs("company")&"','"&rs("dept")&"','"&rs("from_time")&"','"&rs("to_time")&"','"&rs("overtime_amt")&"','"&rs("sign_no")&"','"&rs("acpt_no")&"','"&rs("work_gubun")&"','"&rs("work_memo")&"')"
	
	dbconn.execute(sql)

	rs.movenext()
loop
rs.close()		

'일반경비중 잡비
sql = "select * from general_cost where (emp_company = '"&view_condi&"') and (cost_reg = '0') and (tax_bill_yn <> 'Y' or isnull(tax_bill_yn)) and (slip_gubun = '비용') and (account = '잡비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
order_sql = " ORDER BY emp_company,team,org_name,emp_no,slip_date ASC"

Rs.Open Sql, Dbconn, 1
do until rs.eof
    overtime_count = overtime_count + 1
    sum_overtime_pay = sum_overtime_pay + int(rs("cost"))
	rs.movenext()
loop
rs.close()


sql = "select * from general_cost where (emp_company = '"&view_condi&"') and (cost_reg = '0') and (tax_bill_yn <> 'Y' or isnull(tax_bill_yn)) and (slip_gubun = '비용') and (account = '잡비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
order_sql = " ORDER BY emp_company,team,org_name,emp_no,slip_date ASC"

Rs.Open Sql, Dbconn, 1

do until rs.eof
    
	i = i + 1
	max_seq = "000000" + cstr(i)
	ot_seq = right(max_seq,6)
	
    sql="insert into pay_overtime_cost (work_date,mg_ce_id,ot_seq,emp_name,cost_detail,emp_company,bonbu,saupbu,team,org_name,reside_place,company,dept,from_time,to_time,overtime_amt,sign_no,acpt_no,work_gubun,work_memo) values ('"&rs("slip_date")&"','"&rs("emp_no")&"','"&ot_seq&"','"&rs("emp_name")&"','"&rs("account")&"','"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','"&rs("reside_place")&"','"&rs("company")&"','"&rs("customer")&"','','','"&rs("cost")&"','"&rs("sign_no")&"',0,'"&rs("account_item")&"','"&rs("slip_memo")&"')"
	
	dbconn.execute(sql)

	rs.movenext()
loop
rs.close()		


Sql = "select * from pay_overtime_cost where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' ORDER BY emp_company,team,org_name,mg_ce_id,work_date ASC"

Rs.Open Sql, Dbconn, 1

title_line = ""+ view_condi +" - 야·특근 수당(일반경비 잡비포함) "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
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
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
			
			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_overtime_report3.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
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
								<strong>야·특근기간(From) : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>(To) : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <label>
								<strong>귀속년월 : </strong>
                                <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
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
							<col width="14%" >
							<col width="6%" >
							<col width="13%" >
							<col width="8%" >
							<col width="13%" >
                            <col width="11%" >
                            <col width="5%" >
                            <col width="5%" >
							<col width="6%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">소속</th>
								<th scope="col">구분</th>
								<th scope="col">작업일시</th>
								<th scope="col">고객사 명</th>
								<th scope="col">지점명</th>
								<th scope="col">작업자</th>
                                <th scope="col">Sing No.</th>
                                <th scope="col">금액</th>
                                <th scope="col">AS No.</th>
								<th scope="col">작업내용</th>
							</tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
                        <colgroup>
							<col width="14%" >
							<col width="6%" >
							<col width="13%" >
							<col width="8%" >
							<col width="13%" >
                            <col width="11%" >
                            <col width="5%" >
                            <col width="5%" >
							<col width="6%" >
							<col width="*" >
						</colgroup>                                                
						<tbody>
						<%
						     sum_overtime_cnt = 0	 
						     sum_overtime_cost = 0
							 
							 tot_overtime_cnt = 0	 
						     tot_overtime_cost = 0
							 
						if rs.eof or rs.bof then
							bi_team = ""
							bi_ce = ""
					      else						  
							if isnull(rs("team")) or rs("team") = "" then	
								bi_team = ""
							  else
								bi_team = rs("team")
							end if
							if isnull(rs("mg_ce_id")) or rs("mg_ce_id") = "" then	
								bi_mg_ce_id = ""
							  else
								bi_mg_ce_id = rs("mg_ce_id")
							end if
						end if

						do until rs.eof

                              if isnull(rs("team")) or rs("team") = "" then
								     emp_team = ""
							     else
							  	     emp_team = rs("team")
							  end if
							  if isnull(rs("mg_ce_id")) or rs("mg_ce_id") = "" then
								     mg_ce_id = ""
							     else
							  	     mg_ce_id = rs("mg_ce_id")
							  end if
							
							  if bi_mg_ce_id <> mg_ce_id then
							     emp_no = bi_mg_ce_id
							     Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                                 Set rs_emp = DbConn.Execute(SQL)
							     if not Rs_emp.eof then
								      ce_name = rs_emp("emp_name")
							     end if
							     rs_emp.close()
					  %>
                                 <tr>
								    <td colspan="4" bgcolor="#EEFFFF" class="first"><%=ce_name%>&nbsp;&nbsp;&nbsp;계</td>
							        <td colspan="2" bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_overtime_cnt,0)%>&nbsp;건</td>
							        <td colspan="2" bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_overtime_cost,0)%>&nbsp;원</td>
							        <td colspan="2" bgcolor="#EEFFFF" >&nbsp;</th>
						         </tr>
                      <%
							     sum_overtime_cnt = 0	 
								 sum_overtime_cost = 0
								 bi_mg_ce_id = mg_ce_id
							  end if
							  
							  if bi_team <> emp_team then
					  %>
                                 <tr>
								    <td colspan="4" bgcolor="#EEFFFF" class="first"><%=bi_team%>&nbsp;&nbsp;&nbsp;계</td>
							        <td colspan="2" bgcolor="#EEFFFF" class="right"><%=formatnumber(tot_overtime_cnt,0)%>&nbsp;건</td>
							        <td colspan="2" bgcolor="#EEFFFF" class="right"><%=formatnumber(tot_overtime_cost,0)%>&nbsp;원</td>
							        <td colspan="2" bgcolor="#EEFFFF" >&nbsp;</th>
						         </tr>                      
                      
                      <%    
							     tot_overtime_cnt = 0	 
								 tot_overtime_cost = 0
								 bi_team = emp_team
							  end if
							  					                    							  
							  emp_no = rs("mg_ce_id")
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
								   emp_end_date = rs_emp("emp_end_date")
							  end if
							  rs_emp.close()
                              
							  if isNull(emp_end_date) or emp_end_date = "1900-01-01" or emp_end_date = "0000-00-00" then
							          emp_end = ""
								 else 
								      emp_end = "퇴직"
							  end if
							  
							  sum_overtime_cnt = sum_overtime_cnt + 1	 
							  sum_overtime_cost = sum_overtime_cost + int(rs("overtime_amt"))
							  
							  tot_overtime_cnt = tot_overtime_cnt + 1	 
							  tot_overtime_cost = tot_overtime_cost + int(rs("overtime_amt"))

	           			%>
							<tr>
								<td class="left"><%=rs("team")%>-<%=rs("org_name")%></td>
                                <td class="left"><%=rs("cost_detail")%></td>
                                <td class="left"><%=rs("work_date")%>&nbsp;<%=mid(rs("from_time"),1,2)%>:<%=mid(rs("from_time"),3,2)%>∼<%=mid(rs("to_time"),1,2)%>:<%=mid(rs("to_time"),3,2)%></td>
                                <td class="left"><%=rs("company")%></td>
                                <td class="left"><%=rs("dept")%></td>
                                <td><%=emp_name%>(<%=rs("mg_ce_id")%>)<%=emp_end%></td>
                                <td><%=rs("sign_no")%></td>
                                <td class="right"><%=formatnumber(rs("overtime_amt"),0)%></td>
                                <td><%=rs("acpt_no")%></td>
                                <td class="left"><%=rs("work_gubun")%>-<%=rs("work_memo")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()

						         emp_no = bi_mg_ce_id
							     Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                                 Set rs_emp = DbConn.Execute(SQL)
							     if not Rs_emp.eof then
								      ce_name = rs_emp("emp_name")
							     end if
							     rs_emp.close()
						
						%>
                            <tr>
								<td colspan="4" bgcolor="#EEFFFF" class="first"><%=ce_name%>&nbsp;&nbsp;&nbsp;계</td>
							    <td colspan="2" bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_overtime_cnt,0)%>&nbsp;건</td>
							    <td colspan="2" bgcolor="#EEFFFF" class="right"><%=formatnumber(sum_overtime_cost,0)%>&nbsp;원</td>
							    <td colspan="2" bgcolor="#EEFFFF" >&nbsp;</th>
						    </tr>
                            
                            <tr>
								<td colspan="4" bgcolor="#EEFFFF" class="first"><%=bi_team%>&nbsp;&nbsp;&nbsp;계</td>
							    <td colspan="2" bgcolor="#EEFFFF" class="right"><%=formatnumber(tot_overtime_cnt,0)%>&nbsp;건</td>
							    <td colspan="2" bgcolor="#EEFFFF" class="right"><%=formatnumber(tot_overtime_cost,0)%>&nbsp;원</td>
							    <td colspan="2" bgcolor="#EEFFFF" >&nbsp;</th>
						    </tr>  

                            <tr>
								<th colspan="4" class="first">합&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;계</th>
							    <th colspan="2" class="right"><%=formatnumber(overtime_count,0)%>&nbsp;건</th>
							    <th colspan="2" class="right"><%=formatnumber(sum_overtime_pay,0)%>&nbsp;원</th>
							    <th colspan="2">&nbsp;</th>
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
                    <a href="insa_excel_pay_overtime3.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>" class="btnType04">엑셀다운로드</a>
                    <a href="insa_pay_overtime_excel3.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>" class="btnType04">조직/개인별합계</a>
					</div>                  
                  	</td>
				    <td width="50%">
                    </td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('insa_pay_overtime_save3.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>','pop_report','scrollbars=yes,width=1050,height=500')" class="btnType04">야특근수당 처리</a>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

