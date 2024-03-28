<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!DOCTYPE HTML>
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim from_date, to_date, ck_sw, Page
Dim slip_month, emp_yn, emp_name, sort_condi
Dim end_date, be_from_date, be_to_date, pgsize
Dim stpage, condi_sql, order_sql, total_record
Dim be_date, start_page, title_line

'win_sw = "close"

ck_sw = Request("ck_sw")
Page = Request("page")

If ck_sw = "y" Then
	slip_month = Request("slip_month")
	emp_yn = Request("emp_yn")
	emp_name = Request("emp_name")
	sort_condi = Request("sort_condi")
Else
	slip_month = Request.Form("slip_month")
	emp_yn = Request.Form("emp_yn")
	emp_name = Request.Form("emp_name")
	sort_condi = Request.Form("sort_condi")
End if

If slip_month = "" Then
	be_date = DateAdd("m", -1, Now())
	slip_month = Mid(CStr(be_date), 1, 4) & Mid(CStr(be_date), 6, 2)
	emp_yn = "N"
	emp_name = ""
	sort_condi = "emp"
End If

If emp_yn = "N" Then
	emp_name = ""
End If

from_date = Mid(slip_month, 1, 4) & "-" & Mid(slip_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
be_from_date = DateAdd("m", -1, from_date)
be_to_date = Mid(be_from_date, 1, 4) & "-" & Mid(be_from_date, 6, 2) & "-31"

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = Int((page - 1) * pgsize)

' 조건 조회
If emp_yn = "Y" Then
	condi_sql = "AND cslt.emp_name LIKE '%"&emp_name&"%'"
Else
  	condi_sql = ""
End If

' 조회순서
If sort_condi = "emp" Then
	order_sql = "ORDER BY cslt.emp_name ASC "
Else
  	order_sql = "ORDER BY calt.price desc "
End If

' 레코드 건수
total_record = 0

Dim rsCardCnt, total_page

'sql = "select emp_no from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"&condi_sql&" group by emp_no"
objBuilder.Append "SELECT cslt.emp_no FROM card_slip AS cslt "
objBuilder.Append "WHERE (cslt.slip_date >='"&from_date&"' AND cslt.slip_date <='"&to_date&"') "
objBuilder.Append condi_sql
objBuilder.Append "GROUP BY cslt.emp_no "

Set rsCardCnt = Server.CreateObject("ADODB.RecordSet")
rsCardCnt.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

'count 체크 필요[허정호_20210321]
Do Until rsCardCnt.eof
	total_record = total_record + 1
	rsCardCnt.MoveNext()
Loop
rsCardCnt.Close() : Set rsCardCnt = Nothing

'total_record = cint(RsCount(0)) 'Result.RecordCount

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

Dim rsCardTot, sum_cnt, sum_cost, sum_cost_vat, sum_price
' 당월 금액 SUM 처리
'sql = "select count(*) as slip_cnt,sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"&condi_sql
objBuilder.Append "SELECT COUNT(*) AS slip_cnt, SUM(price) AS price, SUM(cost) AS cost, SUM(cost_vat) AS cost_vat "
objBuilder.Append "FROM card_slip AS cslt "
objBuilder.Append "WHERE (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append condi_sql

Set rsCardTot = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

sum_cnt = CDbl(rsCardTot("slip_cnt"))

If rsCardTot("price") = "" Or IsNull(rsCardTot("price")) Then
	sum_cost = 0
	sum_cost_vat = 0
	sum_price = 0
Else
	sum_cost = CDbl(rsCardTot("cost"))
	sum_cost_vat = CDbl(rsCardTot("cost_vat"))
	sum_price = CDbl(rsCardTot("price"))
End If
rsCardTot.Close() : Set rsCardTot = Nothing

Dim rsCardPrevTot, be_sum_cnt, be_sum_cost, be_sum_cost_vat, be_sum_price
' 전월 금액 전체 SUM 처리
'sql = "select count(*) as slip_cnt,sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (slip_date >='"&be_from_date&"' and slip_date <='"&be_to_date&"')"&condi_sql
objBuilder.Append "SELECT COUNT(*) AS slip_cnt, SUM(price) AS price, SUM(cost) AS cost, SUM(cost_vat) AS cost_vat "
objBuilder.Append "FROM card_slip AS cslt "
objBuilder.Append "WHERE (slip_date >='"&be_from_date&"' AND slip_date <='"&be_to_date&"') "
objBuilder.Append condi_sql

Set rsCardPrevTot = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

be_sum_cnt = CDbl(rsCardPrevTot("slip_cnt"))

If rsCardPrevTot("price") = "" Or IsNull(rsCardPrevTot("price")) Then
	be_sum_cost = 0
	be_sum_cost_vat = 0
	be_sum_price = 0
Else
	be_sum_cost = CDbl(rsCardPrevTot("cost"))
	be_sum_cost_vat = CDbl(rsCardPrevTot("cost_vat"))
	be_sum_price = CDbl(rsCardPrevTot("price"))
End If

rsCardPrevTot.Close() : Set rsCardPrevTot = Nothing

Dim rsCardList
'sql = "select card_slip.emp_no,card_slip.emp_name,memb.user_grade,memb.org_name,count(*) as slip_cnt,sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat "
'sql = sql + " from card_slip inner join memb on card_slip.emp_no=memb.user_id where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"&condi_sql&" group by card_slip.emp_no " + order_sql + " limit "& stpage & "," &pgsize

objBuilder.Append "SELECT cslt.emp_no, cslt.emp_name, memt.user_grade, eomt.org_name, eomt.org_company, "
objBuilder.Append "	COUNT(*) AS slip_cnt, SUM(price) AS price, SUM(cost) AS cost, SUM(cost_vat) AS cost_vat "
objBuilder.Append "FROM card_slip AS cslt "
objBuilder.Append "INNER JOIN memb AS memt ON cslt.emp_no = memt.user_id "
objBuilder.Append "INNER JOIN emp_master AS emtt ON cslt.emp_no = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (cslt.slip_date >='"&from_date&"' AND cslt.slip_date <='"&to_date&"')"
objBuilder.Append condi_sql&" GROUP BY cslt.emp_no "&order_sql&" "
objBuilder.Append "LIMIT "&stpage&", "&pgsize

'Response.write objBuilder.ToString()
'Response.end

Set rsCardList = Server.CreateObject("ADODB.RecordSet")
rsCardList.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

title_line = "카드 전표 관리"
%>

<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">-->
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
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

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.slip_month.value == ""){
					alert("사용년월을 입력하세요");
					return false;
				}
				//return true;
				return;
			}

			function condi_view(){

				if (eval("document.frm.emp_yn[0].checked")){
					document.getElementById('emp_name_view').style.display = 'none';
				}

				if (eval("document.frm.emp_yn[1].checked")){
					document.getElementById('emp_name_view').style.display = '';
				}
			}
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/card_slip_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/finance/card_person_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>사용년월&nbsp;</strong>(예201401) :
                                	<input name="slip_month" type="text" value="<%=slip_month%>" style="width:60px">
								</label>
                                <label>
								<strong>검색조건</strong>
                                  <input type="radio" name="emp_yn" value="N" <%If emp_yn = "N" Then%>checked<%End If %> style="width:30px" id="Radio1" onClick="condi_view()">전체 </label>
                                  <input type="radio" name="emp_yn" value="Y" <%If emp_yn = "Y" Then%>checked<%End If %> style="width:30px" id="Radio2" onClick="condi_view()">직원명
                                </label>
								&nbsp;&nbsp;
                                <label>
                                	<input name="emp_name" type="text" value="<%=emp_name%>" style="width:80px; display:none" id="emp_name_view">
								</label>
                                <label>
								<strong>조회순서</strong>
                                  <input type="radio" name="sort_condi" value="emp" <% if sort_condi = "emp" then %>checked<% end if %> style="width:30px" id="Radio1">직원순
                                  <input type="radio" name="sort_condi" value="price" <% if sort_condi = "price" then %>checked<% end if %> style="width:30px" id="Radio2">금액순
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="*" >
							<col width="5%" >
							<col width="8%" >
							<col width="7%" >
							<col width="8%" >
							<col width="5%" >
							<col width="8%" >
							<col width="7%" >
							<col width="8%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">직원명</th>
								<th rowspan="2" scope="col">조직명</th>
								<th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">전 월</th>
								<th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">당월</th>
								<th rowspan="2" scope="col">증감액</th>
								<th rowspan="2" scope="col">증감율</th>
								<th rowspan="2" scope="col">당월세부<p>내역조회</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">건수</th>
							  <th scope="col">공급가액</th>
							  <th scope="col">부가세</th>
							  <th scope="col">합계</th>
							  <th scope="col">건수</th>
							  <th scope="col">공급가액</th>
							  <th scope="col">부가세</th>
							  <th scope="col">합계</th>
						  </tr>
						</thead>
						<tbody>
							<tr>
								<th class="first">총계</th>
								<th><%=formatnumber(total_record,0)%>&nbsp;건</th>
							  	<th class="right"><%=formatnumber(be_sum_cnt,0)%></th>
							  	<th class="right"><%=formatnumber(be_sum_cost,0)%></th>
								<th class="right"><%=formatnumber(be_sum_cost_vat,0)%></th>
							  	<th class="right"><%=formatnumber(be_sum_price,0)%></th>
							  	<th class="right"><%=formatnumber(sum_cnt,0)%></th>
							  	<th class="right"><%=formatnumber(sum_cost,0)%></th>
								<th class="right"><%=formatnumber(sum_cost_vat,0)%></th>
							  	<th class="right"><%=formatnumber(sum_price,0)%></th>
							  	<th class="right"><%=formatnumber(sum_price-be_sum_price,0)%></th>
								<th class="right"><%=formatnumber((sum_price-be_sum_price)/be_sum_price*100,2)%>%</th>
								<th>&nbsp;</th>
							</tr>
						<%
						Dim rsCardPrevCnt, be_cnt, be_cost, be_cost_vat, be_price
						Dim incr_per
						Do Until rsCardList.EOF
							' 전월 금액 전체 SUM 처리
							'sql = "select count(*) as slip_cnt,sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (slip_date >='"&be_from_date&"' and slip_date <='"&be_to_date&"') and emp_no = '"&rs("emp_no")&"'"
							objBuilder.Append "SELECT COUNT(*) AS slip_cnt, SUM(price) AS price, sum(cost) as cost, sum(cost_vat) as cost_vat "
							objBuilder.Append "FROM card_slip "
							objBuilder.Append "where (slip_date >='"&be_from_date&"' and slip_date <='"&be_to_date&"') "
							objBuilder.Append "and emp_no = '"&rsCardList("emp_no")&"' "

							Set rsCardPrevCnt = DBconn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							be_cnt = cdbl(rsCardPrevCnt("slip_cnt"))

							if rsCardPrevCnt("price") = "" or isnull(rsCardPrevCnt("price")) then
								be_cost = 0
								be_cost_vat = 0
								be_price = 0
							else
								be_cost = CDbl(rsCardPrevCnt("cost"))
								be_cost_vat = cdbl(rsCardPrevCnt("cost_vat"))
								be_price = cdbl(rsCardPrevCnt("price"))
							end If
							rsCardPrevCnt.close()

							if be_price = 0 then
								incr_per = 100
							else
								incr_per = (cdbl(rsCardList("price")) - be_price) / be_price * 100
							end if
						%>
							<tr>
								<td class="first"><%=rsCardList("emp_name")%>&nbsp;<%=rsCardList("user_grade")%></td>
								<td class="left"><%=rsCardList("org_name")%></td>
							  	<td class="right"><%=formatnumber(be_cnt, 0)%></td>
							  	<td class="right"><%=formatnumber(be_cost, 0)%></td>
							  	<td class="right"><%=formatnumber(be_cost_vat, 0)%></td>
							  	<td class="right"><%=formatnumber(be_price, 0)%></td>
							  	<td class="right"><%=formatnumber(rsCardList("slip_cnt"), 0)%></td>
							  	<td class="right"><%=formatnumber(rsCardList("cost"), 0)%></td>
							  	<td class="right"><%=formatnumber(rsCardList("cost_vat"), 0)%></td>
							  	<td class="right"><%=formatnumber(rsCardList("price"), 0)%></td>
							  	<td class="right"><%=formatnumber(cdbl(rsCardList("price")) - be_price, 0)%></td>
							  	<td class="right"><%=formatnumber(incr_per, 2)%>%</td>
								<td>
                               	<input type="hidden" name="emp_no" value='rsCardList("emp_no")'%>
								<a href="#" onClick="pop_Window('/person_card_slip_view.asp?slip_month=<%=slip_month%>&emp_no=<%=rsCardList("emp_no")%>','카드전표수정','scrollbars=yes,width=900,height=500')">조회</a>
                                </td>
							</tr>
						<%
							rsCardList.MoveNext()
						Loop
						Set rsCardPrevCnt = Nothing
						rsCardList.close() : Set rsCardList = Nothing
						DBConn.Close : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page, i

                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="/finance/card_person_mg.asp?page=<%=first_page%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="/finance/card_person_mg.asp?page=<%=intstart -1%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[이전]</a>
                    <% end if %>
                      <% for i = intstart to intend %>
						<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
						<% else %>
                        <a href="/finance/card_person_mg.asp?page=<%=i%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
					<% If intend < total_page then %>
                        <a href="/finance/card_person_mg.asp?page=<%=intend+1%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[다음]</a>
                        <a href="/finance/card_person_mg.asp?page=<%=total_page%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[마지막]</a>
                     <%	else %>
                        [다음]&nbsp;[마지막]
                     <% end if %>
                    </div>
                    </td>
				    <td width="25%">
					<div class="btnCenter">
					</div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>

