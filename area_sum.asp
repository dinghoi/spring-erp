<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim company_tab(150)
dim area_tab
area_tab = array("서울","경기","부산","대구","인천","광주","대전","울산","강원","경남","경북","세종","충남","충북","전남","전북","제주")
dim as_cnt(16)
dim as_per(16)

'ck_sw=Request("ck_sw")
c_name = "전체"

'If ck_sw = "n" Then
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	company = request.form("company")
'Else
'	from_date=Request("from_date")
'	to_date=Request("to_date")
'	company = "전체"
'End if

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	company = "전체"
End If

if company = "전체" then
	sql = "select count(*) as err_tot from as_acpt "
	sql = sql + "WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
  else
	sql = "select count(*) as err_tot from as_acpt "
	sql = sql + "WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"') "
	sql = sql + " and company = '" + company + "'"
end if

Rs.Open Sql, Dbconn, 1
err_tot = cint(rs("err_tot"))
if rs.eof then
	err_tot = 0
end if

rs.close()
for i = 0 to 16
	sido = area_tab(i)
	if company = "전체" then
		sql = "select sido,COUNT(*) AS err_cnt FROM as_acpt"
		sql = sql + " WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		sql = sql + " GROUP BY sido"
		sql = sql + " HAVING (sido = '"+sido+"')"
	  else
		sql = "select company,sido,COUNT(*) AS err_cnt FROM as_acpt"
		sql = sql + " WHERE (CAST(acpt_date as date) >= '" + from_date + "' AND CAST(acpt_date as date) <= '"+to_date+"')"
		sql = sql + " GROUP BY company,sido"
		sql = sql + " HAVING (company = '"+company+"') AND (sido = '"+sido+"')"
	end if

	Rs.Open Sql, Dbconn, 1

	if rs.eof then
		as_cnt(i) = 0
		as_per(i) = 0
		else
		as_cnt(i) = cint(rs("err_cnt"))
		as_per(i) = formatnumber((as_cnt(i)/err_tot * 100),2)
	end if
	rs.close()

next

title_line = "지역별 통계 현황"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
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
				if (document.frm.from_date.value > document.frm.to_date.value) {
					alert ("시작일이 종료일보다 클수가 없습니다");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/sum_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=area_sum.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%="1900-01-01"%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<strong>회사</strong>
							  	<%
									sql="select * from trade where use_sw = 'Y'  and (trade_id = '매출' or trade_id = '공통') order by trade_name asc"
                                    rs_trade.Open Sql, Dbconn, 1
                                %>
        						<select name="company" id="company" style="width:150px">
									<option value="전체">전체</option>
          					<%
								While not rs_trade.eof
							%>
          							<option value='<%=rs_trade("trade_name")%>' <%If rs_trade("trade_name") = company  then %>selected<% end if %>><%=rs_trade("trade_name")%></option>
          					<%
									rs_trade.movenext()
								Wend
								rs_trade.Close()
							%>
        						</select>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="*" >
							<col width="10%" >
							<col width="15%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">시도</th>
								<th scope="col">그래프</th>
								<th scope="col">건수</th>
								<th scope="col">백분률(%)</th>
							</tr>
						</thead>
						<tbody>
							<tr>
                              <th>총계</th>
                              <td class="left">&nbsp;</th>
                              <th><%=formatnumber(clng(err_tot),0)%></th>
                              <th>100%</th>
							</tr>
						<%
                    	for i = k to 16
                		%>
							<tr>
                              <td><%=area_tab(i)%></td>
                              <td class="left"><img src="image/graph02.gif" height="15px" width="<%=as_per(i)%>%"></td>
                              <td><%=formatnumber(clng(as_cnt(i)),0)%></td>
                              <td><%=as_per(i)%>%</td>
							</tr>
                		<%
						next
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</div>
	</body>
</html>

