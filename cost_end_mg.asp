<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_cost = Server.CreateObject("ADODB.Recordset")
Set Rs_insa = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


' 2019.02.22 박정신 요구 'N/W 1사업부','N/W 2사업부',"SI3사업부","솔루션사업부"	는 나오지않도록 조건으로 처리..
sql = "SELECT *                                " & chr(13) & _
      "  FROM emp_org_mst                      " & chr(13) & _
      " WHERE (org_level = '사업부')           " & chr(13) & _
      "   AND (org_name <> '총괄대표')         " & chr(13) & _
      "   AND (    ISNULL(org_end_date)        " & chr(13) & _
      "         OR org_end_date = '0000-00-00' " & chr(13) & _
      "       )                                " & chr(13)
' org_end_date = '' or   ....   date형을 '' 으로 비교할수없다.   Warning: Incorrect date value: '' for column 'org_end_date' at row 1

if cost_grade = "0" then
    sql = sql                             & chr(13) & _
        " GROUP BY org_saupbu, org_name " & chr(13) & _
	    " ORDER BY org_name ASC         " & chr(13)
else
    sql = sql                                                          & chr(13) & _
        "   AND (org_name = '"&saupbu&"' OR org_empno ='"&emp_no&"') " & chr(13) & _
	    " GROUP BY org_name                                          " & chr(13)
end if

Rs.Open Sql, Dbconn, 1
'Response.write "<pre>"&Sql&"</pre>"

sql = " SELECT MAX(org_month)  max_org_month  " & chr(13) & _
      "   FROM emp_org_mst_month              " & chr(13)
set Rs_insa=dbconn.execute(sql)
if	isnull(Rs_insa("max_org_month"))  then
    max_org_month = "000000"
else
    max_org_month = Rs_insa("max_org_month")
end if

title_line = "비용 마감 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
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
			function frmcheck () {
					document.frm.submit();
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="cost_end_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건검색</dt>
						<dd>
						<p>
							<label>&nbsp;&nbsp;<strong>최신정보로 다시 조회하기&nbsp;</strong></label>
							<a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
							(영업부 월마감,인사부 조직코드마감, 인사마감[<%=max_org_month%>] 확인)
						</p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*%" >
							<col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="13%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">사업부</th>
								<th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">현 재 마 감 현 황</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">새로운 마감 처리</th>
								<th rowspan="2" scope="col">보고자료</th>
								<th rowspan="2" scope="col">본부장보고</th>
								<th rowspan="2" scope="col">CEO보고</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">마감년월</th>
							  <th scope="col">마감상태</th>
							  <th scope="col">마감자</th>
							  <th scope="col">처리일자</th>
							  <th scope="col">마감취소</th>
							  <th scope="col">마감년월</th>
							  <th scope="col">마감처리</th>
              				</tr>
						</thead>
						<tbody>
						<%
							do until rs.eof

								cancel_yn = "N"
								if rs("org_bonbu") = "직할사업부" then
									if rs("org_saupbu") = "공항지원사업부" or rs("org_saupbu") = "KAL지원사업부" then
										jik_yn = "N"
								  else
										jik_yn = "Y"
							  	end if
							  else
							  	jik_yn = "N"
								end if

								sql="SELECT MAX(end_month) AS max_month "&_
								    "  FROM cost_end "&_
								    " WHERE saupbu='"&rs("org_name")&"'"
								set rs_max=dbconn.execute(sql)
								'Response.write sql &"<br>"

								sql="SELECT * "&_
								    "  FROM cost_end "&_
								    " WHERE saupbu='"&rs("org_name")&"' "&_
                                    "   AND end_month ='"&rs_max("max_month")&"'"
'response.write sql & "<br>"
								set rs_cost=dbconn.execute(sql)

								if rs_cost.eof or rs_cost.bof then
									new_date = dateadd("m",-1,now())
									new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
									now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
									end_month = "없음"

									if end_month = "없음" then
'										new_date = "2015-01-01"
										new_date = rs("org_date")
										new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
									end if

									end_view = ""
									batch_yn = ""
									bonbu_yn = ""
									ceo_yn = ""
									batch_view = ""
									ceo_view = ""
									reg_name = ""
									reg_id = ""
									reg_date = ""
							  else
									new_date = dateadd("m",1,datevalue(mid(rs_cost("end_month"),1,4) + "-" + mid(rs_cost("end_month"),5,2) + "-01"))
									new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
									now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)

									if  rs_cost("end_yn") = "Y" then
										end_view = "마감"
								  elseif rs_cost("end_yn") = "C" then
										new_month = rs_cost("end_month")
										end_view = "취소"
								  else
										end_view = "진행"
									end if

									end_yn = rs_cost("end_yn")
									end_month = rs_cost("end_month")
									reg_name = rs_cost("reg_name")
									reg_id = rs_cost("reg_id")
									reg_date = rs_cost("reg_date")

									if rs_cost("batch_yn") = "Y" then
										batch_view = "자료생성"
								  else
								  	batch_view = "미생성"
									end if
									if rs_cost("bonbu_yn") = "Y" then
										bonbu_view = "승인완료"
									end if
									if rs_cost("ceo_yn") = "Y" then
										ceo_view = "승인완료"
									end if
									if rs_cost("batch_yn") = "Y" and rs_cost("bonbu_yn") = "N" then
										bonbu_view = "진행중"
								  	ceo_view = ""
									end if
									if rs_cost("bonbu_yn") = "Y" and rs_cost("ceo_yn") = "N" then
										ceo_view = "진행중"
									end if
									if rs_cost("batch_yn") = "N" and rs_cost("bonbu_yn") = "N" and rs_cost("ceo_yn") = "N" then
										bonbu_view = ""
									  ceo_view = ""
									end if

									batch_yn = rs_cost("batch_yn")
									bonbu_yn = rs_cost("bonbu_yn")
									ceo_yn = rs_cost("ceo_yn")
								end if

								if jik_yn = "Y" then
									if ceo_yn = "N" then
										cancel_yn = "Y"
									end if
							  else
							  	if bonbu_yn = "N" then
										cancel_yn = "Y"
									end if
								end if
							%>
							<tr>
								<td class="first"><%=rs("org_name")%></td>
								<td><%=end_month%></td>
								<td>
									<%
									if end_view = "취소" then
										Response.write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									else
										Response.write end_view
									end if
									%>&nbsp;
								</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
									<%
									if cancel_yn = "Y" then
										Response.write "<a href='cost_end_cancel.asp?saupbu="&rs("org_name")&"&end_month="&end_month&"' class='btnType03'>마감취소</a>"
									else
										Response.write "취소불가"
									end if
									%>
								</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									if now_month > new_month then
										Response.write "<a href='cost_end_pro.asp?saupbu="&rs("org_name")&"&end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>마감</a>"
									else
										Response.write "마감불가"
									end if
									%>
                				</td>
								<td><%=batch_view%>&nbsp;</td>
								<td><%=bonbu_view%>&nbsp;</td>
								<td><%=ceo_view%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
							<%
							if cost_grade = "0" then
								sql="SELECT MAX(end_month) AS max_month "&_
								    "  FROM cost_end "&_
								    " WHERE saupbu='사업부외나머지'"
								set rs_max=dbconn.execute(sql)

							sql="SELECT * "&_
							    "  FROM cost_end "&_
							    " WHERE saupbu='사업부외나머지' "&_
							    "   AND end_month ='"&rs_max("max_month")&"'"
							set rs_cost=dbconn.execute(sql)

							if rs_cost.eof or rs_cost.bof then
								new_date = "2015-01-01"
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
								end_month = "없음"
								end_view = ""
								batch_yn = ""
								bonbu_yn = ""
								ceo_yn = ""
								batch_view = ""
								ceo_view = ""
								reg_name = ""
								reg_id = ""
								reg_date = ""
							else
								new_date = dateadd("m",1,datevalue(mid(rs_cost("end_month"),1,4) + "-" + mid(rs_cost("end_month"),5,2) + "-01"))
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)

								if  rs_cost("end_yn") = "Y" then
									end_view = "마감"
								elseif rs_cost("end_yn") = "C" then
									new_month = rs_cost("end_month")
									end_view = "취소"
								else
									end_view = "진행"
								end if

								end_yn = rs_cost("end_yn")
								end_month = rs_cost("end_month")
								reg_name = rs_cost("reg_name")
								reg_id = rs_cost("reg_id")
								reg_date = rs_cost("reg_date")

								if rs_cost("batch_yn") = "Y" then
									batch_view = "자료생성"
								else
									batch_view = "미생성"
								end if
								if rs_cost("bonbu_yn") = "Y" then
									bonbu_view = "승인완료"
								end if
								if rs_cost("ceo_yn") = "Y" then
									ceo_view = "승인완료"
								end if
								if rs_cost("batch_yn") = "Y" and rs_cost("bonbu_yn") = "N" then
									bonbu_view = "진행중"
								  ceo_view = ""
								end if
								if rs_cost("bonbu_yn") = "Y" and rs_cost("ceo_yn") = "N" then
								  	ceo_view = "진행중"
								end if
								if rs_cost("batch_yn") = "N" and rs_cost("bonbu_yn") = "N" and rs_cost("ceo_yn") = "N" then
									bonbu_view = ""
								  ceo_view = ""
								end if

								batch_yn = rs_cost("batch_yn")
								bonbu_yn = rs_cost("bonbu_yn")
								ceo_yn = rs_cost("ceo_yn")
							end if
							if jik_yn = "Y" then
								if ceo_yn = "N" then
									cancel_yn = "Y"
								end if
							else
								if bonbu_yn = "N" then
									cancel_yn = "Y"
								end if
							end if
							%>

							<tr bgcolor="#FFE8E8">
								<td class="first">사업부외나머지</td>
								<td><%=end_month%></td>
								<td>
									<%
									if end_view = "취소" then
										Response.write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									else
										Response.write end_view
									end if
									%>&nbsp;
								</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
									<%
									if cancel_yn = "Y" then
										Response.write "<a href='cost_bonbu_end_cancel.asp?saupbu=사업부외나머지&end_month="&end_month&"' class='btnType03'>마감취소</a>"
									else
										Response.write "취소불가"
									end if
									%>
								</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									if now_month > new_month then
										Response.write "<a href='cost_bonbu_end_pro.asp?saupbu=사업부외나머지&end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>마감</a>"
									else
										Response.write "마감불가"
									end if
									%>
								</td>
								<td><%=batch_view%>&nbsp;</td>
								<td><%=bonbu_view%>&nbsp;</td>
								<td><%=ceo_view%>&nbsp;</td>
							</tr>
						<%
						end if

						if user_id = "900001" or user_id = "100359" or user_id = "100952" or user_id = "101100" or user_id = "102305" or user_id = "102306" Or user_id = "102592" then
'						if user_id = "900001" then
							sql="SELECT MAX(end_month) AS max_month "&_
							    "  FROM cost_end "&_
							    " WHERE saupbu='상주비용'"
							set rs_max=dbconn.execute(sql)

							sql="SELECT * "&_
							    "  FROM cost_end "&_
							    " WHERE saupbu='상주비용' "&_
							    "   AND end_month ='"&rs_max("max_month")&"'"
							set rs_cost=dbconn.execute(sql)

							if rs_cost.eof or rs_cost.bof then
								new_date = "2015-01-01"
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
								end_month = "없음"
								end_yn = ""
								end_view = ""
								batch_yn = ""
								bonbu_yn = ""
								ceo_yn = ""
								batch_view = ""
								ceo_view = ""
								reg_name = ""
								reg_id = ""
								reg_date = ""
							else
								new_date = dateadd("m",1,datevalue(mid(rs_cost("end_month"),1,4) + "-" + mid(rs_cost("end_month"),5,2) + "-01"))
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)

								if  rs_cost("end_yn") = "Y" then
									end_view = "마감"
								elseif rs_cost("end_yn") = "C" then
									new_month = rs_cost("end_month")
									end_view = "취소"
								else
									end_view = "진행"
								end if

								end_yn = rs_cost("end_yn")
								end_month = rs_cost("end_month")
								reg_name = rs_cost("reg_name")
								reg_id = rs_cost("reg_id")
								reg_date = rs_cost("reg_date")

								if rs_cost("batch_yn") = "Y" then
									batch_view = "자료생성"
								else
									batch_view = "미생성"
								end if
								if rs_cost("bonbu_yn") = "Y" then
									bonbu_view = "승인완료"
								end if
								if rs_cost("ceo_yn") = "Y" then
									ceo_view = "승인완료"
								end if
								if rs_cost("batch_yn") = "Y" and rs_cost("bonbu_yn") = "N" then
									bonbu_view = "진행중"
								  ceo_view = ""
								end if
								if rs_cost("bonbu_yn") = "Y" and rs_cost("ceo_yn") = "N" then
									ceo_view = "진행중"
								end if
								if rs_cost("batch_yn") = "N" and rs_cost("bonbu_yn") = "N" and rs_cost("ceo_yn") = "N" then
									bonbu_view = ""
								  ceo_view = ""
								end if

								batch_yn = rs_cost("batch_yn")
								bonbu_yn = rs_cost("bonbu_yn")
								ceo_yn = rs_cost("ceo_yn")
							end if

							if jik_yn = "Y" then
								if ceo_yn = "N" then
									cancel_yn = "Y"
								end if
							else
								if bonbu_yn = "N" then
									cancel_yn = "Y"
								end if
							end if
						%>

							<tr bgcolor="#FFFFCC">
								<td class="first">상주비용</td>
								<td><%=end_month%></td>
								<td>
									<%
									if end_view = "취소" then
										Response.write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									else
										Response.write end_view
									end if
									%>&nbsp;
                				</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  <td>
									<%
									if cancel_yn = "Y" then
										Response.write "<a href='company_cost_end_cancel.asp?end_month="&end_month&"'  class='btnType03'>마감취소</a>"
									else
										Response.write "취소불가"
									end if
									%>
								</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									if now_month > new_month then
										Response.write "<a href='company_cost_end_pro.asp?end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>마감</a>"
									else
										Response.write "마감불가"
									end if
									%>
								</td>
							  	<td colspan="3">&nbsp;</td>
							</tr>
							<%
							sql="SELECT MAX(end_month) AS max_month "&_
							    "  FROM cost_end "&_
							    " WHERE saupbu='공통비/직접비배분'"
							set rs_max=dbconn.execute(sql)

							sql="SELECT * "&_
							    "  FROM cost_end "&_
							    " WHERE saupbu='공통비/직접비배분' "&_
							    "   AND end_month ='"&rs_max("max_month")&"'"
							set rs_cost=dbconn.execute(sql)
							if rs_cost.eof or rs_cost.bof then
								new_date = "2015-01-01"
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
								end_month = "없음"
								end_view = ""
								batch_yn = ""
								bonbu_yn = ""
								ceo_yn = ""
								batch_view = ""
								ceo_view = ""
								reg_name = ""
								reg_id = ""
								reg_date = ""
							else
								new_date = dateadd("m",1,datevalue(mid(rs_cost("end_month"),1,4) + "-" + mid(rs_cost("end_month"),5,2) + "-01"))
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)

								if  rs_cost("end_yn") = "Y" then
									end_view = "마감"
								elseif rs_cost("end_yn") = "C" then
									new_month = rs_cost("end_month")
									end_view = "취소"
								else
									end_view = "진행"
								end if

								end_yn = rs_cost("end_yn")
								end_month = rs_cost("end_month")
								reg_name = rs_cost("reg_name")
								reg_id = rs_cost("reg_id")
								reg_date = rs_cost("reg_date")

								if rs_cost("batch_yn") = "Y" then
									batch_view = "자료생성"
								else
								  	batch_view = "미생성"
								end if
								if rs_cost("bonbu_yn") = "Y" then
									bonbu_view = "승인완료"
								end if
								if rs_cost("ceo_yn") = "Y" then
									ceo_view = "승인완료"
								end if
								if rs_cost("batch_yn") = "Y" and rs_cost("bonbu_yn") = "N" then
									bonbu_view = "진행중"
								  	ceo_view = ""
								end if
								if rs_cost("bonbu_yn") = "Y" and rs_cost("ceo_yn") = "N" then
								  ceo_view = "진행중"
								end if
								if rs_cost("batch_yn") = "N" and rs_cost("bonbu_yn") = "N" and rs_cost("ceo_yn") = "N" then
									bonbu_view = ""
								  	ceo_view = ""
								end if

								batch_yn = rs_cost("batch_yn")
								bonbu_yn = rs_cost("bonbu_yn")
								ceo_yn = rs_cost("ceo_yn")
							end if
							if jik_yn = "Y" then
								if ceo_yn = "N" then
									cancel_yn = "Y"
								end if
							else
								if bonbu_yn = "N" then
									cancel_yn = "Y"
								end if
							end if
							%>
							<tr bgcolor="#CCFFFF">
								<td class="first">공통비/직접비배분</td>
					  	  		<td><%=end_month%></td>
								<td>
									<%
									if end_view = "취소" then
										Response.write "<span style='color:#F30; font-weight:bold'>"&end_view&"</span>"
									else
										Response.write end_view
									end if
									%>&nbsp;
								</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
							  		<%
							  		if cancel_yn = "Y" then
							  			Response.write "<a href='company_as_sum_cancel.asp?end_month="&end_month&"' class='btnType03'>마감취소</a>"
									else
										Response.write "취소불가"
									end if
									%>
                				</td>
								<td>
									<input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true">
								</td>
								<td>
									<%
									if now_month > new_month then
										Response.write "<a href='company_as_sum_pro.asp?end_month="&new_month&"&end_yn="&end_yn&"' class='btnType03'>마감</a>"
									else
										Response.write "마감불가"
									end if
									%>
								</td>
								<td colspan="3">&nbsp;</td>
						  	</tr>
							<%
							end if
							%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</div>
	</body>
</html>