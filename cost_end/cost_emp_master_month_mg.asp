<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim page, page_cnt, pg_cnt, be_pg, curr_date
Dim ck_sw, srchEmpName, srchWord, srchCategory, srchEmpMonth, view_sort
Dim view_condi, pgsize, start_page, stpage, end_yn
Dim listCostCenter
Dim order_sql, where_sql, total_record, total_page, title_line
Dim rs_cost_end, rs_etc, rsCount, rs
'Dim sql

page = Request("page")
page_cnt = Request.Form("page_cnt")
pg_cnt = CInt(Request("pg_cnt"))

ck_sw = Request("ck_sw")
srchEmpName = Request("srchEmpName")
srchWord = Request("srchWord")
srchCategory = Request("srchCategory")

srchEmpMonth = Request("srchEmpMonth")
view_sort = Request("view_sort")

If ck_sw = "y" Then
	view_condi = Request("view_condi")
Else
	view_condi = Request.Form("view_condi")
End if

If view_condi = "" Then
	'view_condi = "케이원정보통신"
	view_condi = "케이원"
End If

be_pg = "/cost/cost_emp_master_month_mg.asp"
curr_date = DateValue(Mid(CStr(Now()), 1, 10))

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

'//마지막 월
If Trim(srchEmpMonth&"") = "" Then
	srchEmpMonth = Left(Replace(DateAdd("m", -1, Now()), "-", ""), 6)
End If

'sql="select max(pmg_yymm) as max_pmg_yymm from pay_month_give "
'set rs_max=dbconn.execute(sql)
'If Not(rs_max.bof Or rs_max.eof) Then
'	empMonth = rs_max("max_pmg_yymm")
'End If
'rs_max.close : Set rs_max = Nothing

'//검색년월 마감 여부
end_yn = "N"

'sql="select count(1) as cnt from cost_end where end_yn='Y' and end_month ='"&srchEmpMonth&"'"
objBuilder.Append "SELECT COUNT(1) AS cnt "
objBuilder.Append "FROM cost_end "
objBuilder.Append "WHERE end_yn = 'Y' AND end_month = '"&srchEmpMonth&"' "

Set rs_cost_end = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not(rs_cost_end.EOF Or rs_cost_end.BOF) Then
	If CInt(rs_cost_end("cnt")) > 0 Then
		end_yn = "Y"
	End If
End If

rs_cost_end.close() : Set rs_cost_end = Nothing

'//비용구분
'sql = "select emp_etc_name from emp_etc_code where emp_etc_type = '70' order by emp_etc_code asc"
objBuilder.Append "SELECT emp_etc_name "
objBuilder.Append "FROM emp_etc_code "
objBuilder.Append "WHERE emp_etc_type = '70' "
objBuilder.Append "ORDER BY emp_etc_code ASC "

Set rs_etc = Server.CreateObject("ADODB.Recordset")
rs_etc.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Set listCostCenter = getRsToDic(rs_etc)
rs_etc.close : Set rs_etc = Nothing

If view_sort = "" Then
	view_sort = "ASC"
End If

'order_sql = "ORDER BY eomt.org_company " & view_sort & " "

'where_sql = where_sql & " AND A.pmg_id=1 " & chr(13) ' 왜 이 라인에 주석을 잡았었지? 박정신부장 문의 2180-08-14 (비용현황관리/지원월별현황 에서 2건이상나오는 문제..)
'where_sql = where_sql & " AND B.cost_except in ('0','1') " & chr(13)	'기존 주석 처리 코드

'회사명 검색
'If view_condi <> "전체" Then
'    where_sql = where_sql & " AND A.pmg_company='" & view_condi & "' " & chr(13)
'End If

'이름 검색
'If Trim(srchWord & "") <>"" Then
'    where_sql = where_sql & " AND B." & srchCategory & " like '%" & srchWord & "%' " & chr(13)
'End If

'Sql = "SELECT count(*) FROM pay_month_give  A ,emp_master_month B  " & where_sql

objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '" & srchEmpMonth & "' "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE pmgt.pmg_id = '1' "
objbuilder.Append "	AND pmgt.pmg_yymm = '" & srchEmpMonth & "' "

If view_condi <> "전체" Then
	objBuilder.Append "AND eomt.org_company = '" & view_condi & "' "
End If

If Trim(srchWord) <> "" Then
	objBuilder.Append "AND emmt."& srchCategory &" LIKE '%" & srchWord & "%' "
End If

'objBuilder.Append order_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(RsCount(0)) 'Result.RecordCount

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

'sql = " SELECT  A.pmg_yymm           " & chr(13) & _
'      "       , A.pmg_company        " & chr(13) & _
'      "       , A.pmg_saupbu         " & chr(13) & _
'      "       , A.pmg_give_total     " & chr(13) & _
'      "       , A.cost_group         " & chr(13) & _
'      "       , A.cost_center        " & chr(13) & _
'      "       , B.emp_no             " & chr(13) & _
'      "       , B.emp_name           " & chr(13) & _
'      "       , B.emp_job            " & chr(13) & _
'      "       , B.emp_type           " & chr(13) & _
'      "       , B.emp_saupbu         " & chr(13) & _
'      "       , B.emp_org_name       " & chr(13) & _
'      "       , B.emp_company        " & chr(13) & _
'      "       , B.emp_bonbu          " & chr(13) & _
'      "       , B.emp_team           " & chr(13) & _
'      "       , B.emp_reside_company " & chr(13) & _
'      "       , B.emp_reside_place   " & chr(13) & _
'      "    FROM pay_month_give A     " & chr(13) & _
'      "       , emp_master_month B   " & chr(13) & _
'      where_sql                        & chr(13) & _
'      order_sql                        & chr(13) & _
'      " LIMIT "& stpage & ", " & pgsize & chr(13)

objBuilder.Append "SELECT pmgt.pmg_yymm, pmgt.pmg_company, pmgt.pmg_give_total, "
objBuilder.Append "	pmgt.cost_group, pmgt.cost_center, "
objBuilder.Append "	emmt.emp_no, emmt.emp_name, emmt.emp_job, emmt.emp_type, "
objBuilder.Append "	emmt.emp_org_name, emmt.emp_company, emmt.emp_bonbu, emmt.emp_saupbu, emmt.emp_team, "
objBuilder.Append "	emmt.emp_reside_company, emmt.emp_reside_place, emmt.emp_stay_name, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, eomt.org_code, "
objBuilder.Append "	eomt.org_reside_company, eomt.org_reside_place "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '" & srchEmpMonth & "' "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "	AND eomt.org_month = '" & srchEmpMonth & "' "
objBuilder.Append "WHERE pmgt.pmg_id = '1' "
objBuilder.Append "	AND pmgt.pmg_yymm = '" & srchEmpMonth & "' "

If view_condi <> "전체" Then
	objBuilder.Append "AND eomt.org_company = '" & view_condi & "' "
End If

If Trim(srchWord) <> "" Then
	objBuilder.Append "AND emmt."& srchCategory &" LIKE '%" & srchWord & "%' "
End If

objBuilder.Append "ORDER BY eomt.org_name " & view_sort & ", "
objBuilder.Append "	FIELD(emmt.emp_job, '사장', '부사장', '총괄대표', '전무이사', '상무이사', '이사', "
objBuilder.Append "		'전문위원', '연구소장', '부장', '차장', '과장', '수석연구원', '책임연구원', "
objBuilder.Append "		'대리', '대리1급', '대리2급', '주임연구원', '연구원', '사원') "

objBuilder.Append "LIMIT "& stpage & ", " & pgsize

'Response.write objBuilder.ToString()

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

title_line = " 직원 월별 현황 " & "(" & srchEmpMonth & ")"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용관리 시스템</title>
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

			function searchEnter(name){
				$("input[name = "+name+"]").on("keyup", function(e){
					if(e.keyCode === 13){
						frmcheck();
					}
				});
			}

			var empMonth = "<%=srchEmpMonth%>";

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.view_condi.value == ""){
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}
				return true;
			}

			function callbackOrgSelect(empNo, costGroup){
				var tEmpNo = "";
				$(".tableList tbody tr").each(function(){
					tEmpNo = $(this).find("td").eq(0).text();
					if( tEmpNo == empNo ){
						$(this).find("input[name='cost_group']").val(costGroup);
					}
				});
			}

			function callbackOrgCodeSelect(empNo, orgCode){
				var tEmpNo = "";
				$(".tableList tbody tr").each(function(){
					tEmpNo = $(this).find("td").eq(0).text();
					if( tEmpNo == empNo ){
						$(this).find("input[name='cost_org_code']").val(orgCode);
					}
				});
			}

            function changeEmpMasterMonth(empNo){
				empMonth = $("#srchEmpMonth").val();

				var tEmpNo = "";
				var tCostCenter = "";
				var tCostGroup = "";

				var tEmpOrgCode = "";
				var tEmpOrgName = "";
				var tEmpCompany = "";
				var tEmpBonbu = "";
				var tEmpSaupbu = "";
				var tEmpTeam = "";
				var tEmpResideCompany = "";
				var tEmpResidePlace = "";

				$(".tableList tbody tr").each(function(){
					tEmpNo = $(this).find("td").eq(0).text();

					if( tEmpNo == empNo ){
						tCostCenter = $(this).find("select option:selected").val();
						tCostGroup = $(this).find("input[name='cost_group']").val();
						//alert(tEmpNo + "::" + tCostCenter + "::"+tCostGroup);
						tEmpOrgCode = $(this).find("input[name='emp_org_code']").val();
						tEmpOrgName = $(this).find("input[name='emp_org_name']").val();
						tEmpCompany = $(this).find("input[name='emp_company']").val();
						tEmpBonbu = $(this).find("input[name='emp_bonbu']").val();
						tEmpSaupbu = $(this).find("input[name='emp_saupbu']").val();
						tEmpTeam = $(this).find("input[name='emp_team']").val();

						tCostOrgCode = $(this).find("input[name='cost_org_code']").val();

						return false;
					}
				});

				if(empMonth==null || empMonth==""){
					alert("년월 정보가 없습니다.");
					return false;
				}
				if(tEmpNo==null || tEmpNo==""){
					alert("사번정보가 없습니다.");
					return false;
				}
				if(tCostCenter==null || tCostCenter==""){
					alert("비용구분을 선택해주세요.");
					return false;
				}
				if(tCostGroup==null || tCostGroup==""){
					alert("비용그룹을 입력해주세요.");
					return false;
				}

				var params = {
                                "empMonth"    : empMonth
                                ,"empNo"      : tEmpNo
                                ,"costCenter" : escape(tCostCenter)
                                ,"costGroup"  : escape(tCostGroup)
                                ,"empOrgCode" : escape(tEmpOrgCode)
                                ,"empOrgName" : escape(tEmpOrgName)
                                ,"empCompany" : escape(tEmpCompany)
                                ,"empBonbu"   : escape(tEmpBonbu)
                                ,"empSaupbu"  : escape(tEmpSaupbu)
                                ,"empTeam"    : escape(tEmpTeam)
								,"costOrgCode"    : escape(tCostOrgCode)
                             };
				$.ajax({
 					 url: "/cost/cost_emp_master_month_mg_save.asp"
					,type: 'post'
					,data: params
					,dataType: "json"
					,contentType: "application/x-www-form-urlencoded; charset=euc-kr"
					,beforeSend: function(jqXHR){
							jqXHR.overrideMimeType("application/x-www-form-urlencoded; charset=euc-kr");
						}
					//,success:function(data, status, request){
					,success: function(data){
						var result = data.result;
						if( result=="succ"){
							alert("변경됐습니다.");
						}else if( result=="invalid" ){
							alert("입력하신 정보가 정확하지 않습니다.");
						}else if(result=="fail"){
							alert("저장 실패했습니다.");
						}
					}
					,error: function(jqXHR, status, errorThrown){
						alert("에러가 발생하였습니다.\n상태코드 : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
					}
				});
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="<%=be_pg%>" method="post" name="frm">

				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>검색</dt>
						<dd>
							<p>
								<label for="srchEmpMonth"><strong>마감년월 : </strong></label>
								<input type="text" name="srchEmpMonth" id="srchEmpMonth" value="<%=srchEmpMonth%>" />
								<label for="view_condi"><strong>회사 : </strong></label>
								<%
								Call SelectEmpOrgList("view_condi", "view_condi", "width:150px", view_condi)
								%>
								<!--<label for="srchEmpName"><strong>이름 : </strong></label>-->
								<select name="srchCategory" id="srchCategory">
									<option value="emp_name"<%If srchCategory="emp_name" Then%>selected<%End If%>>이름</option>
									<option value="emp_no"<%If srchCategory="emp_no" Then%>selected<%End If%>>사번</option>
								</select>
								<input type="text" name="srchWord" id="srchWord" style="width: 100px; text-align: left; -ms-ime-mode: active;" value="<%=srchWord%>" onkeypress="searchEnter('srchWord');"/>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>

				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%">
							<col width="4%">
							<col width="3%">
							<col width="16%">
							<col width="32%">
							<col width="5%">
							<col width="5%">
							<col width="8%">
							<col width="24%">
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">조직명</th>
								<th scope="col">조직</th>
								<th scope="col">상주처</th>
								<!--<th scope="col">실근무지</th>-->
								<th scope="col">상주회사</th>
								<th scope="col">비용구분</th>
								<th scope="col">비용그룹</th>
							</tr>
						</thead>
					<tbody>
						<%
						Dim j
						Dim tEmpNo, tCostCenter, tCostGroup, tEmpOrgName, tEmpCompany
						Dim tEmpBonbu,tEmpTeam, tEmpResideCompany, tEmpResidePlace
						Dim tEmpOrgCode, org, mg_org, row
						Dim tEmpSaupbu

						Int j = 0

						Do Until rs.EOF
							tEmpNo 				= rs("emp_no")
							tCostCenter 		= rs("cost_center")
							tCostGroup 			= rs("cost_group")

							'tEmpOrgName 		= rs("emp_org_name")
							'tEmpCompany 		= rs("emp_company")
							'tEmpBonbu 			= rs("emp_bonbu")
							'tEmpSaupbu 			= rs("emp_saupbu")
							'tEmpTeam 			= rs("emp_team")

							tEmpOrgName 		= rs("org_name")
							tEmpCompany 		= rs("org_company")
							tEmpBonbu 			= rs("org_bonbu")
							tEmpSaupbu 			= rs("org_saupbu")
							tEmpTeam 			= rs("org_team")
							tEmpOrgCode			= rs("org_code")

							tEmpResideCompany	= rs("emp_reside_company")
							tEmpResidePlace 	= rs("emp_reside_place")
							'tEmpResideCompany	= rs("org_reside_company")
							'tEmpResidePlace 	= rs("org_reside_place")

							j = j + 1
						    %>
							<tr>
								<td class="first"><%=rs("emp_no")%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("emp_job")%></td>
								<td>
									<input type="hidden" id="emp_org_code<%=j%>" name="emp_org_code" value="<%=tEmpOrgCode%>" />
									<input id="emp_org_name<%=j%>" name="emp_org_name" type="text" style="width:120px" readonly="true" value="<%=tEmpOrgName%>">
                                    <!--a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%'="org"%>&mg_org=<%'=mg_org%>&view_condi=<%'=view_condi%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a-->
                                    <a href="#" class="btnType03" onClick="pop_Window('/insa/insa_org_select.asp?gubun=<%="org"%>&mg_org=<%=mg_org%>&view_condi=<%=view_condi%>&org_id=<%=j%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a>
								</td>
								<td>
									<input type="text" id="emp_company<%=j%>" name="emp_company" readonly="true" value="<%=tEmpCompany%>" style="width:80px" />
									<input type="text" id="emp_bonbu<%=j%>" name="emp_bonbu" readonly="true" value="<%=tEmpBonbu%>" style="width:80px" />
									<input type="text" id="emp_saupbu<%=j%>" name="emp_saupbu" readonly="true" value="<%=tEmpSaupbu%>" style="width:80px" />
									<input type="text" id="emp_team<%=j%>" name="emp_team" readonly="true" value="<%=tEmpTeam%>" style="width:80px" />
									<input type="hidden" id="emp_reside_company<%=j%>" name="emp_reside_company" readonly="true" value="<%=tEmpResideCompany%>" />
									<input type="hidden" id="emp_reside_place<%=j%>" name="emp_reside_place" readonly="true" value="<%=tEmpResidePlace%>"   />
									<input type="hidden" id="emp_org_level<%=j%>" name="emp_org_level" readonly="true" value="" />
									<input type="hidden" id="emp_type<%=j%>" name="emp_type" value="<%=tEmpSaupbu%>" >

									<input type="hidden" id="emp_saupbu<%'=j%>" name="emp_saupbu" >
                                </td>
								<td><%=rs("org_reside_place")%></td>
								<td>
									<%'=rs("emp_stay_name")%>
									<%=rs("org_reside_company")%>
								</td>
                                <td>
									<select id="cost_center<%=j%>" name="cost_center" style="width:90px">
										<option value="">선택</option>
										<%
											If IsObject(listCostCenter) Then
												If listCostCenter.count > 0 Then
													For i=0 to listCostCenter.count-1
														Set row = listCostCenter.item(i)
                                                        %>
                                                        <option value='<%=row("emp_etc_name")%>' <%If tCostCenter = row("emp_etc_name") Then %>selected<%End If %>><%=row("emp_etc_name")%></option>
                                                        <%
													Next
												End If
											End If
										%>
									</select>
								</td>
               	                <td>
									<input type="text" id="cost_group<%=j%>" name="cost_group" value="<%=tCostGroup%>" readonly="readonly" />
									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_org_select.asp?gubun=costEmp&target=<%=rs("emp_no")%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a>

									<input type="hidden" name = "cost_org_code" value="" />

									<%' If end_yn <> "Y" Then %>
									    <a href="#" class="btnType04" onClick="changeEmpMasterMonth('<%=tEmpNo%>')">적용</a>
									<%' End If %>
								</td>
							</tr>
						    <%
							rs.MoveNext()
						Loop

						rs.Close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page, i

                intstart = (Int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                If intend > total_page Then
                    intend = total_page
                End If
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
					<!--
                    <a href="insa_excel_emp.asp?view_condi=<%=view_condi%>" class="btnType04">엑셀다운로드</a>
					-->
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="<%=be_pg%>?page=<%=first_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[처음]</a>
                  	<%If intstart > 1 Then %>
                        <a href="<%=be_pg%>?page=<%=intstart -1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[이전]</a>
                    <%End If %>
                    <%For i = intstart To intend %>
                  	<%	If i = int(page) Then %>
                        <b>[<%=i%>]</b>
                    <%	Else %>
                        <a href="<%=be_pg%>?page=<%=i%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                    <%	End If %>
                    <%Next %>

                  	<%If intend < total_page Then %>
                        <a href="<%=be_pg%>?page=<%=intend+1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[다음]</a> <a href="<%=be_pg%>?page=<%=total_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[마지막]</a>
                    <%Else %>
                        [다음]&nbsp;[마지막]
                    <%End If %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
					<!--
                    <a href="#" onClick="pop_Window('insa_emp_add01.asp?view_condi=<%=view_condi%>&u_type=<%=""%>','insa_emp_add01_popup','scrollbars=yes,width=1250,height=600')" class="btnType04">신규채용등록</a>
					-->
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