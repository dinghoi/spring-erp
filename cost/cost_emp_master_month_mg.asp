<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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

Dim pg_url, i

page = f_Request("page")
page_cnt = f_Request("page_cnt")
pg_cnt = CInt(f_Request("pg_cnt"))
srchEmpName = f_Request("srchEmpName")
srchWord = f_Request("srchWord")
srchCategory = f_Request("srchCategory")
srchEmpMonth = f_Request("srchEmpMonth")
view_sort = f_Request("view_sort")
view_condi = f_Request("view_condi")

If view_condi = "" Then
	view_condi = "케이원"
End If

curr_date = DateValue(Mid(CStr(Now()), 1, 10))

'//마지막 월
If Trim(srchEmpMonth&"") = "" Then
	srchEmpMonth = Left(Replace(DateAdd("m", -1, Now()), "-", ""), 6)
End If

be_pg = "/cost/cost_emp_master_month_mg.asp"

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_sort="&view_sort&"&view_condi="&view_condi&"&srchEmpMonth="&srchEmpMonth&"&srchCategory="&srchCategory&"&srchWord="&srchWord&"&srchEmpName="&srchEmpName

'//검색년월 마감 여부
end_yn = "N"

objBuilder.Append "SELECT COUNT(*) AS cnt "
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
objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
objBuilder.Append "WHERE emp_etc_type = '70' ORDER BY emp_etc_code ASC "

'Response.write objBuilder.toString()

Set rs_etc = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Set listCostCenter = getRsToDic(rs_etc)

rs_etc.close : Set rs_etc = Nothing

If view_sort = "" Then
	view_sort = "ASC "
End If

objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '" & srchEmpMonth & "' "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE pmgt.pmg_id = '1' "
objbuilder.Append "	AND pmgt.pmg_yymm = '"&srchEmpMonth&"' "

If view_condi <> "전체" Then
	objBuilder.Append "AND eomt.org_company = '"&view_condi&"' "
End If

If Trim(srchWord) <> "" Then
	objBuilder.Append "AND emmt."&srchCategory&" LIKE '%"&srchWord&"%' "
End If

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(RsCount(0)) 'Result.RecordCount

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

objBuilder.Append "SELECT pmgt.pmg_yymm, pmgt.pmg_company, pmgt.pmg_give_total, "
objBuilder.Append "	pmgt.cost_group, pmgt.cost_center, pmgt.pmg_company, pmgt.pmg_bonbu, pmgt.pmg_saupbu, pmgt.pmg_team, pmgt.pmg_org_name, "
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

Set rs = DBConn.Execute(objBuilder.ToString())
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
							location.reload();
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
								<input type="text" name="srchEmpMonth" id="srchEmpMonth" value="<%=srchEmpMonth%>" onkeypress="if(event.keyCode == '13'){frmcheck();}"/>
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
							<col width="9%">
							<col>
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

							tEmpOrgName 		= rs("pmg_org_name")
							tEmpCompany 		= rs("pmg_company")
							tEmpBonbu 			= rs("pmg_bonbu")
							tEmpSaupbu 			= rs("pmg_saupbu")
							tEmpTeam 			= rs("pmg_team")
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
                                    <a href="#" class="btnType03" onClick="pop_Window('/insa/insa_org_select.asp?gubun=org&mg_org=<%=mg_org%>&view_condi=<%=view_condi%>&org_id=<%=j%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a>
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
									<select id="cost_center<%=j%>" name="cost_center" style="width:100px">
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
               	                <td style="text-align:left;padding-left:2px;">
									<input type="text" id="cost_group<%=j%>" name="cost_group" value="<%=tCostGroup%>" readonly="readonly" style="width:130px;" />
									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_org_select.asp?gubun=costEmp&target=<%=rs("emp_no")%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a>

									<input type="hidden" name = "cost_org_code" value="" />

									<%'비용 마감 전에만 적용 가능 처리[허정호_20220210]
									'If end_yn <> "Y" Then
									%>
									    <a href="#" class="btnType04" onClick="changeEmpMasterMonth('<%=tEmpNo%>')">적용</a>
									<%
									'End If
									%>
								</td>
							</tr>
						    <%
							rs.MoveNext()
						Loop

						rs.Close() : Set rs = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
						<div class="btnCenter">
						<%
						If CInt(total_record) > 0 Then
						%>
							<a href="/cost/excel/cost_emp_master_excel.asp?srchEmpMonth=<%=srchEmpMonth%>" class="btnType04">추가 인원(엑셀)</a>
							<a href="/cost/excel/cost_appoint_excel.asp?srchEmpMonth=<%=srchEmpMonth%>" class="btnType04">이동 발령(엑셀)</a>
						<%
						End If
						%>
						</div>
                  	</td>
				    <td>
                    <%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)

					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
					<td width="20%">
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