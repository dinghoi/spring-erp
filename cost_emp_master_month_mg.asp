<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim page_cnt
Dim pg_cnt

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "cost_emp_master_month_mg.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")
srchEmpName=Request("srchEmpName")
srchWord=Request("srchWord")
srchCategory=Request("srchCategory")


srchEmpMonth=Request("srchEmpMonth")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
  else
	view_condi=Request.form("view_condi")
End if

If view_condi = "" Then
	view_condi = "케이원정보통신"
End If

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_max = Server.CreateObject("ADODB.Recordset")
Set rs_cost_end = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect


'//마지막 월
If Trim(srchEmpMonth&"")="" Then
	srchEmpMonth = Left(Replace(DateAdd("m",-1,now()),"-",""),6)
End If
'sql="select max(pmg_yymm) as max_pmg_yymm from pay_month_give "
'set rs_max=dbconn.execute(sql)
'If Not(rs_max.bof Or rs_max.eof) Then
'	empMonth = rs_max("max_pmg_yymm")
'End If
'rs_max.close : Set rs_max = Nothing

'//검색년월 마감 여부
end_yn = "N"
sql="select count(1) as cnt from cost_end where end_yn='Y' and end_month ='"&srchEmpMonth&"'"
set rs_cost_end=dbconn.execute(sql)
if Not(rs_cost_end.eof or rs_cost_end.bof) then
	If CInt(rs_cost_end("cnt"))>0 Then
		end_yn = "Y"
	End If
End If
rs_cost_end.close() : Set rs_cost_end = Nothing

'//비용구분
sql="select emp_etc_name from emp_etc_code where emp_etc_type = '70' order by emp_etc_code asc"
rs_etc.Open sql, Dbconn, 1
Set listCostCenter = getRsToDic(rs_etc)
rs_etc.close : Set rs_etc = nothing


view_sort = request("view_sort")

if view_sort = "" then
	view_sort = "ASC"
end if

order_Sql = " ORDER BY A.pmg_company " & view_sort

where_sql = " WHERE 1=1 " & chr(13)
where_sql = where_sql & " AND A.pmg_id=1 " & chr(13) ' 왜 이 라인에 주석을 잡았었지? 박정신부장 문의 2180-08-14 (비용현황관리/지원월별현황 에서 2건이상나오는 문제..)
where_sql = where_sql & " AND A.pmg_emp_no =  B.emp_no " & chr(13)
'where_sql = where_sql & " AND B.cost_except in ('0','1') " & chr(13)
where_sql = where_sql & " AND A.pmg_yymm = '" & srchEmpMonth & "' " & chr(13)
where_sql = where_sql & " AND B.emp_month ='" & srchEmpMonth & "' " & chr(13)

if view_condi <> "전체" then		'//회사명 검색	
    where_sql = where_sql & " AND A.pmg_company='" & view_condi & "' " & chr(13)
end if

'IF Trim(srchEmpName & "") <>"" Then		'//이름 검색
'where_sql = where_sql & " AND B.emp_name like '%" & srchEmpName & "%' " & chr(13)
'End IF
IF Trim(srchWord & "") <>"" Then		'//이름 검색
    where_sql = where_sql & " AND B." & srchCategory & " like '%" & srchWord & "%' " & chr(13)
End IF

Sql = "SELECT count(*) FROM pay_month_give  A ,emp_master_month B  " & where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
ELSE
	total_page = int((tottal_record / pgsize) + 1)
END If

sql = " SELECT  A.pmg_yymm           " & chr(13) & _
      "       , A.pmg_company        " & chr(13) & _
      "       , A.pmg_saupbu         " & chr(13) & _
      "       , A.pmg_give_total     " & chr(13) & _
      "       , A.cost_group         " & chr(13) & _
      "       , A.cost_center        " & chr(13) & _
      "       , B.emp_no             " & chr(13) & _
      "       , B.emp_name           " & chr(13) & _
      "       , B.emp_job            " & chr(13) & _
      "       , B.emp_type           " & chr(13) & _
      "       , B.emp_saupbu         " & chr(13) & _
      "       , B.emp_org_name       " & chr(13) & _
      "       , B.emp_company        " & chr(13) & _
      "       , B.emp_bonbu          " & chr(13) & _
      "       , B.emp_team           " & chr(13) & _
      "       , B.emp_reside_company " & chr(13) & _
      "       , B.emp_reside_place   " & chr(13) & _
      "    FROM pay_month_give A     " & chr(13) & _
      "       , emp_master_month B   " & chr(13) & _
      where_sql                        & chr(13) & _
      order_sql                        & chr(13) & _
      " LIMIT "& stpage & ", " & pgsize & chr(13)
Rs.Open Sql, Dbconn, 1
'Response.write "<pre>"&Sql&"</pre>"

title_line = " 직원 월별 현황 " & "(" & srchEmpMonth & ")"

%>
<!DOCTYPE html>
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
		</script>
		<script type="text/javascript">
			var empMonth = "<%=srchEmpMonth%>";
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

			function callbackOrgSelect(empNo, costGroup){
				var tEmpNo = "";
				$(".tableList tbody tr").each(function(){
					tEmpNo = $(this).find("td").eq(0).text();
					if( tEmpNo == empNo ){
						$(this).find("input[name='cost_group']").val(costGroup);
					}
				});
			}

            function changeEmpMasterMonth(empNo)
            {
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
						return false;
					}
				});

				if( empMonth==null || empMonth==""){ alert("년월 정보가 없습니다."); return false; }
				if( tEmpNo==null || tEmpNo==""){ alert("사번정보가 없습니다."); return false; }
				if( tCostCenter==null || tCostCenter==""){ alert("비용구분을 선택해주세요."); return false; }
				if( tCostGroup==null || tCostGroup==""){ alert("비용그룹을 입력해주세요."); return false; }

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
                             };
				$.ajax({
 					 url: "cost_emp_master_month_mg_save.asp"
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
							  Sql="select * from emp_org_mst where (org_level = '회사') ORDER BY org_code ASC"
	                          rs_org.Open Sql, Dbconn, 1	
							  %>
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
								<!--<label for="srchEmpName"><strong>이름 : </strong></label>-->
								<select name="srchCategory" id="srchCategory">
									<option value="emp_name"<% If srchCategory="emp_name" Then Response.write " selected=""selected"""%>>이름</option>
									<option value="emp_no"<% If srchCategory="emp_no" Then Response.write " selected=""selected"""%>>사번</option>
								</select>
								<input type="text" name="srchWord" id="srchWord" style="width: 100px; text-align: left; -ms-ime-mode: active;" value="<%=srchWord%>" />
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>                      

				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="16%" >
							<col width="42%" >
							<col width="8%" >
							<col width="22%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">소속</th>
								<th scope="col">조직</th>
								<th scope="col">비용구분</th>
								<th scope="col">비용그룹</th>
							</tr>
						</thead>
					<tbody>
						<%
						int j=0
						do until rs.eof
							
							tEmpNo 				= rs("emp_no")
							tCostCenter 		= rs("cost_center")
							tCostGroup 			= rs("cost_group")
							tEmpOrgName 		= rs("emp_org_name")
							tEmpCompany 		= rs("emp_company")
							tEmpBonbu 			= rs("emp_bonbu")
							tEmpSaupbu 			= rs("emp_saupbu")
							tEmpTeam 			= rs("emp_team")
							tEmpResideCompany	= rs("emp_reside_company")
							tEmpResidePlace 	= rs("emp_reside_place")
							
							j=j+1
						    %>
							<tr>
								<td class="first"><%=rs("emp_no")%></td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("emp_job")%></td>
								<td>
									<input type="hidden" id="emp_org_code<%=j%>" name="emp_org_code" value="<%=tEmpOrgCode%>" />
									<input id="emp_org_name<%=j%>" name="emp_org_name" type="text" style="width:120px" readonly="true" value="<%=tEmpOrgName%>">
                                    <!--a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="org"%>&mg_org=<%=mg_org%>&view_condi=<%=view_condi%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a-->
                                    <a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=<%="org"%>&mg_org=<%=mg_org%>&view_condi=<%=view_condi%>&org_id=<%=j%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a>
								</td>
								<td>
									<input type="text" 	 id="emp_company<%=j%>" 	    name="emp_company" 			readonly="true" value="<%=tEmpCompany%>"style="width:100px" />
									<input type="text" 	 id="emp_bonbu<%=j%>" 		    name="emp_bonbu" 			readonly="true" value="<%=tEmpBonbu%>"  style="width:120px" />
									<input type="text" 	 id="emp_saupbu<%=j%>" 		    name="emp_saupbu" 			readonly="true" value="<%=tEmpSaupbu%>" style="width:120px" />
									<input type="text" 	 id="emp_team<%=j%>" 		    name="emp_team" 			readonly="true" value="<%=tEmpTeam%>"   style="width:120px" />
									<input type="hidden" id="emp_reside_company<%=j%>"  name="emp_reside_company" 	readonly="true" value="<%=tEmpResideCompany%>" />
									<input type="hidden" id="emp_reside_place<%=j%>"    name="emp_reside_place" 	readonly="true" value="<%=tEmpResidePlace%>"   />
									<input type="hidden" id="emp_org_level<%=j%>" 	    name="emp_org_level" 		readonly="true" value="" />
									<input type="hidden" id="emp_type<%=j%>" 			name="emp_type">
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
                                                        <option value='<%=row("emp_etc_name")%>' <%If tCostCenter = row("emp_etc_name") then %>selected<% end if %>><%=row("emp_etc_name")%></option>
                                                        <%
													Next
												End If
											End If
										%>  
									</select>
								</td>
               	                <td>
									<input type="text" id="cost_group<%=j%>" name="cost_group" value="<%=tCostGroup%>" readonly="readonly" />
									<a href="#" class="btnType03" onClick="pop_Window('insa_org_select.asp?gubun=costEmp&target=<%=rs("emp_no")%>','orgselect','scrollbars=yes,width=800,height=400')">선택</a>
									<%' If end_yn<>"Y" Then %>
									    <a href="#" class="btnType04" onClick="changeEmpMasterMonth('<%=tEmpNo%>')">적용</a>
									<%' End If %>
								</td>
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
					<!--
                    <a href="insa_excel_emp.asp?view_condi=<%=view_condi%>" class="btnType04">엑셀다운로드</a>
					-->
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="<%=be_pg%>?page=<%=first_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="<%=be_pg%>?page=<%=intstart -1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="<%=be_pg%>?page=<%=i%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="<%=be_pg%>?page=<%=intend+1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[다음]</a> <a href="<%=be_pg%>?page=<%=total_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&srchEmpMonth=<%=srchEmpMonth%>&srchCategory=<%=srchCategory%>&srchWord=<%=srchWord%>&srchEmpName=<%=srchEmpName%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
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

