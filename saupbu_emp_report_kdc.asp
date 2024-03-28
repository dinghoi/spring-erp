<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	 
cost_month=Request.form("cost_month")
sales_saupbu=Request.form("sales_saupbu")

if cost_month = "" then
	before_date = dateadd("m",-1,now())
	cost_month = mid(cstr(before_date),1,4) + mid(cstr(before_date),6,2)
	sales_saupbu = "전체"
end If
cost_date = mid(cstr(cost_month),1,4) + "-" + mid(cstr(cost_month),5,2) + "-01"
start_date = dateadd("m",-1,cost_date)
cost_year = mid(cost_month,1,4)

'sql = "select * from emp_master_month where emp_month = '"&cost_month&"' and mg_saupbu = '"&sales_saupbu&"' and (emp_end_date = '1900-01-01' or isnull(emp_end_date) or emp_end_date >= '"&cost_date&"') order by emp_bonbu, emp_saupbu, emp_team, emp_org_name, emp_reside_place, emp_reside_company, emp_name"

if sales_saupbu = "전체" then
	sql = "     SELECT emp_master_month.*                                      " & chr(13) &_
	      "          , pay_month_give.pmg_job_support                          " & chr(13) &_
	      "          , pay_month_give.pmg_give_total                           " & chr(13) &_
	      "       FROM emp_master_month                                        " & chr(13) &_
	      " INNER JOIN pay_month_give                                          " & chr(13) &_
	      "         ON (emp_master_month.emp_no = pay_month_give.pmg_emp_no)   " & chr(13) &_
	      "        AND (emp_master_month.emp_month = pay_month_give.pmg_yymm)  " & chr(13) &_
	      "      WHERE (emp_master_month.emp_month='"&cost_month&"')           " & chr(13) &_
	      "        AND (pmg_id = '1')                                          " & chr(13) &_
	      "   ORDER BY cost_except                                             " & chr(13) &_
	      "          , emp_bonbu                                               " & chr(13) &_
	      "          , emp_saupbu                                              " & chr(13) &_
	      "          , emp_team                                                " & chr(13) &_
	      "          , emp_org_name                                            " & chr(13) &_
	      "          , emp_reside_place                                        " & chr(13) &_
	      "          , emp_reside_company                                      " & chr(13) &_
	      "          , emp_name                                                "
		  '"        AND (emp_master_month.cost_center <> '손익제외')            " & chr(13) &_
else	
	sql = "     SELECT emp_master_month.*                                      " & chr(13) &_
	      "          , pay_month_give.pmg_job_support                          " & chr(13) &_
	      "          , pay_month_give.pmg_give_total                           " & chr(13) &_
	      "       FROM emp_master_month                                        " & chr(13) &_
	      " INNER JOIN pay_month_give                                          " & chr(13) &_
	      "         ON (emp_master_month.emp_no = pay_month_give.pmg_emp_no)   " & chr(13) &_
	      "        AND (emp_master_month.emp_month = pay_month_give.pmg_yymm)  " & chr(13) &_
	      "      WHERE (emp_master_month.emp_month='"&cost_month&"')           " & chr(13) &_
	      "        AND (emp_master_month.mg_saupbu = '"&sales_saupbu&"')       " & chr(13) &_
	      "        AND (pmg_id = '1')                                          " & chr(13) &_
	      "   ORDER BY cost_except                                             " & chr(13) &_
	      "          , emp_bonbu                                               " & chr(13) &_
	      "          , emp_saupbu                                              " & chr(13) &_
	      "          , emp_team                                                " & chr(13) &_
	      "          , emp_org_name                                            " & chr(13) &_
	      "          , emp_reside_place                                        " & chr(13) &_
	      "          , emp_reside_company                                      " & chr(13) &_
	      "          , emp_name                                                        "
	      '"        AND (emp_master_month.cost_center <> '손익제외')            " & chr(13) &_
end if
'Response.write "<pre>"&sql & "</pre><br>"
rs.Open sql, Dbconn, 1

title_line = "사업부별 인원 현황"
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
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("근무년월을 입력하세요.");
					return false;
				}	
				return true;
			}

			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}

            $(document).ready(function(){
                $("input[name=cost_except]").change(function(){
                    if ("<%=sales_saupbu%>" != "KDC사업부")  
                    {
                        alert("KDC사업부만 동작합니다!");
                        return ;
                    }
                    var emp_month = $(this).attr("emp_month"); // 
                    var emp_no    = $(this).attr("emp_no");    //
                    var chked     = $(this).is(":checked");    // 체크여부

                    // alert("emp_month= "+emp_month+", emp_no= "+emp_no);

                    $.ajax({
                             url: "ajax_set_empMasterMonth_costExcept.asp"
                            ,type: 'post'
                            ,data:  { "emp_month" : emp_month
                                    , "emp_no"    : emp_no
                                    , "chked"     : chked
                                    }
                            ,dataType: "json"
                            ,success: function(data){
        						var result = data.result;
        						if( result=="succ"){
        							if(chked)
                                    {                        
                                        alert("수익제외 설정!");
                                    }
                                    else
                                    {
                                        alert("수익제외 해제!");
                                    }
                                }
                            }
                            ,error: function(jqXHR, status, errorThrown){
                                alert("에러가 발생하였습니다.\n상태코드 : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
                            }
                    });                    
                });
            });
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				
                <form action="saupbu_emp_report_kdc.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>근무년월&nbsp;</strong>(예201401) : 
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
								</label>
                                <label>
								<strong>사업부 &nbsp;:</strong>
                                <select name="sales_saupbu" id="sales_saupbu" style="width:150px">
                                    <option value="전체" <% if sales_saupbu = "전체" then %>selected<% end if %>>전체</option>
                                    <% 
                                    'sql_org="select saupbu from sales_org order by sort_seq"
                                    sql_org="select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"
                                    rs_org.Open sql_org, Dbconn, 1

                                    do until rs_org.eof
                                        %>
                                            <option value='<%=rs_org("saupbu")%>' <%If rs_org("saupbu") = sales_saupbu  then %>selected<% end if %>><%=rs_org("saupbu")%></option>
                                        <%
                                        rs_org.movenext()  
                                    loop 
                                    rs_org.Close()
                                    %>
                                </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
                
                <table cellpadding="0" cellspacing="0" width="100%">
                <tr>
                    <td>
                    <DIV id="topLine2" style="width:1200px;overflow:hidden;">
                    <div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="*" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="4%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
                            <col width="2%" >
							<col width="1%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">본부</th>
								<th scope="col">사업부</th>
								<th scope="col">팀</th>
								<th scope="col">상주처</th>
								<th scope="col">상주회사</th>
								<th scope="col">사번</th>
								<th scope="col">사원명</th>
								<th scope="col">직위</th>
								<th scope="col">퇴사일</th>
								<th scope="col">비용구분</th>
								<th scope="col">관리본부</th>
								<th scope="col">급여총액</th>
								<th scope="col">야특근</th>
                                <th scope="col">손익 제외</th>
								<th scope="col"></th>
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
							<col width="3%" >
							<col width="*" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="4%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
                            <col width="2%" >
							<col width="1%" >
						</colgroup>
						<tbody>
						<%
						i = 0
						j = 0
						team_sum = 0
						team_overtim_sum = 0
						tot_sum = 0
						tot_overtime_sum = 0
						bi_team = "first"
						do until rs.eof
							if bi_team = "first" then
								bi_team = rs("emp_team")
							end if
							if bi_team <> rs("emp_team") then
                                %>
                                <tr bgcolor="#FFFFCC">
                                    <td colspan="2" class="first">소계</td>
                                    <td>인원수&nbsp;&nbsp;<%=j%></td>
                                    <td><%=bI_team%>&nbsp;</td>
                                    <td colspan="8">&nbsp;</td>
                                    <td class="right">
                                    <% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
                                        <%=formatnumber(team_sum,0)%>
                                    <% else	%>
                                        ********
                                    <% end if	%>
                                    </td>
                                    <td class="right">
                                    <% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
                                        <%=formatnumber(team_overtime_sum,0)%>
                                    <% else	%>
                                        ********
                                    <% end if	%>
                                    </td>
                                    <td></td>
                                    <td></td>
                                </tr>
                                <%
								j = 0
								bi_team = rs("emp_team")								
								team_sum = 0
								team_overtime_sum = 0
							end if
							
                            ' 손익제외건은 누락 '2019.08.27
                            if  (rs("cost_except")<>"2") then 
                                i = i + 1
                                j = j + 1
                            end if 
						  	pmg_give_total = rs("pmg_give_total")
						  	pmg_job_support = rs("pmg_job_support")

							team_sum = team_sum + pmg_give_total
							team_overtime_sum = team_overtime_sum + pmg_job_support
							tot_sum = tot_sum + pmg_give_total
							tot_overtime_sum = tot_overtime_sum + pmg_job_support
                            %>
                            <tr>
                                <td class="first"><%=i%></td>
                                <td><%=rs("emp_bonbu")%>&nbsp;</td>
                                <td><%=rs("emp_saupbu")%>&nbsp;</td>
                                <td><%=rs("emp_team")%>&nbsp;</td>
                                <td><%=rs("emp_reside_place")%>&nbsp;</td>
                                <td><%=rs("emp_reside_company")%>&nbsp;</td>
                                <td><%=rs("emp_no")%></td>
                                <td><%=rs("emp_name")%></td>
                                <td><%=rs("emp_job")%></td>
                                <td><%=emp_end_date%>&nbsp;</td>
                                <td><%=rs("cost_center")%></td>
                                <td><%=rs("mg_saupbu")%>&nbsp;</td>
                                <td class="right">
                                <% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001,100952" then	%>
                                    <%=formatnumber(pmg_give_total,0)%>
                                <% else	%>
                                    ********
                                <% end if	%>
                                </td>
                                <td class="right">
                                <% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001,100952" then	%>
                                    <%=formatnumber(pmg_job_support,0)%>
                                <% else	%>
                                    ********
                                <% end if	%>
                                </td>
                                <td>
                                    <!-- 손익제외 여부를 표시 (2019.08.27) -->
                                    <input type="checkbox" name="cost_except" emp_month="<%=rs("emp_month")%>" emp_no="<%=rs("emp_no")%>"  <% if (rs("cost_except")="2") then %>checked<% end if %>>
                                </td>
                                <td></td>
                            </tr>
                            <%
							rs.movenext()
						loop
						%>
							<tr bgcolor="#FFFFCC">
								<td colspan="2" class="first">소계</td>
								<td>인원수&nbsp;&nbsp;<%=j%></td>
								<td><%=bI_team%>&nbsp;</td>
								<td colspan="8">&nbsp;</td>
								<td class="right">
                                <% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
                                    <%=formatnumber(team_sum,0)%>
                                <% else	%>
                                    ********
                                <% end if	%>
                                </td>
                                <td class="right">
                                <% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
                                    <%=formatnumber(team_overtime_sum,0)%>
                                <% else	%>
                                    ********
                                <% end if	%>
                                </td>
                                <td></td>
                            </tr>
                            <tr bgcolor="#FFE8E8">
                                <td colspan="2" class="first">총계</td>
                                <td>인원수&nbsp;&nbsp;<%=i%></td>
                                <td>&nbsp;</td>
                                <td colspan="8">&nbsp;</td>
                                <td class="right">
                                <% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
                                    <%=formatnumber(tot_sum,0)%>
                                <% else	%>
                                    ********
                                <% end if	%>
                                </td>
                                <td class="right">
                                <% if (position = "사업부장" and sales_saupbu = saupbu) or user_id = "900001" then	%>
                                    <%=formatnumber(tot_overtime_sum,0)%>
                                <% else	%>
                                    ********
                                <% end if	%>
								</td>
								<td></td>
                                <td></td>
							</tr>
						</tbody>
						</table>
                        </DIV>
						</td>
                    </tr>
					</table>
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td width="25%">
                            <div class="btnCenter">
                            <a href="saupbu_emp_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">엑셀다운로드</a>
                            </div>                  
                        </td>
                        <td width="50%"></td>
                        <td width="25%"></td>
                    </tr>
				    </table>
			    </form>
			    <br>
		    </div>				
	    </div>        				
	</body>
</html>

