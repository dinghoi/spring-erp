<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
    dim saupbu_tab(10,2)

    for i = 1 to 10
        saupbu_tab(i,1) = ""
        saupbu_tab(i,2) = 0
    next

    ck_sw=Request("ck_sw")
        
    If ck_sw = "y" Then
        cost_month=Request("cost_month")
        saupbu = Request("saupbu")
    else
        cost_month=Request.form("cost_month")
        saupbu = Request.form("saupbu")
    End if


    if cost_month = "" then
        before_date = dateadd("m",-1,now())
        cost_month = mid(cstr(before_date),1,4) + mid(cstr(before_date),6,2)
    end If

    cost_year = mid(cost_month,1,4)
    cost_mm = mid(cost_month,5)

	sql = "    SELECT saupbu          /* 사업부 명 */    " & chr(13) &_
	      "         , saupbu_person   /* 사업부 인력 */  " & chr(13) &_
	      "         , tot_person      /* 총인력 */       " & chr(13) &_
	      "         , saupbu_per      /* 차지율 */       " & chr(13) &_
	      "         , saupbu_cost_amt /* 전사공통비1 */  " & chr(13) &_

          "         , saupbu_sale     /* 사업부 매출 */  " & chr(13) &_
          "         , tot_sale        /* 총 매출 */      " & chr(13) &_
          "         , sale_per        /* 차지율 */       " & chr(13) &_
          "         , saupbu_sale_amt /* 전사공통비2 */  " & chr(13) &_

	      "         , tot_cost_amt                       " & chr(13) &_
	      "      FROM management_cost                    " & chr(13) &_
	      "     WHERE cost_month ='"&cost_month&"'       " & chr(13) &_
	      "  GROUP BY saupbu                             " & chr(13) &_
	      "  ORDER BY saupbu                             "
	rs.Open sql, Dbconn, 1
'Response.write "<pre>"&sql&"</pre><br>"
    if saupbu = "" then
        if rs.eof then
            saupbu = ""
        else
            saupbu = rs("saupbu")
        end if
    end if

    title_line = "공통비 인원 및 매출 배분 기준 현황"
'	Response.write sql
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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

			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("발생년월을 입력하세요.");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<h3 class="stit">1. 전사공통비 배부 기준은 사업부별 손익에는 인원수, 고객사별손익은 해당 사업부내의 매출액 비율로 배부함. </h3>
				<h3 class="stit">2. 고객사별손익에 직접비 배분은 사업부내의 매출액 비율로 배부함. </h3>
				<form action="management_cost_report2.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>발생년월&nbsp;</strong>(예201401) : 
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="52%" height="356" valign="top">
				      <h3 class="stit">* 사업부별 인원 현황 및 비율</h3>
                      <table cellpadding="0" cellspacing="0" class="tableList">
                            <colgroup>
                                <col width="*" >
                                <col width="12%" >
                                <col width="12%" >
                                <col width="12%" >
                                <col width="12%" >

                                <col width="12%" >
                                <col width="12%" >
                                <col width="12%" >
                            </colgroup>
				        <thead>
                            <tr>
                                <th class="first" scope="col" rowspan="2">사업부</th>
                                <th scope="col" colspan="4" style="border-bottom:1px solid #e3e3e3;">전사공통비(인원)</th>
                                <th scope="col" colspan="3" style="border-bottom:1px solid #e3e3e3;">전사공통비(매출)</th>
                            </tr>
                            <tr>
                                <th scope="col" style="border-left:1px solid #e3e3e3;">사업부인력</th>
                                <th scope="col">차지율(%)</th>
                                <th scope="col">전사공통비</th>
                                <th scope="col">직접비</th>

                                <th scope="col">사업부매출</th>
                                <th scope="col">차지율(%)</th>
                                <th scope="col">전사공통비</th>
                            </tr>
			            </thead>
			            <tbody>
						<%
						tot_saupbu_person   = 0
						tot_saupbu_cost_amt = 0
						tot_saupbu_per      = 0
						tot_saupbu_direct   = 0
                        
                        tot_saupbu_sale     = 0
                        tot_sale_per        = 0
                        tot_saupbu_sale_amt = 0

						i = 0
                        do until rs.eof
							i = i + 1
'select saupbu,company,sum(cost_amt) as cost 
'from saupbu_sales 
'where substring(sales_date,1,7) = '2019-06' 
'  and saupbu = '"&saupbu&"'           " & chr(13) &_
'group by saupbu,company  
                            sql = "select sum(cost_amt_"&cost_mm&")      " & chr(13) &_
                                    "  from company_cost                   " & chr(13) &_
                                    " where (cost_center = '직접비' )      " & chr(13) &_
                                    "   and (saupbu = '"&rs("saupbu")&"' ) " & chr(13) &_
                                    "   and cost_year ='"&cost_year&"'     "
                            'Response.write "<pre>"&sql&"</pre><br>"                               
							set rs_etc=dbconn.execute(sql)

							if rs_etc(0) = "" or isnull(rs_etc(0)) then
								direct_cost = 0
							  else
								direct_cost = Cdbl(rs_etc(0))
							end if
							rs_etc.close()
							saupbu_tab(i,1) = rs("saupbu")
							saupbu_tab(i,2) = direct_cost

							tot_saupbu_person   = tot_saupbu_person + Cdbl(rs("saupbu_person"))
							tot_saupbu_cost_amt = tot_saupbu_cost_amt + Cdbl(rs("saupbu_cost_amt"))
							tot_saupbu_per      = tot_saupbu_per + rs("saupbu_per")
							tot_saupbu_direct   = tot_saupbu_direct + direct_cost

                            tot_saupbu_sale     = tot_saupbu_sale + rs("saupbu_sale")
                            tot_sale_per        = tot_sale_per + rs("sale_per")
                            tot_saupbu_sale_amt = tot_saupbu_sale_amt + rs("saupbu_sale_amt")
                        %>
				        <tr>
                            <td class="first"><a href="management_cost_report2.asp?saupbu=<%=rs("saupbu")%>&cost_month=<%=cost_month%>&ck_sw=<%="y"%>"><%=rs("saupbu")%></a></td>
                            <td class="right"><%=formatnumber(rs("saupbu_person"),0)%>&nbsp;</td>
                            <td class="right"><%=formatnumber(rs("saupbu_per")*100,3)%>%&nbsp;</td>
                            <td class="right"><%=formatnumber(rs("saupbu_cost_amt"),0)%>&nbsp;</td>
                            <td class="right"><%=formatnumber(direct_cost,0)%>&nbsp;</td>

                            <td class="right"><%=formatnumber(rs("saupbu_sale"),0)%>&nbsp;</td>
                            <td class="right"><%=formatnumber(rs("sale_per")*100,3)%>%&nbsp;</td>
                            <td class="right"><%=formatnumber(rs("saupbu_sale_amt"),0)%>&nbsp;</td>
			            </tr>
				        <%
							rs.movenext()
						loop
						rs.close()
						%>
				        <tr bgcolor="#FFE8E8">
                            <td class="first">계</td>
                            <td class="right"><%=formatnumber(tot_saupbu_person,0)%>&nbsp;</td>
                            <td class="right"><%=formatnumber(tot_saupbu_per*100,3)%>%&nbsp;</td>
                            <td class="right"><%=formatnumber(tot_saupbu_cost_amt,0)%>&nbsp;</td>
                            <td class="right"><%=formatnumber(tot_saupbu_direct,0)%>&nbsp;</td>

                            <td class="right"><%=formatnumber(tot_saupbu_sale,0)%>&nbsp;</td>
                            <td class="right"><%=formatnumber(tot_sale_per*100,3)%>%&nbsp;</td>
                            <td class="right"><%=formatnumber(tot_saupbu_sale_amt,0)%>&nbsp;</td>
			            </tr>
			            </tbody>
			          </table>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="46%" valign="top">
				        <h3 class="stit">* 사업부내 회사별 매출액 비율</h3>
				        <table cellpadding="0" cellspacing="0" summary="" class="tableList">
				        <colgroup>
				          <col width="20%" >
				          <col width="*" >
				          <col width="20%" >
			            </colgroup>
				        <thead>
                            <tr>
                                <th class="first" scope="col">사업부</th>
                                <th scope="col">고객사</th>
                                <th scope="col">매출</th>
                            </tr>
                        </thead>
			            <tbody>
                            <%
                            tot_cost_amt = 0
                            tot_charge_per = 0
                            tot_company_cost = 0

                            salesDate = LEFT (cost_month, 4) & "-" & RIGHT (cost_month, 2)
                            sql = "    SELECT saupbu, company, sum(cost_amt) as cost_amt   " & chr(13) &_
                                  "      FROM saupbu_sales                                 " & chr(13) &_
                                  "     WHERE substring(sales_date,1,7) = '"&salesDate&"'  " & chr(13) &_
                                  "       AND saupbu ='"&saupbu&"'                         " & chr(13) &_
                                  "  GROUP BY saupbu ,company                              " 
'Response.write "<pre>"&sql&"</pre><br>"
                            rs.Open sql, Dbconn, 1
                            do until rs.eof
                                tot_cost_amt = tot_cost_amt + rs("cost_amt")
                                %>
                                <tr>
                                    <td class="first"><%=rs("saupbu")%></td>
                                    <td><%=rs("company")%>&nbsp;</td>
                                    <td class="right"><%=formatnumber(rs("cost_amt"),0)%>&nbsp;</td>
                                </tr>
                                <%
                                rs.movenext()
                            loop
                            rs.close()
                            %>
                            <tr bgcolor="#FFE8E8">
                                <td class="first">계</td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><%=formatnumber(tot_cost_amt,0)%>&nbsp;</td>
                            </tr>
			            </tbody>
			            </table>
			          </td>
			        </tr>
				    <tr>
				      <td width="46%">&nbsp;</td>
				      <td width="2%">&nbsp;</td>
				      <td width="52%">&nbsp;</td>
			        </tr>
			      </table>
				     
			          </td>
			        </tr>
				    <tr>
				      <td width="46%">&nbsp;</td>
				      <td width="2%">&nbsp;</td>
				      <td width="52%">&nbsp;</td>
			        </tr>
			      </table>
                </div>

			</div>				
	</div>        				
	</body>
</html>

