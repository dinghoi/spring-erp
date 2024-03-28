<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim pummok_tab(4,20)
dim cost_tab(6,40)

slip_no = request("slip_no")
title_line = "전표 History"

Sql="select * from sales_slip where slip_no = '"&slip_no&"' order by slip_id, slip_seq"
Rs.Open Sql, Dbconn, 1

if rs("slip_id") = "1" then	
	view_slip_id = "대기전표"
  else
	view_slip_id = "수주전표"
end if

if rs("slip_stat") = "2" then	
	slip_stat_view = "부분수주"
  elseif rs("slip_stat") = "3" then	
	slip_stat_view = "수주완료"
  elseif rs("slip_stat") = "4" then	
	slip_stat_view = "취소진행"
  elseif rs("slip_stat") = "5" then	
	slip_stat_view = "전표취소"
  else
	slip_stat_view = "기초전표"
end if

view_att_file = rs("att_file")
if rs("sales_yn") = "Y" then	
	view_sales_yn = "매출"
  else
	view_sales_yn = "비매출"
end if
if rs("bill_issue_yn") = "Y" then	
	view_bill_issue_yn = "발행"
  else
	view_bill_issue_yn = "미발행"
end if

slip_memo = rs("slip_memo")
cancel_memo = slip_memo
if slip_memo = "" or isnull(slip_memo) then
	slip_memo = rs("slip_memo")
  else
	slip_memo = replace(slip_memo,chr(10),"<br>")
end if

buy_cost = rs("buy_cost")
sales_cost = rs("sales_cost")
sales_cost_vat = rs("sales_cost_vat")
sales_price = rs("sales_price")
margin_cost = rs("margin_cost")
if rs("sales_cost") = 0 then
	margin_per = 0
  else
	margin_per = rs("margin_cost")/rs("sales_cost") * 100
end if
view_att_file = rs("att_file")
sign_yn = rs("sign_yn")
path = "/sales_file"

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
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
						
				{
				a=confirm('전표을 취소하겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = false; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 13; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
                factory.printing.rightMargin = 13; //오른쯕 여백 설정
                factory.printing.bottomMargin = 15; //바닦 여백 설정
        //		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
        //		factory.printing.printer = ""; //프린터 할 프린터 이름
        //		factory.printing.paperSize = "A4"; //용지선택
        //		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
        //		factory.printing.collate = true; //순서대로 출력하기
        //		factory.printing.copies = "1"; //인쇄할 매수
        //		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
        //		factory.printing.Printer(true); //출력하기
                factory.printing.Preview(); //윈도우를 통해서 출력
                factory.printing.Print(false); //윈도우를 통해서 출력
            }
//			function approve_request(slip_id,slip_no,slip_seq) 
			function approve_request() 
				{
				a=confirm('결재 요청하시겠습니까?')
				if (a==true) {
//					document.frm.action = "sales_slip_approve_ok.asp?slip_id="+slip_id+'&slip_no='+slip_no+'&slip_seq='+slip_seq;
					document.frm.action = "sales_slip_approve_ok.asp";
					document.frm.submit();
				}
				return false;
				}
        </script>
	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>전표번호</th>
							  <td class="left"><%=slip_no%></td>
							  <th>매출조직</th>
							  <td class="left"><%=rs("sales_company")%>&nbsp;<%=rs("sales_company")%></td>
							  <th>영업사원</th>
							  <td class="left"><%=rs("emp_name")%>&nbsp;<%=rs("org_name")%></td>
						    </tr>
							<tr>
							  <th>거래처</th>
							  <td class="left"><%=rs("trade_name")%></td>
							  <th>사업자번호</th>
							  <td class="left"><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%></td>
							  <th>거래처<br>
						      담당자</th>
							  <td class="left"><%=rs("trade_person")%>&nbsp;</td>
                          </tr>
							<tr>
							  <th>연락처</th>
							  <td class="left"><%=rs("trade_person_tel_no")%>&nbsp;</td>
							  <th>계산서 메일</th>
							  <td class="left"><%=rs("trade_email")%></td>
							  <th>매출구분</th>
							  <td class="left"><%=view_sales_yn%></td>
                          </tr>
							<tr>
							  <th>매출일자</th>
							  <td class="left"><%=rs("sales_date")%></td>
							  <th>제품출고<br>
요청일</th>
							  <td class="left"><%=rs("out_request_date")%></td>
							  <th>계산서<br>
발행여부</th>
							  <td class="left"><%=view_bill_issue_yn%></td>
						    </tr>
							<tr>
							  <th>계산서<br>
발행예정일</th>
							  <td class="left"><%=rs("bill_due_date")%></td>
							  <th>계산서발행일</th>
							  <td class="left"><%=rs("bill_issue_date")%>&nbsp;</td>
							  <th>수금상태</th>
							  <td class="left"><%=rs("collect_stat")%></td>
						    </tr>
							<tr>
							  <th>수금방법</th>
							  <td class="left"><%=rs("bill_collect")%></td>
							  <th>수금예정일</th>
							  <td class="left"><%=rs("collect_due_date")%></td>
							  <th>수금일</th>
							  <td class="left"><%=rs("collect_date")%>&nbsp;</td>
						    </tr>
							<tr>
							  <th>비고</th>
							  <td colspan="5" class="left"><%=slip_memo%></td>
						    </tr>
						</tbody>
					</table>
				<h3 class="stit">* 전표 History</h3>
           		<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="8%" >
							<col width="12%" >
							<col width="*" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">전표유형</th>
								<th scope="col">전표번호</th>
								<th scope="col">전표상태</th>
								<th scope="col">매출일자</th>
								<th scope="col">계산서<br>발행일</th>
								<th scope="col">계산서<br>발행예정일</th>
								<th scope="col">매입총액</th>
								<th scope="col">매출총액</th>
								<th scope="col">마진총액</th>
							</tr>
						</thead>
						<tbody>
			  				<tr bgcolor="#FFE8E8">
								<td colspan="2" class="first"><%=view_slip_id%></td>
								<td>
								<a href="#" onClick="pop_Window('sales_slip_view.asp?slip_id=<%=rs("slip_id")%>&slip_no=<%=rs("slip_no")%>&slip_seq=<%=rs("slip_seq")%>','sales_slip_view_pop','scrollbars=yes,width=1230,height=600')"><%=rs("slip_no")%>-<%=rs("slip_seq")%>
                                </td>
								<td><%=slip_stat_view%></td>
								<td><%=rs("sales_date")%>&nbsp;</td>
								<td><%=rs("bill_issue_date")%>&nbsp;</td>
								<td><%=rs("bill_due_date")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("buy_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("margin_cost"),0)%></td>
							</tr>
						<%
						i = 0
						rs.movenext()
						do until rs.eof
							if rs("slip_id") = "1" then	
								view_slip_id = "대기전표"
							  else
								view_slip_id = "수주전표"
							end if
							if rs("slip_stat") = "2" then	
								slip_stat_view = "부분수주"
							  elseif rs("slip_stat") = "3" then	
								slip_stat_view = "전체수주"
							  elseif rs("slip_stat") = "4" then	
								slip_stat_view = "취소진행"
							  elseif rs("slip_stat") = "5" then	
								slip_stat_view = "전표취소"
							  else
								slip_stat_view = "기초전표"
							end if
							i = i + 1
						%>
			  				<tr>
								<td class="first"><%=i%></td>
								<td><%=view_slip_id%></td>
								<td>
								<a href="#" onClick="pop_Window('sales_slip_view.asp?slip_id=<%=rs("slip_id")%>&slip_no=<%=rs("slip_no")%>&slip_seq=<%=rs("slip_seq")%>&view_only=<%="Y"%>','sales_slip_view_pop','scrollbars=yes,width=1230,height=600')"><%=rs("slip_no")%>-<%=rs("slip_seq")%>
                                </td>
								<td><%=slip_stat_view%></td>
								<td><%=rs("sales_date")%>&nbsp;</td>
								<td><%=rs("bill_issue_date")%>&nbsp;</td>
								<td><%=rs("bill_due_date")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("buy_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("margin_cost"),0)%></td>
							</tr>
						<%
							rs.movenext()
						loop
						%>
						</tbody>
				  </table>                    
			    <br>
     				<div class="noprint">
                        <div align=center>
                            <span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>
                            <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                        </div>
					</div>
					<br>
					<input type="hidden" name="slip_no" value="<%=slip_no%>">
				</form>
                </div>
			</div>
		</div>
	</body>
</html>

