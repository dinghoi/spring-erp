<!--#include virtual="/include/google_analytics.asp" -->
			<div id="header">
				<h1><img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/></h1>
				<h2><img src="/image/meterials_control_title.gif" alt="상품자재관리 시스템" width="198" height="25"/></h2>
				<div class="login">
					<p>
                    <strong><%=request.cookies("nkpmg_user")("coo_user_name")%>&nbsp;<%=request.cookies("nkpmg_user")("coo_user_grade")%></strong>님 안녕하세요.
					</p>
				</div>	
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_met_grade")
				%>  
                <div id="gnb">
                    <ul>
                        <li class="dep1">
					<% if in_empno = "900002" then	%>						
                        <a href="met_buy_ing_mg.asp">구매발주 관리</a>
					<%	 else	%>
                        <a>구매발주 관리</a>
                    <% end if	%>                          
                        <li class="dep1">
					<% if met_grade = "0" then	%>						
                      <% '  <a href="met_stock_in_mg.asp">입고 관리</a></li>    %>
                        <a href="met_stock_in_report01.asp">입고 관리</a></li>    
					<%	 else	%>
                        <a>입고 관리</a>
                    <% end if	%>            
                      <% ' <li class="dep1"><a href="met_stock_out_reg_mg.asp">출고 관리</a></li> %> 
                        <li class="dep1"><a href="met_stock_out_reg_ing01.asp">출고 관리</a></li>
                        <li class="dep1">
					<% if in_empno = "900002" then	%>						
                        <a href="met_stock_move_reg_mg.asp">창고이동 관리</a></li> 
					<%	 else	%>
                        <a>창고이동 관리</a>
                    <% end if	%>                                    
                        <li class="dep1"><a href="met_stock_jaego_mg.asp">재고 관리</a></li>
                        <li class="dep1">
					<% if met_grade = "0" then	%>						
                        <a href="met_stock_pum_jaego_mg.asp">현황 및 출력</a></li>    
					<%	 else	%>
                        <a>현황 및 출력</a>
                    <% end if	%>     
                        <li class="dep1">               
					<% if met_grade = "0" then	%>						
                        <a href="met_goods_code_mg.asp">코드 관리</a></li>    
					<%	 else	%>
                        <a>코드 관리</a>
                    <% end if	%>                                                        
                       <% ' <li class="dep1"><a href="meterials_basic_mg.asp">기본정보 관리</a></li> 
                        '<li class="dep1"><a href="meterials_jaego_analysis_mg.asp">재고평가 관리</a></li>   
                        '<li class="dep1"><a href="meterials_system_popup.asp">현황/조회 관리</a></li>
						'<li class="dep1"><a href="meterials_control_mg.asp">구매발주 관리</a></li>     %>
                    </ul>
                </div>
			</div>
