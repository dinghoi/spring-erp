<!--#include virtual="/include/google_analytics.asp" -->
			<div id="header">
				<h1><!--<img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/>-->
					<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
				</h1>
				<h2>
					<div style="margin:6px;">
						<!--<img src="/image/cost_title.gif" alt="비용 관리 시스템" width="189" height="25"/>-->
						<img src="/image/cost_title.gif" alt="비용 관리 시스템" width="166" height="22"/>
					</div>
				</h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%>님</strong> 안녕하세요
					</p>
				</div>
                <div id="gnb">
                    <ul>
                        <li class="dep1"><a href="/person_cost_report.asp">비용등록관리</a></li>
                        <li class="dep1"><a href="/emp_cost_mg.asp">비용현황관리</a></li>
                        <li class="dep1">
                        <%
                        '최길성, 박정신, 강경진, 허정호, 하장호
                        If cost_grade = "0" AND (user_id = "100031" Or user_id = "100359" Or user_id = "101880" Or user_id = "102592" Or user_id = "100953") Then
                        %>
                            <a href="/cost_grade_mg.asp">비용코드관리</a>
                        <%Else	%>
                            <a>비용코드관리</a>
                        <%End If	%>
						</li>
                    </ul>
                </div>
			</div>
