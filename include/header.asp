<!--#include virtual="/include/google_analytics.asp" -->
			<div id="header">
				<h1><!--<img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/>-->
					<img src="/images/logo_long_ko_main.png" alt="K-ONE" width="120" height="40"/>
				</h1>
				<h2>
					<div style="margin:6px;">
						<!--<img src="/image/main_title.gif" alt="A/S 관리 시스템" width="225" height="25"/>-->
						<img src="/image/main_title.gif" alt="A/S 관리 시스템" width="198" height="22"/>
					</div>
				</h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%>님</strong> 안녕하세요.
					</p>
				</div>
                <div id="gnb">
                    <ul>
                        <li class="dep1"><a href="/as_list_ce.asp">A/S 관리</a></li>
                        <li class="dep1">
					<%If c_grade < "1" Then	%>
                        <a href="/large_data_up.asp">다량처리</a>
					<%Else	%>
                        <a>다량처리</a>
                    <%End If%>
                        </li>
                        <li class="dep1">
                    <%If c_grade < "2" Then	%>
                        <a href="/waiting.asp?pg_name=<%="day_sum.asp"%>">총괄현황</a>
                    <%Else	%>
                        <a>총괄현황</a>
                    <%End If	%>
                        </li>
                        <li class="dep1">
                    <%If c_grade < "2" Then	%>
                        <a href="/waiting.asp?pg_name=<%="area_term_pro.asp"%>">통계현황</a>
					<%Else%>
                        <a>통계현황</a>
                    <%End If%>
                        </li>
                        <li class="dep1">
                    <%If c_grade < "1" Then	%>
                        <a href="/ce_mg_list.asp">CE 관리</a></li>
					<%Else	%>
                        <a>CE 관리</a>
					<%End If%>
                        </li>
                        <li class="dep1">
                    <%If c_grade < "1" Then	%>
                        <a href="/etc_code_mg.asp">코드 관리</a>
					<%Else	%>
                        <a>코드 관리</a>
                    <%End If	%>
                        </li>
                    </ul>
                </div>
			</div>
