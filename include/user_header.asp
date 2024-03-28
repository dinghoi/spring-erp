<!--#include virtual="/include/google_analytics.asp" -->
			<div id="header">
				<h1><img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/></h1>
				<h2><img src="/image/main_title.gif" alt="A/S 관리 시스템" width="198" height="25"/></h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%></strong>님 안녕하세요.
					</p>
					<a href="#" onclick="javascript:pop_user_mod();"><img src="/image/user_mod.gif" alt="개인정보변경"/></a>
                    <a href="logout.asp"><img src="/image/logout.gif" alt="로그아웃"/></a>
				</div>	
                <div id="gnb">
                    <ul>
                        <li class="dep1"><a href="as_list_ce_user.asp">A/S 관리</a></li>    
                        <li class="dep1">  
						
                    <%	if team <> "외주관리" then	%>						
                        <a href="waiting.asp?pg_name=<%="day_sum_user.asp"%>">현황관리</a>
					<%	  else	%>
                        <a>현황관리</a>
						<%	end if	%>
                         <li>
						 <%	if user_name <> "롯데렌탈" then	%>						
                        <a href="waiting.asp?pg_name=<%="nkp_main2.asp"%>">업무공유</a>
					<%else%>
                        <a>업무공유</a>
						<%	end if	%>
						</li>
                    </ul>
                </div>
			</div>
