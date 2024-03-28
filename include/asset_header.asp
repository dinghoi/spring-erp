<!--#include virtual="/include/google_analytics.asp" -->
			<div id="header">
				<h1><img src="/image/com_logo.jpg" alt="홈페이지" width="116" height="30"/></h1>
				<h2><img src="/image/asset_title.gif" alt="자산 관리 시스템" width="225" height="25"/></h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%></strong>님 안녕하세요.
					</p>
					<a href="#" onclick="javascript:pop_user_mod();"><img src="/image/user_mod.gif" alt="개인정보변경"/></a>
                    <a href="logout.asp"><img src="/image/logout.gif" alt="로그아웃"/></a>
				</div>	
                <div id="gnb">
                    <ul>
                        <li class="dep1">                       
                    <%	if reside = "3" or reside = "0" then	%>						
                        <a href="as_list_asset.asp">A/S 관리</a>
					<%	  else	%>
                        <a>A/S 관리</a>
					<%	end if	%>
						</li>

                        <li class="dep1">                       
                    <%	if reside = "3" or reside = "0" then	%>						
                        <a href="waiting.asp?pg_name=<%="day_sum_asset.asp"%>">A/S 현황 관리</a>
					<%	  else	%>
                        <a>A/S 현황 관리</a>
					<%	end if	%>
						</li>

                        <li class="dep1">                       
                    <%	if reside = "2" or reside = "0" then	%>						
                        <a href="asset_process_mg.asp">자산관리</a>
					<%	  else	%>
                        <a>자산관리</a>
					<%	end if	%>
						</li>
                    </ul>
                </div>
			</div>
