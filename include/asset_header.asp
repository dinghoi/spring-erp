<!--#include virtual="/include/google_analytics.asp" -->
			<div id="header">
				<h1><img src="/image/com_logo.jpg" alt="Ȩ������" width="116" height="30"/></h1>
				<h2><img src="/image/asset_title.gif" alt="�ڻ� ���� �ý���" width="225" height="25"/></h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%></strong>�� �ȳ��ϼ���.
					</p>
					<a href="#" onclick="javascript:pop_user_mod();"><img src="/image/user_mod.gif" alt="������������"/></a>
                    <a href="logout.asp"><img src="/image/logout.gif" alt="�α׾ƿ�"/></a>
				</div>	
                <div id="gnb">
                    <ul>
                        <li class="dep1">                       
                    <%	if reside = "3" or reside = "0" then	%>						
                        <a href="as_list_asset.asp">A/S ����</a>
					<%	  else	%>
                        <a>A/S ����</a>
					<%	end if	%>
						</li>

                        <li class="dep1">                       
                    <%	if reside = "3" or reside = "0" then	%>						
                        <a href="waiting.asp?pg_name=<%="day_sum_asset.asp"%>">A/S ��Ȳ ����</a>
					<%	  else	%>
                        <a>A/S ��Ȳ ����</a>
					<%	end if	%>
						</li>

                        <li class="dep1">                       
                    <%	if reside = "2" or reside = "0" then	%>						
                        <a href="asset_process_mg.asp">�ڻ����</a>
					<%	  else	%>
                        <a>�ڻ����</a>
					<%	end if	%>
						</li>
                    </ul>
                </div>
			</div>
