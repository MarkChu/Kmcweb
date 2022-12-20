<div class="left-menu">
    <div><i class="fas fa-user"></i> 登入者：<%=session("username")%></div>
    <ul class="left-nav">
        <li><a href="?m=motion"><i class="fas fa-book-open"></i> 議案查詢</a></li>
        <%if session("weblaw")>0 then%>
        <li><a href="?m=weblaw1"><i class="fas fa-list"></i> 地方議政法規</a></li>
        <li><a href="?m=weblaw2"><i class="fas fa-list"></i> 相關法令解釋</a></li>
        <li><a href="?m=weblaw3"><i class="fas fa-list"></i> 財政預算法規</a></li>
        <li><a href="?m=weblaw4"><i class="fas fa-list"></i> 其他法規</a></li>
        <%end if%>
        <%if session("sms")>0 then%>
        <li><a href="?m=sms"><i class="fas fa-sms"></i> 簡訊發送</a></li>
        <%end if%>
        <li><a href="?m=news"><i class="fas fa-newspaper"></i> 線上新聞檢索</a></li>
        <%if session("account")>0 then%>
        <li><a href="?m=account"><i class="fas fa-chalkboard"></i> 會計專區</a></li>
        <%end if%>
        <li><a href="?logout=true"><i class="fas fa-sign-out-alt"></i> 登出</a></li>
    </ul>
</div>