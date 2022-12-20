<%
if request("logout")="true" then
    session.contents.removeall   
    session.abandon
    response.redirect "default.asp"
    response.end
end if

if trim(session("userid")&"")="" then
    session.contents.removeall   
    session.abandon
    response.redirect "default.asp"
    response.end
end if
Session.Timeout = "60"

if request("m")>"" then
    session("menu") = request("m")
else
    if session("menu")="" then
        session("menu") = "news"
    end if
end if

select case session("menu")
    case "motion"
        pagetitle = "議案查詢"
        page = "motion"
        redim submenu(3,2)
        submenu(0,0) = "議案查詢"
        submenu(1,0) = "motion"

        submenu(0,1) = "第十一至十四屆議決案暨處理情形"
        submenu(1,1) = "motion_history"

        submenu(0,2) = "第一至第十屆議決案暨處理情形"
        submenu(1,2) = "motion_history1"

    case "sms"
        pagetitle = "簡訊發送"
        page = "sms"
        redim submenu(3,2)
        submenu(0,0) = "簡訊首頁"
        submenu(1,0) = "sms"

        submenu(0,1) = "發送簡訊"
        submenu(1,1) = "sms_send"

        submenu(0,2) = "發送紀錄"
        submenu(1,2) = "sms_history"
    case "news"
        pagetitle = "線上新聞檢索"
        page = "news"
    case "account"
        pagetitle = "會計專區"
        page = "account"
        redim submenu(3,1)
        submenu(0,0) = "會計資料列表"
        submenu(1,0) = "account"

        submenu(0,1) = "新增/修改資料"
        submenu(1,1) = "account_add"

    case "weblaw1"
        pagetitle = "地方議政法規"
        page = "weblaw1"
        redim submenu(3,1)
        submenu(0,0) = "地方議政法規列表"
        submenu(1,0) = "weblaw1"

        submenu(0,1) = "新增/修改資料"
        submenu(1,1) = "weblaw1_add"

    case "weblaw2"
        pagetitle = "相關法令解釋"
        page = "weblaw2"
        redim submenu(3,1)
        submenu(0,0) = "相關法令解釋列表"
        submenu(1,0) = "weblaw2"

        submenu(0,1) = "新增/修改資料"
        submenu(1,1) = "weblaw2_add"

    case "weblaw3"
        pagetitle = "財政預算法規"
        page = "weblaw3"
        redim submenu(3,1)
        submenu(0,0) = "財政預算法規列表"
        submenu(1,0) = "weblaw3"

        submenu(0,1) = "新增/修改資料"
        submenu(1,1) = "weblaw3_add"

    case "weblaw4"
        pagetitle = "其他法規"
        page = "weblaw4"
        redim submenu(3,1)
        submenu(0,0) = "其他法規列表"
        submenu(1,0) = "weblaw4"

        submenu(0,1) = "新增/修改資料"
        submenu(1,1) = "weblaw4_add"        


end select

if request("p")>"" then
    page = request("p")
end if
%>