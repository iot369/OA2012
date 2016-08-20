

<head>
<link rel="stylesheet" type="text/css" href="main.css">

<%
'Option Explicit
'response.expires=0
'on error resume next

dim r_title,r_url,r_adcontent,r_masteremail,r_showlogintype,r_badword                                     '##设置聊天室主信息
    r_title          = "伴江行聊天室"                                                             '##聊天室标题
    r_url            = "http://office.11k.net/"                                                             '##主页路径
    r_adcontent      = "<font color=#cccccc><font class=3dfont >欢 迎 光 临 上 电 在 线 会 议 室</font></font>"          '##广告，可使用html语言
    r_masteremail    = "admin@semmw.com"                                                                  '##管理员信箱,可能设为你的信箱
    r_showlogintype  = 1                                                                                  '##是否显示聊天室进入方式 1-显示 0-不显示
    r_badword        = "你妈|你妈|他妈|爹|爸爸|爷爷|吊|逼|操|草|靠|fuck|shit"                                                              '##要过滤的非法字符，中间用"|"间隔
    r_refresh        = "up"                                                                             '##刷新方式 down 向下，up 向上


dim m_fontsize,m_text,m_text1,m_text2,m_text3,m_bg,m_bg1,m_bg2,m_bg3  '##主页面颜色及字号设定
    m_fontsize       = "9pt"                                                                              '##字体大小
    m_text           = "#000000"                                                                          '##字体颜色
    m_text1          = "#a0s0ff"                                                                          '##表格字体颜色一
    m_text2          = "#FFFFFF"                                                                          '##表格字体颜色二
    m_text3          = "#000000"                                                                          '##表格字体颜色三
    m_bg             = "#FFFFFF"                                                                          '##主页面背景颜色
    m_bg1            = "#FFFFFF"                                                                          '##表格背景色一
    m_bg2            = "#6666CC"                                                                          '##表格背景色二
    m_bg3            = "#FFFFFF"                                                                          '##表格背景色三

dim m_button,m_buttonover,m_buttondown                                                                    '##按钮样式

    m_button         = "background-color: #76ee00; color: #000000; text-align: center; vertical-align: middle; border: #FFFFFF 1px solid"
    m_buttonover     = "background-color: #7d26cd; color: #FFFFFF; text-align: center; vertical-align: middle; border: #000000 1px solid"
    m_buttondown     = "background-color: #76ee00; color: #FFFFFF; text-align: center; vertical-align: middle; border: #000000 1px solid"

dim level1rate,level2rate,level3rate,level4rate,level5rate,level6rate,level7rate,level8rate,level9rate    '##积分标准
    level1rate       = 0
    level2rate       = 299
    level3rate       = 599
    level4rate       = 1199
    level5rate       = 2399
    level6rate       = 3599
    level7rate       = 5999
    level8rate       = 9999
    level9rate       = 9999
%>
</head>

