

<head>
<link rel="stylesheet" type="text/css" href="main.css">

<%
'Option Explicit
'response.expires=0
'on error resume next

dim r_title,r_url,r_adcontent,r_masteremail,r_showlogintype,r_badword                                     '##��������������Ϣ
    r_title          = "�齭��������"                                                             '##�����ұ���
    r_url            = "http://office.11k.net/"                                                             '##��ҳ·��
    r_adcontent      = "<font color=#cccccc><font class=3dfont >�� ӭ �� �� �� �� �� �� �� �� ��</font></font>"          '##��棬��ʹ��html����
    r_masteremail    = "admin@semmw.com"                                                                  '##����Ա����,������Ϊ�������
    r_showlogintype  = 1                                                                                  '##�Ƿ���ʾ�����ҽ��뷽ʽ 1-��ʾ 0-����ʾ
    r_badword        = "����|����|����|��|�ְ�|үү|��|��|��|��|��|fuck|shit"                                                              '##Ҫ���˵ķǷ��ַ����м���"|"���
    r_refresh        = "up"                                                                             '##ˢ�·�ʽ down ���£�up ����


dim m_fontsize,m_text,m_text1,m_text2,m_text3,m_bg,m_bg1,m_bg2,m_bg3  '##��ҳ����ɫ���ֺ��趨
    m_fontsize       = "9pt"                                                                              '##�����С
    m_text           = "#000000"                                                                          '##������ɫ
    m_text1          = "#a0s0ff"                                                                          '##���������ɫһ
    m_text2          = "#FFFFFF"                                                                          '##���������ɫ��
    m_text3          = "#000000"                                                                          '##���������ɫ��
    m_bg             = "#FFFFFF"                                                                          '##��ҳ�汳����ɫ
    m_bg1            = "#FFFFFF"                                                                          '##��񱳾�ɫһ
    m_bg2            = "#6666CC"                                                                          '##��񱳾�ɫ��
    m_bg3            = "#FFFFFF"                                                                          '##��񱳾�ɫ��

dim m_button,m_buttonover,m_buttondown                                                                    '##��ť��ʽ

    m_button         = "background-color: #76ee00; color: #000000; text-align: center; vertical-align: middle; border: #FFFFFF 1px solid"
    m_buttonover     = "background-color: #7d26cd; color: #FFFFFF; text-align: center; vertical-align: middle; border: #000000 1px solid"
    m_buttondown     = "background-color: #76ee00; color: #FFFFFF; text-align: center; vertical-align: middle; border: #000000 1px solid"

dim level1rate,level2rate,level3rate,level4rate,level5rate,level6rate,level7rate,level8rate,level9rate    '##���ֱ�׼
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

