<%
Const MaxPageSize=20        '查看统计记录时，每页最多显示多少条记录
Const ExpireTime=24        '同一IP每隔多少时间后访问才继续计数，单位小时，默认为24小时
Const MaxRecord=300        '后台管理时显示多少条记录，默认为100条
Const OnlineTime=20        '在线人数截取时间，单位分钟，默认为20分钟
Const TimeZone=8        '服务器所在时区，中国为东8区，所以默认为8
Const Language="CHS"        '默认语言，默认为简体中文CHS
Const Skin =  "1"       '	系统默认风格，可选范围0－4
Const Sysmode =  "1|10|2"       '第一个参数默认为0，日ip小于1000的设置为0；大于1000以上的设置为1，默认自动清理10天没有访问且访问数据小于5次的内容。
Const RecordNum=100        '最后详细来访信息记录多少猹记录，默认为100条。因涉及到对数据库的操作，请登陆后台管理后修改此值，在此修改无效。
Const YVisitor=0        '原网站访问量
Const YPageView=0        '原网站浏览量
%>