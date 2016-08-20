<%
'Option Explicit
'response.expires=0
'on error resume next

dim strConnString '##连接数据库信息

    'strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;DBQ=" & server.MapPath("database.mdb") & ";uid=admin;PWD=;"         '## MS Access 2000
    'strConnString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & server.MapPath("mdbpath.mdb") & ";uid=admin;PWD=;"  '## MS Access
    'strConnString = "DRIVER={SQL Server};server=SERVER_NAME;uid=SQL_USER;pwd=PASSWORD;database=DATABASE_NAME"             '## MS SQL Server 7
    'strConnString = "DSN=DSN_name;UID=USER;PWD=PASSWORD"                                                                  '## Use DSN
    'my_Conn.open "DSN_name,User,Password"

set my_Conn = Server.CreateObject("ADODB.Connection")
    strConnString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & server.MapPath("3a4b5c.mdb") & ";uid=admin;PWD=;"
    my_Conn.Open strConnString

dim dbtable_change,dbtable_kill,dbtable_gbook,dbtable_user,dbtable_function        '##数据库结构

    dbtable_change         = "chat_change"                                         '##加分记录库
    dbtable_kill           = "chat_kill"                                           '##踢人记录库
    dbtable_gbook          = "chat_gbook"                                          '##留言记录库
    dbtable_user           = "userinfo"                                            '##用户信息库
    dbtable_function       = "function"                                            '##聊天动作库

dim dbfield_change_id,dbfield_change_change,dbfield_kill_id,dbfield_kill_kill      '##加分记录库，踢人记录库字段结构

dbfield_change_id      = "id"
dbfield_change_change  = "change"

dbfield_kill_id        = "id"
dbfield_kill_kill      = "kill"

dim dbfield_function_id,dbfield_function_command,dbfield_function_xiang           '##动作库字段结构

dbfield_function_id    = "id"
dbfield_function_cmd   = "cmd"
dbfield_function_xiang = "xiang"

dim dbfield_gook_id,dbfield_gook_name,dbfield_gook_lyname,dbfield_gook_email      '##留言记录库字段结构
dim dbfield_gook_homepage,dbfield_gook_addtime,dbfield_gook_message
dim dbfield_gook_comefrom,dbfield_gook_picture

dbfield_gbook_id       = "id"
dbfield_gbook_name     = "name"
dbfield_gbook_lyname   = "lyname"
dbfield_gbook_email    = "email"
dbfield_gbook_homepage = "homepage"
dbfield_gbook_addtime  = "addtime"
dbfield_gbook_message  = "message"
dbfield_gbook_comefrom = "comefrom"
dbfield_gbook_picture  = "picture"

dim dbfield_user_id,dbfield_user_username,dbfield_user_password                  '##用户信息库字段结构
dim dbfield_user_email,dbfield_user_oicq,dbfield_user_homepage
dim dbfield_user_comefrom,dbfield_user_rate,dbfield_user_ip
dim dbfield_user_lasttime,dbfield_user_sex,dbfield_user_manager

dbfield_user_id        = "id"
dbfield_user_username  = "username"
dbfield_user_password  = "password"
dbfield_user_email     = "d_email"
dbfield_user_oicq      = "d_oicq"
dbfield_user_homepage  = "d_homepage"
dbfield_user_comefrom  = "d_comefrom"
dbfield_user_rate      = "d_rate"
dbfield_user_ip        = "d_ip"
dbfield_user_lasttime  = "d_lasttime"
dbfield_user_sex       = "d_sex"
dbfield_user_manager   = "d_manager"
%>