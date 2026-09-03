-- 只读诊断：列出 reco 与 reco_plugin 两个登录名的服务器角色、各库用户映射与库角色。
-- 在 192.168.2.213 与 192.168.2.13 各执行一次，建号脚本执行前后各跑一遍对照。
SET NOCOUNT ON;

SELECT @@SERVERNAME AS server_name, SUSER_SNAME() AS executed_as, GETDATE() AS checked_at;

-- 1. 登录名与服务器角色
SELECT p.name AS login_name, p.type_desc, p.is_disabled, p.create_date,
       STUFF((SELECT ', ' + r.name
              FROM sys.server_role_members m
              JOIN sys.server_principals r ON r.principal_id = m.role_principal_id
              WHERE m.member_principal_id = p.principal_id
              FOR XML PATH('')), 1, 2, '') AS server_roles
FROM sys.server_principals p
WHERE p.name IN (N'reco', N'reco_plugin')
ORDER BY p.name;

-- 2. 服务器级显式权限
SELECT p.name AS login_name, sp.permission_name, sp.state_desc
FROM sys.server_permissions sp
JOIN sys.server_principals p ON p.principal_id = sp.grantee_principal_id
WHERE p.name IN (N'reco', N'reco_plugin')
ORDER BY p.name, sp.permission_name;

-- 3. 每个库里的用户映射与库角色
DECLARE @sql nvarchar(max) = N'';
SELECT @sql = @sql + N'
SELECT N''' + REPLACE(d.name, '''', '''''') + N''' AS database_name, l.name AS login_name, u.name AS db_user,
       STUFF((SELECT '', '' + r.name
              FROM ' + QUOTENAME(d.name) + N'.sys.database_role_members rm
              JOIN ' + QUOTENAME(d.name) + N'.sys.database_principals r ON r.principal_id = rm.role_principal_id
              WHERE rm.member_principal_id = u.principal_id
              FOR XML PATH('''')), 1, 2, '''') AS db_roles
FROM ' + QUOTENAME(d.name) + N'.sys.database_principals u
JOIN sys.server_principals l ON l.sid = u.sid
WHERE l.name IN (N''reco'', N''reco_plugin'')
UNION ALL '
FROM sys.databases d
WHERE d.state_desc = N'ONLINE' AND d.name NOT IN (N'tempdb');

IF LEN(@sql) > 0
BEGIN
    SET @sql = LEFT(@sql, LEN(@sql) - LEN(N'UNION ALL ')) + N' ORDER BY database_name, login_name;';
    EXEC sp_executesql @sql;
END;

-- 4. 尚未给 reco_plugin 建用户的在线用户库（建号脚本执行后此列表应为空）
DECLARE @missing TABLE (database_name sysname);
DECLARE @check nvarchar(max) = N'';
SELECT @check = @check + N'
IF NOT EXISTS (SELECT 1 FROM ' + QUOTENAME(d.name) + N'.sys.database_principals WHERE name = N''reco_plugin'')
    SELECT N''' + REPLACE(d.name, '''', '''''') + N''';'
FROM sys.databases d
WHERE d.state_desc = N'ONLINE' AND d.database_id > 4;

IF LEN(@check) > 0
BEGIN
    INSERT INTO @missing (database_name) EXEC sp_executesql @check;
END;
SELECT database_name AS database_without_plugin_user FROM @missing ORDER BY database_name;
