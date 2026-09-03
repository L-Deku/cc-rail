-- 在 192.168.2.13（2020 定额库 RecoData2020 + 2020 项目库）上创建插件专用最小权限登录名。
-- 用法（SSMS 开启 SQLCMD 模式，或 sqlcmd -v PluginPassword="..."）：
--   :setvar PluginPassword "在这里填 reco_plugin 的密码（与 .213 相同）"
-- 密码不要写进本文件、不要提交到 Git。脚本可重复执行（幂等）。
-- 权限：RecoData2020 只读；现有全部用户库只读；model 预置只读（新建项目库自动继承）。本服务器上没有学习库，不授写权限。
SET NOCOUNT ON;

IF N'$(PluginPassword)' = N'' OR N'$(PluginPassword)' = N'$' + N'(PluginPassword)'
BEGIN
    RAISERROR(N'请先用 :setvar PluginPassword 提供 reco_plugin 的密码。', 16, 1);
    RETURN;
END;

IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = N'reco_plugin')
BEGIN
    CREATE LOGIN [reco_plugin] WITH PASSWORD = N'$(PluginPassword)', CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF, DEFAULT_DATABASE = [RecoData2020];
    PRINT N'已创建登录名 reco_plugin。';
END
ELSE
BEGIN
    ALTER LOGIN [reco_plugin] WITH PASSWORD = N'$(PluginPassword)';
    PRINT N'登录名 reco_plugin 已存在，已同步密码。';
END;

IF EXISTS (SELECT 1 FROM sys.server_role_members m JOIN sys.server_principals r ON r.principal_id = m.role_principal_id
           JOIN sys.server_principals p ON p.principal_id = m.member_principal_id WHERE p.name = N'reco_plugin' AND r.name <> N'public')
    RAISERROR(N'reco_plugin 不应属于任何服务器角色，请人工核对后移除。', 16, 1);

DECLARE @db sysname, @sql nvarchar(max);

DECLARE db_cursor CURSOR LOCAL FAST_FORWARD FOR
    SELECT name FROM sys.databases
    WHERE (database_id > 4 OR name = N'model') AND state_desc = N'ONLINE' AND is_read_only = 0
      AND name NOT IN (N'RecoLearning')
    ORDER BY CASE WHEN name = N'RecoData2020' THEN 0 WHEN name = N'model' THEN 1 ELSE 2 END, name;
OPEN db_cursor;
FETCH NEXT FROM db_cursor INTO @db;
WHILE @@FETCH_STATUS = 0
BEGIN
    BEGIN TRY
        SET @sql = N'USE ' + QUOTENAME(@db) + N';
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N''reco_plugin'')
    CREATE USER [reco_plugin] FOR LOGIN [reco_plugin];
EXEC sp_addrolemember N''db_datareader'', N''reco_plugin'';';
        EXEC sp_executesql @sql;
        PRINT N'已授权只读：' + @db;
    END TRY
    BEGIN CATCH
        PRINT N'失败：' + @db + N' - ' + ERROR_MESSAGE();
    END CATCH;
    FETCH NEXT FROM db_cursor INTO @db;
END;
CLOSE db_cursor; DEALLOCATE db_cursor;

PRINT N'完成。请再执行 Diagnose-RecoLogins.sql 核对结果。';
