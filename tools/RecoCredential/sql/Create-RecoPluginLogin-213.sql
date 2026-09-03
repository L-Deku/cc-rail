-- 在 192.168.2.213（学习库 RecoLearning + 2024 定额库 RecoData2024 + 2024 项目库）上创建插件专用最小权限登录名。
-- 用法（SSMS 开启 SQLCMD 模式，或 sqlcmd -v PluginPassword="..."）：
--   :setvar PluginPassword "在这里填 reco_plugin 的密码"
-- 密码不要写进本文件、不要提交到 Git。脚本可重复执行（幂等）。
-- 权限：RecoLearning 读写；RecoData2024 只读；现有全部用户库只读；model 预置只读（新建项目库自动继承）。
SET NOCOUNT ON;

IF N'$(PluginPassword)' = N'' OR N'$(PluginPassword)' = N'$' + N'(PluginPassword)'
BEGIN
    RAISERROR(N'请先用 :setvar PluginPassword 提供 reco_plugin 的密码。', 16, 1);
    RETURN;
END;

IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = N'reco_plugin')
BEGIN
    CREATE LOGIN [reco_plugin] WITH PASSWORD = N'$(PluginPassword)', CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF, DEFAULT_DATABASE = [RecoLearning];
    PRINT N'已创建登录名 reco_plugin。';
END
ELSE
BEGIN
    ALTER LOGIN [reco_plugin] WITH PASSWORD = N'$(PluginPassword)';
    PRINT N'登录名 reco_plugin 已存在，已同步密码。';
END;

-- 明确只保留 CONNECT SQL，不加入任何服务器角色。
IF EXISTS (SELECT 1 FROM sys.server_role_members m JOIN sys.server_principals r ON r.principal_id = m.role_principal_id
           JOIN sys.server_principals p ON p.principal_id = m.member_principal_id WHERE p.name = N'reco_plugin' AND r.name <> N'public')
    RAISERROR(N'reco_plugin 不应属于任何服务器角色，请人工核对后移除。', 16, 1);

DECLARE @db sysname, @roles nvarchar(200), @sql nvarchar(max);

-- 固定库：RecoLearning 读写，RecoData2024 只读，model 只读。
DECLARE fixed_cursor CURSOR LOCAL FAST_FORWARD FOR
    SELECT N'RecoLearning', N'db_datareader,db_datawriter'
    UNION ALL SELECT N'RecoData2024', N'db_datareader'
    UNION ALL SELECT N'RecoData2020', N'db_datareader'   -- .213 上也有一份 2020 定额库，正式目录的 2020 宿主实际连 .213
    UNION ALL SELECT N'model', N'db_datareader';
OPEN fixed_cursor;
FETCH NEXT FROM fixed_cursor INTO @db, @roles;
WHILE @@FETCH_STATUS = 0
BEGIN
    IF DB_ID(@db) IS NULL
        PRINT N'跳过：库不存在 ' + @db;
    ELSE
    BEGIN
        SET @sql = N'USE ' + QUOTENAME(@db) + N';
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N''reco_plugin'')
    CREATE USER [reco_plugin] FOR LOGIN [reco_plugin];
EXEC sp_addrolemember N''db_datareader'', N''reco_plugin'';'
            + CASE WHEN @roles LIKE N'%db_datawriter%' THEN N'
EXEC sp_addrolemember N''db_datawriter'', N''reco_plugin'';' ELSE N'' END;
        EXEC sp_executesql @sql;
        PRINT N'已授权 ' + @db + N'：' + @roles;
    END;
    FETCH NEXT FROM fixed_cursor INTO @db, @roles;
END;
CLOSE fixed_cursor; DEALLOCATE fixed_cursor;

-- 其余在线用户库（项目库）：只读。
DECLARE db_cursor CURSOR LOCAL FAST_FORWARD FOR
    SELECT name FROM sys.databases
    WHERE database_id > 4 AND state_desc = N'ONLINE' AND is_read_only = 0
      AND name NOT IN (N'RecoLearning', N'RecoData2024', N'RecoData2020')
    ORDER BY name;
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
