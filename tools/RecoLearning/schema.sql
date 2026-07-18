-- tools/RecoLearning/schema.sql  (幂等:可重复执行)
IF OBJECT_ID('dbo.BindingLog') IS NULL
CREATE TABLE dbo.BindingLog (
  id            BIGINT IDENTITY(1,1) PRIMARY KEY,
  occurred_at   DATETIME2(0) NOT NULL,
  imported_at   DATETIME2(0) NOT NULL DEFAULT SYSDATETIME(),
  source        NVARCHAR(50)  NOT NULL,
  method        NVARCHAR(50)  NOT NULL DEFAULT(''),
  project_id    NVARCHAR(200) NOT NULL DEFAULT(''),
  entry_code    NVARCHAR(100) NOT NULL DEFAULT(''),
  entry_name    NVARCHAR(500) NOT NULL DEFAULT(''),
  quantity_name NVARCHAR(1000) NOT NULL,
  quantity_unit NVARCHAR(50)  NOT NULL DEFAULT(''),
  target_kind   NVARCHAR(20)  NOT NULL,
  target_code   NVARCHAR(100) NOT NULL,
  target_name   NVARCHAR(500) NOT NULL DEFAULT(''),
  target_unit   NVARCHAR(50)  NOT NULL DEFAULT(''),
  group_key     CHAR(32)      NOT NULL,
  event_hash    CHAR(32)      NOT NULL,
  extra         NVARCHAR(MAX) NULL
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='UX_BindingLog_event_hash')
  CREATE UNIQUE INDEX UX_BindingLog_event_hash ON dbo.BindingLog(event_hash);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_BindingLog_target_code')
  CREATE INDEX IX_BindingLog_target_code ON dbo.BindingLog(target_code);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_BindingLog_group_key')
  CREATE INDEX IX_BindingLog_group_key ON dbo.BindingLog(group_key);

IF OBJECT_ID('dbo.QuantityAlias') IS NULL
CREATE TABLE dbo.QuantityAlias (
  alias_hash    CHAR(32) PRIMARY KEY,
  raw_name      NVARCHAR(1000) NOT NULL,
  quantity_unit NVARCHAR(50) NOT NULL DEFAULT(''),
  signature     NVARCHAR(450) NOT NULL,
  seen_count    INT NOT NULL DEFAULT(0),
  first_seen    DATETIME2(0) NULL,
  last_seen     DATETIME2(0) NULL
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_QuantityAlias_signature')
  CREATE INDEX IX_QuantityAlias_signature ON dbo.QuantityAlias(signature);

IF OBJECT_ID('dbo.QuotaBox') IS NULL
CREATE TABLE dbo.QuotaBox (
  box_id          NVARCHAR(64) PRIMARY KEY,
  target_set_hash CHAR(32) NOT NULL,
  status          NVARCHAR(20) NOT NULL DEFAULT('active'),
  created_at      DATETIME2(0) NOT NULL DEFAULT SYSDATETIME()
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='UX_QuotaBox_target_set')
  CREATE UNIQUE INDEX UX_QuotaBox_target_set ON dbo.QuotaBox(target_set_hash);

IF OBJECT_ID('dbo.QuotaBoxTarget') IS NULL
CREATE TABLE dbo.QuotaBoxTarget (
  box_id      NVARCHAR(64) NOT NULL,
  target_kind NVARCHAR(20) NOT NULL,
  target_code NVARCHAR(100) NOT NULL,
  target_name NVARCHAR(500) NOT NULL DEFAULT(''),
  target_unit NVARCHAR(50) NOT NULL DEFAULT(''),
  CONSTRAINT PK_QuotaBoxTarget PRIMARY KEY (box_id, target_kind, target_code)
);

IF OBJECT_ID('dbo.SignatureBoxMap') IS NULL
CREATE TABLE dbo.SignatureBoxMap (
  signature       NVARCHAR(450) NOT NULL,
  box_id          NVARCHAR(64) NOT NULL,
  weight          INT NOT NULL DEFAULT(0),
  accepted_count  INT NOT NULL DEFAULT(0),
  corrected_count INT NOT NULL DEFAULT(0),
  rejected_count  INT NOT NULL DEFAULT(0),
  last_used_at    DATETIME2(0) NULL,
  CONSTRAINT PK_SignatureBoxMap PRIMARY KEY (signature, box_id)
);

IF OBJECT_ID('dbo.EntryQuota') IS NULL
CREATE TABLE dbo.EntryQuota (
  method        NVARCHAR(50) NOT NULL,
  method_no     NVARCHAR(100) NOT NULL DEFAULT(''),
  entry_code    NVARCHAR(100) NOT NULL,
  entry_name    NVARCHAR(500) NOT NULL DEFAULT(''),
  target_kind   NVARCHAR(20) NOT NULL DEFAULT('quota'),
  quota_code    NVARCHAR(100) NOT NULL,
  quota_name    NVARCHAR(500) NOT NULL DEFAULT(''),
  quota_unit    NVARCHAR(50) NOT NULL DEFAULT(''),
  project_count INT NOT NULL DEFAULT(0),
  source        NVARCHAR(50) NOT NULL DEFAULT(''),
  last_seen     NVARCHAR(50) NOT NULL DEFAULT(''),
  CONSTRAINT PK_EntryQuota PRIMARY KEY (method, entry_code, target_kind, quota_code)
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_EntryQuota_quota')
  CREATE INDEX IX_EntryQuota_quota ON dbo.EntryQuota(method, quota_code);

IF OBJECT_ID('dbo.ChapterEntry') IS NULL
CREATE TABLE dbo.ChapterEntry (
  method     NVARCHAR(50) NOT NULL,
  method_no  NVARCHAR(100) NOT NULL DEFAULT(''),
  entry_code NVARCHAR(100) NOT NULL,
  entry_name NVARCHAR(500) NOT NULL DEFAULT(''),
  unit       NVARCHAR(50) NOT NULL DEFAULT(''),
  entry_type NVARCHAR(50) NOT NULL DEFAULT(''),
  level      INT NOT NULL DEFAULT(0),
  CONSTRAINT PK_ChapterEntry PRIMARY KEY (method, entry_code)
);

IF OBJECT_ID('dbo.EngineeringTemplate') IS NULL
CREATE TABLE dbo.EngineeringTemplate (
  method           NVARCHAR(50) NOT NULL,
  engineering_type NVARCHAR(50) NOT NULL,
  entry_code       NVARCHAR(100) NOT NULL,
  box_id           NVARCHAR(64) NOT NULL,
  sample_count     INT NOT NULL DEFAULT(0),
  last_seen        DATETIME2(0) NULL,
  CONSTRAINT PK_EngineeringTemplate PRIMARY KEY (method, engineering_type, entry_code, box_id)
);
