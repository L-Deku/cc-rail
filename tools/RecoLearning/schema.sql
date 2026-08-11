-- tools/RecoLearning/schema.sql  (幂等:可重复执行)
IF OBJECT_ID('dbo.BindingLog') IS NULL
CREATE TABLE dbo.BindingLog (
  id            BIGINT IDENTITY(1,1) PRIMARY KEY,
  occurred_at   DATETIME2(0) NOT NULL,
  imported_at   DATETIME2(0) NOT NULL DEFAULT SYSDATETIME(),
  source        NVARCHAR(50)  NOT NULL,
  method        NVARCHAR(50)  NOT NULL DEFAULT(''),
  software_partition NVARCHAR(10) NOT NULL,
  method_no     NVARCHAR(100) NOT NULL,
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
IF COL_LENGTH('dbo.BindingLog','software_partition') IS NULL
  ALTER TABLE dbo.BindingLog ADD software_partition NVARCHAR(10) NULL;
IF COL_LENGTH('dbo.BindingLog','method_no') IS NULL
  ALTER TABLE dbo.BindingLog ADD method_no NVARCHAR(100) NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='UX_BindingLog_event_hash')
  CREATE UNIQUE INDEX UX_BindingLog_event_hash ON dbo.BindingLog(event_hash);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_BindingLog_target_code')
  CREATE INDEX IX_BindingLog_target_code ON dbo.BindingLog(target_code);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_BindingLog_group_key')
  CREATE INDEX IX_BindingLog_group_key ON dbo.BindingLog(group_key);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_BindingLog_recommend_entry')
  CREATE INDEX IX_BindingLog_recommend_entry
    ON dbo.BindingLog(method, target_kind, entry_code, target_code)
    INCLUDE(entry_name);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_BindingLog_recommend_source')
  CREATE INDEX IX_BindingLog_recommend_source
    ON dbo.BindingLog(method, source, target_kind, id)
    INCLUDE(project_id, target_code);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_BindingLog_partition_entry')
  CREATE INDEX IX_BindingLog_partition_entry
    ON dbo.BindingLog(software_partition, method_no, target_kind, entry_code, target_code)
    INCLUDE(entry_name);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_BindingLog_partition_source')
  CREATE INDEX IX_BindingLog_partition_source
    ON dbo.BindingLog(software_partition, source, target_kind, id)
    INCLUDE(project_id, target_code, method_no);

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
  software_partition NVARCHAR(10) NOT NULL,
  signature       NVARCHAR(450) NOT NULL,
  method          NVARCHAR(50) NOT NULL DEFAULT(''),
  box_id          NVARCHAR(64) NOT NULL,
  weight          INT NOT NULL DEFAULT(0),
  accepted_count  INT NOT NULL DEFAULT(0),
  corrected_count INT NOT NULL DEFAULT(0),
  rejected_count  INT NOT NULL DEFAULT(0),
  last_used_at    DATETIME2(0) NULL,
  CONSTRAINT PK_SignatureBoxMap PRIMARY KEY (software_partition, signature, box_id)
);
IF COL_LENGTH('dbo.SignatureBoxMap','software_partition') IS NULL
  ALTER TABLE dbo.SignatureBoxMap ADD software_partition NVARCHAR(10) NULL;
IF COL_LENGTH('dbo.SignatureBoxMap','method') IS NULL
  ALTER TABLE dbo.SignatureBoxMap ADD method NVARCHAR(50) NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_SignatureBoxMap_partition')
  CREATE INDEX IX_SignatureBoxMap_partition
    ON dbo.SignatureBoxMap(software_partition, signature, box_id)
    INCLUDE(method, weight, accepted_count, corrected_count, rejected_count, last_used_at);

IF OBJECT_ID('dbo.QuantityFormulaRule') IS NULL
CREATE TABLE dbo.QuantityFormulaRule (
  rule_hash        CHAR(32) PRIMARY KEY,
  anchor_signature NVARCHAR(450) NOT NULL,
  target_kind      NVARCHAR(20) NOT NULL,
  target_code      NVARCHAR(100) NOT NULL,
  target_unit      NVARCHAR(50) NOT NULL,
  formula_template NVARCHAR(2000) NOT NULL,
  method           NVARCHAR(50) NOT NULL DEFAULT(''),
  software_partition NVARCHAR(10) NOT NULL,
  method_no        NVARCHAR(100) NOT NULL,
  entry_code       NVARCHAR(100) NOT NULL DEFAULT(''),
  sample_count     INT NOT NULL DEFAULT(0),
  first_seen       DATETIME2(0) NULL,
  last_seen        DATETIME2(0) NULL
);
IF COL_LENGTH('dbo.QuantityFormulaRule','software_partition') IS NULL
  ALTER TABLE dbo.QuantityFormulaRule ADD software_partition NVARCHAR(10) NULL;
IF COL_LENGTH('dbo.QuantityFormulaRule','method_no') IS NULL
  ALTER TABLE dbo.QuantityFormulaRule ADD method_no NVARCHAR(100) NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_QuantityFormulaRule_lookup')
  CREATE INDEX IX_QuantityFormulaRule_lookup
    ON dbo.QuantityFormulaRule(anchor_signature, target_code, target_unit, method);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_QuantityFormulaRule_method')
  CREATE INDEX IX_QuantityFormulaRule_method
    ON dbo.QuantityFormulaRule(method, target_code)
    INCLUDE(anchor_signature, target_kind, target_unit, formula_template, entry_code, sample_count, last_seen);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_QuantityFormulaRule_partition')
  CREATE INDEX IX_QuantityFormulaRule_partition
    ON dbo.QuantityFormulaRule(software_partition, method_no, anchor_signature, target_code, target_unit)
    INCLUDE(target_kind, formula_template, entry_code, sample_count, last_seen);

IF OBJECT_ID('dbo.QuantityFormulaOperand') IS NULL
CREATE TABLE dbo.QuantityFormulaOperand (
  rule_hash         CHAR(32) NOT NULL,
  operand_index     INT NOT NULL,
  operand_signature NVARCHAR(450) NOT NULL,
  operand_name      NVARCHAR(1000) NOT NULL DEFAULT(''),
  operand_unit      NVARCHAR(50) NOT NULL DEFAULT(''),
  CONSTRAINT PK_QuantityFormulaOperand PRIMARY KEY (rule_hash, operand_index)
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_QuantityFormulaOperand_signature')
  CREATE INDEX IX_QuantityFormulaOperand_signature ON dbo.QuantityFormulaOperand(operand_signature);

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

IF OBJECT_ID('dbo.SheetTemplateRow') IS NULL
CREATE TABLE dbo.SheetTemplateRow (
  id               BIGINT IDENTITY(1,1) PRIMARY KEY,
  method           NVARCHAR(50) NOT NULL DEFAULT(''),
  workbook         NVARCHAR(260) NOT NULL,
  worksheet        NVARCHAR(100) NOT NULL DEFAULT(''),
  excel_row        INT NOT NULL DEFAULT(0),
  signature        NVARCHAR(450) NOT NULL,
  box_id           NVARCHAR(64) NOT NULL DEFAULT(''),
  entry_code       NVARCHAR(100) NOT NULL DEFAULT(''),
  engineering_type NVARCHAR(50) NOT NULL DEFAULT(''),
  sample_count     INT NOT NULL DEFAULT(0),
  last_seen        DATETIME2(0) NULL
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_SheetTemplateRow_sheet')
  CREATE INDEX IX_SheetTemplateRow_sheet ON dbo.SheetTemplateRow(worksheet, workbook);

IF OBJECT_ID('dbo.SignatureEntryMap') IS NULL
CREATE TABLE dbo.SignatureEntryMap (
  software_partition NVARCHAR(10) NOT NULL,
  method_no     NVARCHAR(100) NOT NULL,
  signature    NVARCHAR(450) NOT NULL,
  target_code  NVARCHAR(100) NOT NULL,
  method       NVARCHAR(50) NOT NULL DEFAULT(''),
  entry_code   NVARCHAR(100) NOT NULL,
  entry_name   NVARCHAR(500) NOT NULL DEFAULT(''),
  sample_count INT NOT NULL DEFAULT(0),
  last_used_at DATETIME2(0) NULL,
  CONSTRAINT PK_SignatureEntryMap PRIMARY KEY (software_partition, method_no, signature, target_code, entry_code)
);
IF COL_LENGTH('dbo.SignatureEntryMap','software_partition') IS NULL
  ALTER TABLE dbo.SignatureEntryMap ADD software_partition NVARCHAR(10) NULL;
IF COL_LENGTH('dbo.SignatureEntryMap','method_no') IS NULL
  ALTER TABLE dbo.SignatureEntryMap ADD method_no NVARCHAR(100) NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_SignatureEntryMap_method')
  CREATE INDEX IX_SignatureEntryMap_method
    ON dbo.SignatureEntryMap(method, target_code)
    INCLUDE(signature, entry_code, entry_name, sample_count, last_used_at);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_SignatureEntryMap_partition')
  CREATE INDEX IX_SignatureEntryMap_partition
    ON dbo.SignatureEntryMap(software_partition, method_no, target_code)
    INCLUDE(signature, entry_code, entry_name, sample_count, last_used_at);

IF OBJECT_ID('dbo.EngineeringTemplate') IS NULL
CREATE TABLE dbo.EngineeringTemplate (
  software_partition NVARCHAR(10) NOT NULL,
  method_no        NVARCHAR(100) NOT NULL,
  method           NVARCHAR(50) NOT NULL,
  engineering_type NVARCHAR(50) NOT NULL,
  entry_code       NVARCHAR(100) NOT NULL,
  box_id           NVARCHAR(64) NOT NULL,
  sample_count     INT NOT NULL DEFAULT(0),
  last_seen        DATETIME2(0) NULL,
  CONSTRAINT PK_EngineeringTemplate PRIMARY KEY (software_partition, method_no, engineering_type, entry_code, box_id)
);
IF COL_LENGTH('dbo.EngineeringTemplate','software_partition') IS NULL
  ALTER TABLE dbo.EngineeringTemplate ADD software_partition NVARCHAR(10) NULL;
IF COL_LENGTH('dbo.EngineeringTemplate','method_no') IS NULL
  ALTER TABLE dbo.EngineeringTemplate ADD method_no NVARCHAR(100) NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_EngineeringTemplate_partition')
  CREATE INDEX IX_EngineeringTemplate_partition
    ON dbo.EngineeringTemplate(software_partition, method_no, engineering_type, entry_code, box_id)
    INCLUDE(method, sample_count, last_seen);
