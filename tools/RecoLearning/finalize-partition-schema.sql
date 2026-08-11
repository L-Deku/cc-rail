-- DDL-2. This file is executed only by Migrate-LearningPartitionSchema.ps1 -Finalize
-- inside the same transaction that truncates every derived aggregate table.
SET NOCOUNT ON;
SET XACT_ABORT ON;
IF @@TRANCOUNT = 0
BEGIN
  RAISERROR('DDL-2 requires an existing migration transaction.', 16, 1);
  RETURN;
END;

TRUNCATE TABLE dbo.QuantityFormulaOperand;
TRUNCATE TABLE dbo.QuantityFormulaRule;
TRUNCATE TABLE dbo.SignatureEntryMap;
TRUNCATE TABLE dbo.EngineeringTemplate;
TRUNCATE TABLE dbo.SignatureBoxMap;
TRUNCATE TABLE dbo.QuotaBoxTarget;
TRUNCATE TABLE dbo.QuotaBox;
TRUNCATE TABLE dbo.QuantityAlias;
TRUNCATE TABLE dbo.SheetTemplateRow;

IF EXISTS (SELECT 1 FROM dbo.BindingLog WHERE software_partition IS NULL OR method_no IS NULL)
BEGIN
  RAISERROR('BindingLog partition columns still contain NULL.', 16, 1);
  RETURN;
END;

-- SQL Server blocks ALTER COLUMN while any nonclustered index references the column.
-- DDL-2 owns the surrounding transaction, so these indexes are restored below before commit.
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.BindingLog') AND name='IX_BindingLog_partition_entry')
  DROP INDEX IX_BindingLog_partition_entry ON dbo.BindingLog;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.BindingLog') AND name='IX_BindingLog_partition_source')
  DROP INDEX IX_BindingLog_partition_source ON dbo.BindingLog;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.SignatureBoxMap') AND name='IX_SignatureBoxMap_partition')
  DROP INDEX IX_SignatureBoxMap_partition ON dbo.SignatureBoxMap;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.QuantityFormulaRule') AND name='IX_QuantityFormulaRule_partition')
  DROP INDEX IX_QuantityFormulaRule_partition ON dbo.QuantityFormulaRule;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.SignatureEntryMap') AND name='IX_SignatureEntryMap_partition')
  DROP INDEX IX_SignatureEntryMap_partition ON dbo.SignatureEntryMap;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id=OBJECT_ID('dbo.EngineeringTemplate') AND name='IX_EngineeringTemplate_partition')
  DROP INDEX IX_EngineeringTemplate_partition ON dbo.EngineeringTemplate;

ALTER TABLE dbo.BindingLog ALTER COLUMN software_partition NVARCHAR(10) NOT NULL;
ALTER TABLE dbo.BindingLog ALTER COLUMN method_no NVARCHAR(100) NOT NULL;

DECLARE @constraint_name SYSNAME;
DECLARE @sql NVARCHAR(MAX);

SELECT @constraint_name = kc.name
FROM sys.key_constraints kc
WHERE kc.parent_object_id = OBJECT_ID('dbo.SignatureBoxMap') AND kc.[type] = 'PK';
IF @constraint_name IS NOT NULL
BEGIN
  SET @sql = N'ALTER TABLE dbo.SignatureBoxMap DROP CONSTRAINT ' + QUOTENAME(@constraint_name) + N';';
  EXEC sys.sp_executesql @sql;
END;
ALTER TABLE dbo.SignatureBoxMap ALTER COLUMN software_partition NVARCHAR(10) NOT NULL;
ALTER TABLE dbo.SignatureBoxMap ADD CONSTRAINT PK_SignatureBoxMap
  PRIMARY KEY (software_partition, signature, box_id);

SELECT @constraint_name = NULL;
SELECT @constraint_name = kc.name
FROM sys.key_constraints kc
WHERE kc.parent_object_id = OBJECT_ID('dbo.SignatureEntryMap') AND kc.[type] = 'PK';
IF @constraint_name IS NOT NULL
BEGIN
  SET @sql = N'ALTER TABLE dbo.SignatureEntryMap DROP CONSTRAINT ' + QUOTENAME(@constraint_name) + N';';
  EXEC sys.sp_executesql @sql;
END;
ALTER TABLE dbo.SignatureEntryMap ALTER COLUMN software_partition NVARCHAR(10) NOT NULL;
ALTER TABLE dbo.SignatureEntryMap ALTER COLUMN method_no NVARCHAR(100) NOT NULL;
ALTER TABLE dbo.SignatureEntryMap ADD CONSTRAINT PK_SignatureEntryMap
  PRIMARY KEY (software_partition, method_no, signature, target_code, entry_code);

SELECT @constraint_name = NULL;
SELECT @constraint_name = kc.name
FROM sys.key_constraints kc
WHERE kc.parent_object_id = OBJECT_ID('dbo.EngineeringTemplate') AND kc.[type] = 'PK';
IF @constraint_name IS NOT NULL
BEGIN
  SET @sql = N'ALTER TABLE dbo.EngineeringTemplate DROP CONSTRAINT ' + QUOTENAME(@constraint_name) + N';';
  EXEC sys.sp_executesql @sql;
END;
ALTER TABLE dbo.EngineeringTemplate ALTER COLUMN software_partition NVARCHAR(10) NOT NULL;
ALTER TABLE dbo.EngineeringTemplate ALTER COLUMN method_no NVARCHAR(100) NOT NULL;
ALTER TABLE dbo.EngineeringTemplate ADD CONSTRAINT PK_EngineeringTemplate
  PRIMARY KEY (software_partition, method_no, engineering_type, entry_code, box_id);

ALTER TABLE dbo.QuantityFormulaRule ALTER COLUMN software_partition NVARCHAR(10) NOT NULL;
ALTER TABLE dbo.QuantityFormulaRule ALTER COLUMN method_no NVARCHAR(100) NOT NULL;

CREATE INDEX IX_BindingLog_partition_entry
  ON dbo.BindingLog(software_partition, method_no, target_kind, entry_code, target_code)
  INCLUDE(entry_name);
CREATE INDEX IX_BindingLog_partition_source
  ON dbo.BindingLog(software_partition, source, target_kind, id)
  INCLUDE(project_id, target_code, method_no);
CREATE INDEX IX_SignatureBoxMap_partition
  ON dbo.SignatureBoxMap(software_partition, signature, box_id)
  INCLUDE(method, weight, accepted_count, corrected_count, rejected_count, last_used_at);
CREATE INDEX IX_QuantityFormulaRule_partition
  ON dbo.QuantityFormulaRule(software_partition, method_no, anchor_signature, target_code, target_unit)
  INCLUDE(target_kind, formula_template, entry_code, sample_count, last_seen);
CREATE INDEX IX_SignatureEntryMap_partition
  ON dbo.SignatureEntryMap(software_partition, method_no, target_code)
  INCLUDE(signature, entry_code, entry_name, sample_count, last_used_at);
CREATE INDEX IX_EngineeringTemplate_partition
  ON dbo.EngineeringTemplate(software_partition, method_no, engineering_type, entry_code, box_id)
  INCLUDE(method, sample_count, last_seen);
