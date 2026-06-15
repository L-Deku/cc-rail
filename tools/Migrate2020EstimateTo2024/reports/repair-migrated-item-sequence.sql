set nocount on;
set xact_abort on;

begin transaction;

declare @migrated_count int;
declare @zero_count int;
declare @updated_count int;

select @migrated_count = count(*)
from 定额库 d
join 定额库索引 i on i.书号 = d.书号
where i.分类 in (N'概算定额', N'估算定额');

select @zero_count = count(*)
from 定额库 d
join 定额库索引 i on i.书号 = d.书号
where i.分类 in (N'概算定额', N'估算定额')
  and d.条目序号 = 0;

if @migrated_count <> 7385 or @zero_count <> 7385
begin
    rollback transaction;
    raiserror(N'迁移定额数量或条目序号状态与预期不符，已取消修复。', 16, 1);
    return;
end;

update d
set 条目序号 = null
from 定额库 d
join 定额库索引 i on i.书号 = d.书号
where i.分类 in (N'概算定额', N'估算定额')
  and d.条目序号 = 0;

set @updated_count = @@rowcount;

if @updated_count <> 7385
begin
    rollback transaction;
    raiserror(N'实际修复数量与预期不符，已回滚。', 16, 1);
    return;
end;

if exists (
    select 1
    from 定额库 d
    join 定额库索引 i on i.书号 = d.书号
    where i.分类 in (N'概算定额', N'估算定额')
      and d.条目序号 is not null
)
begin
    rollback transaction;
    raiserror(N'修复后仍有迁移定额的条目序号不为空，已回滚。', 16, 1);
    return;
end;

commit transaction;

select @updated_count as UpdatedRows;
