set nocount on;
set xact_abort on;

begin transaction;

update d
set 条目序号 = 0
from 定额库 d
join 定额库索引 i on i.书号 = d.书号
where i.分类 in (N'概算定额', N'估算定额')
  and d.条目序号 is null;

declare @updated_count int = @@rowcount;

if @updated_count <> 7385
begin
    rollback transaction;
    raiserror(N'回滚数量与预期不符，已取消回滚。', 16, 1);
    return;
end;

commit transaction;

select @updated_count as UpdatedRows;
