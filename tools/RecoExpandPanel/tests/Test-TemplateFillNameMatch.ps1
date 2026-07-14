$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$dll = "D:\AI文件\自动预算\RecoQuotaRecommend\bin\RecoExpandPanel.dll"
if (-not (Test-Path $dll)) { throw "找不到 $dll，先构建" }
$type = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel')
$flags = [System.Reflection.BindingFlags]'NonPublic,Static'
$norm = $type.GetMethod('NormalizeMatchText', $flags)
$score = $type.GetMethod('MatchNameScore', $flags)

function N($s) { return $norm.Invoke($null, @($s)) }
function S($a, $b) { return $score.Invoke($null, @((N $a), (N $b))) }

if ((N "铺设 无缝线路（km）") -ne "铺设无缝线路km") { throw "归一化失败: $(N '铺设 无缝线路（km）')" }
Write-Host "PASS 归一化"
if ((S "铺设无缝线路" "铺设无缝线路") -ne 100) { throw "同名应100" }
Write-Host "PASS 同名满分"
if ((S "铺设无缝线路" "无缝线路铺设") -lt 55) { throw "改写措辞应>=55, 实际 $(S '铺设无缝线路' '无缝线路铺设')" }
Write-Host "PASS 改写措辞命中"
if ((S "500m长轨铺设" "25m长轨铺设") -ge 55) { throw "数字不符应<55, 实际 $(S '500m长轨铺设' '25m长轨铺设')" }
Write-Host "PASS 数字不符不误配"
if ((S "土方开挖" "钢筋制作安装") -ge 40) { throw "无关应低分" }
Write-Host "PASS 无关低分"
Write-Host "全部通过"
