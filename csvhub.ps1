$wsobj = new-object -comobject wscript.shell

if($args.Length -eq 0) {
    $wsobj.popup("�t�@�C�����w�肵�Ă��������B")
    exit 1
}

$script:total = (Get-Content $Args[0] | ConvertFrom-CSV | Measure-Object '�����\�萔' -Sum).Sum
$script:key = [string][int64](([datetime]::UtcNow)-(get-date "1/1/1970")).TotalSeconds
$script:i = 0
$script:out = "$($env:USERPROFILE)\Desktop\$($script:key).csv"

Get-Content $Args[0] | `
ConvertFrom-CSV | `
Select-Object @{Name='f1';Expression={$script:key}}, `
@{Name='f2';Expression={(++$script:i)}}, `
@{Name='f3';Expression={''}}, `
@{Name='f4';Expression={''}}, `
@{Name='f5';Expression={''}}, `
@{Name='�o�׌�';Expression={'9998'}}, `
@{Name='�X��';Expression={'401'}}, `
@{Name='f7';Expression={''}}, `
@{Name='f8';Expression={''}}, `
@{Name='f9';Expression={''}}, `
@{Name='f10';Expression={''}}, `
@{Name='f11';Expression={''}}, `
@{Name='f12';Expression={''}}, `
'JAN�R�[�h', `
'�����\�萔', `
@{Name='f13';Expression={''}}, `
@{Name='f14';Expression={''}}, `
@{Name='f15';Expression={''}}, `
@{Name='f16';Expression={''}}, `
@{Name='f17';Expression={''}}, `
@{Name='�X�V�t���O';Expression={'1'}}, `
@{Name='������';Expression={$script:total}}, `
@{Name='f18';Expression={'0'}}, `
@{Name='f19';Expression={'0'}} | `
ConvertTo-Csv -NoTypeInformation | `
Select -Skip 1 | `
% {$_.Replace('"','')} | `
Out-File -Encoding default $script:out


$result = $wsobj.popup($script:key + "�쐬���܂����B")