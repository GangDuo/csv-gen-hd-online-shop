$wsobj = new-object -comobject wscript.shell

if($args.Length -eq 0) {
    $wsobj.popup("�t�@�C�����w�肵�Ă��������B")
    exit 1
}

$script:total = (Get-Content $Args[0] | ConvertFrom-CSV | Measure-Object '�����\�萔' -Sum).Sum
$script:key = "W" + [string][int64](([datetime]::UtcNow)-(get-date "1/1/1970")).TotalSeconds
$script:i = 0
$script:out = "$($env:USERPROFILE)\Desktop\$($script:key).csv"

Get-Content $Args[0] | `
ConvertFrom-CSV | `
Select-Object @{Name='�`�[�ԍ�';Expression={$script:key}}, `
			  @{Name='�`�[�s�ԍ�';Expression={(++$script:i)}}, `
			  @{Name='f3';Expression={''}}, `
			  @{Name='f4';Expression={''}}, `
			  @{Name='f5';Expression={''}}, `
			  @{Name='�o�׌�';Expression={'9998'}}, `
			  @{Name='�X��';Expression={'401'}}, `
			  '�d����R�[�h', `
			  @{Name='�A�C�e��';Expression={'0'}}, `
			  @{Name='�i��';Expression={'0'}}, `
			  @{Name='�i��';Expression={'0'}}, `
			  @{Name='�F�R�[�h';Expression={'0'}}, `
			  @{Name='�T�C�Y�R�[�h';Expression={'0'}}, `
			  'JAN�R�[�h', `
			  '�����\�萔', `
			  @{Name='���P��';Expression={'0'}}, `
			  @{Name='����P��';Expression={'0'}}, `
			  @{Name='�쐬���t';Expression={'0'}}, `
			  @{Name='f16';Expression={''}}, `
			  @{Name='f17';Expression={''}}, `
			  @{Name='�X�V�t���O';Expression={'1'}}, `
			  @{Name='������';Expression={$script:total}}, `
			  @{Name='�d���P��';Expression={'0'}}, `
			  @{Name='���i�敪';Expression={'0'}}, `
			  @{Name='�q���敪';Expression={'0'}} | `
ConvertTo-Csv -NoTypeInformation | `
Select -Skip 1 | `
% {$_.Replace('"','')} | `
Out-File -Encoding default $script:out


$result = $wsobj.popup($script:key + "�쐬���܂����B")