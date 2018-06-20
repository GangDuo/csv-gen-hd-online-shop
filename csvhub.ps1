$wsobj = new-object -comobject wscript.shell

if($args.Length -eq 0) {
    $wsobj.popup("ファイルを指定してください。")
    exit 1
}

$script:total = (Get-Content $Args[0] | ConvertFrom-CSV | Measure-Object '発注予定数' -Sum).Sum
$script:key = "W" + [string][int64](([datetime]::UtcNow)-(get-date "1/1/1970")).TotalSeconds
$script:i = 0
$script:out = "$($env:USERPROFILE)\Desktop\$($script:key).csv"

Get-Content $Args[0] | `
ConvertFrom-CSV | `
Select-Object @{Name='伝票番号';Expression={$script:key}}, `
			  @{Name='伝票行番号';Expression={(++$script:i)}}, `
			  @{Name='f3';Expression={''}}, `
			  @{Name='f4';Expression={''}}, `
			  @{Name='f5';Expression={''}}, `
			  @{Name='出荷元';Expression={'9998'}}, `
			  @{Name='店舗';Expression={'401'}}, `
			  '仕入先コード', `
			  @{Name='アイテム';Expression={'0'}}, `
			  @{Name='品番';Expression={'0'}}, `
			  @{Name='品名';Expression={'0'}}, `
			  @{Name='色コード';Expression={'0'}}, `
			  @{Name='サイズコード';Expression={'0'}}, `
			  'JANコード', `
			  '発注予定数', `
			  @{Name='上代単価';Expression={'0'}}, `
			  @{Name='下代単価';Expression={'0'}}, `
			  @{Name='作成日付';Expression={'0'}}, `
			  @{Name='f16';Expression={''}}, `
			  @{Name='f17';Expression={''}}, `
			  @{Name='更新フラグ';Expression={'1'}}, `
			  @{Name='総数量';Expression={$script:total}}, `
			  @{Name='仕入単価';Expression={'0'}}, `
			  @{Name='商品区分';Expression={'0'}}, `
			  @{Name='客注区分';Expression={'0'}} | `
ConvertTo-Csv -NoTypeInformation | `
Select -Skip 1 | `
% {$_.Replace('"','')} | `
Out-File -Encoding default $script:out


$result = $wsobj.popup($script:key + "作成しました。")