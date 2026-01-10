#Requires -Module ImportExcel
#Requires -Version 7.0

$MOD="ImportExcel"



# This script generates an Excel file that mimics the vns-template.xlsx file.
# It manually builds the Excel package using the EPPlus API to ensure stability and control.

# --- 0. Preamble ---
try {
    # Manually load the EPPlus assembly to ensure types are recognized by the PowerShell session.
    # This is the key to solving the type not found errors.
    Add-Type -Path (Join-Path (Get-Module -ListAvailable ImportExcel | Where-Object { $_.Version.ToString() -eq '7.8.10' } | Select-Object -First 1).ModuleBase "EPPlus.dll")
}
catch {
    Write-Error "EPPlus.dll not found. Please ensure the ImportExcel module (version 7.8.10) is installed correctly."
    return
}

# --- 1. Data Preparation ---
Write-Host "Step 1: Preparing data..."
# Static data loaded from vns-template.csv
$fullData = @(
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-004019'; '[JVN] タイトル'='Oracle MySQL の MySQL Server における Server: Optimizer に関する脆弱性'; '[JVN] CVSS v3 深刻度'='4.9'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:N/AC:L/PR:H/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/4/25 14:49'; '[JVN] CVE'='CVE-2025-21492'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-004019'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='mysql-8.0'; '現行バージョン／更新版バージョン'='8.0.41-0ubuntu0.20.04.1 / 8.0.37-0ubuntu0.20.04.3'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-004020'; '[JVN] タイトル'='Oracle MySQL の MySQL Server における InnoDB に関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:N/AC:L/PR:H/UI:N/S:U/C:N/I:L/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/4/25 15:09'; '[JVN] CVE'='CVE-2025-21497'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-004020'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='mysql-8.0'; '現行バージョン／更新版バージョン'='8.0.41-0ubuntu0.20.04.1 / 8.0.41-0ubuntu0.20.04.1'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-004022'; '[JVN] タイトル'='Oracle MySQL の MySQL Server における Server: Optimizer に関する脆弱性'; '[JVN] CVSS v3 深刻度'='4.9'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:N/AC:L/PR:H/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/4/25 15:13'; '[JVN] CVE'='CVE-2025-21504'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-004022'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='mysql-8.0'; '現行バージョン／更新版バージョン'='8.0.41-0ubuntu0.20.04.1 / 8.0.40-0ubuntu0.20.04.1'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2024-022892'; '[JVN] タイトル'='Xiph.Org の libtheora における不正な認証に関する脆弱性'; '[JVN] CVSS v3 深刻度'='9.8'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:H/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/4/25 17:17'; '[JVN] CVE'='CVE-2024-56431'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2024-022892'; '影響有無'='影響なし(未インストール)'; '対象パッケージ／ソフトウェア'='libtheora'; '現行バージョン／更新版バージョン'=''; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-004045'; '[JVN] タイトル'='Linux の Linux Kernel における NULL ポインタデリファレンスに関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/4/25 20:47'; '[JVN] CVE'='CVE-2025-21917'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-004045'; '影響有無'='修正版適用待ち'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.4.0-208.228 / '; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-004058'; '[JVN] タイトル'='Oracle MySQL の MySQL Server における Server: DML に関する脆弱性'; '[JVN] CVSS v3 深刻度'='4.9'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:N/AC:L/PR:H/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/4/25 23:57'; '[JVN] CVE'='CVE-2025-21580'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-004058'; '影響有無'='修正版配布待ち'; '対象パッケージ／ソフトウェア'='mysql-8.0'; '現行バージョン／更新版バージョン'='8.0.41-0ubuntu0.20.04.1 / '; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2023-029763'; '[JVN] タイトル'='Linux の Linux Kernel における整数オーバーフローの脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/9/26 14:44'; '[JVN] CVE'='CVE-2023-52762'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2023-029763'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-153.163 / 5.15.0-100.110'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2024-028083'; '[JVN] タイトル'='Linux の Linux Kernel における有効期限後のメモリの解放の欠如に関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/9/26 14:51'; '[JVN] CVE'='CVE-2024-35912'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2024-028083'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-153.163 / 5.15.0-116.126'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2024-028084'; '[JVN] タイトル'='Linux の Linux Kernel におけるリソースのロックに関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/9/26 14:53'; '[JVN] CVE'='CVE-2024-35914'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2024-028084'; '影響有無'='影響なし(範囲外)'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-153.163'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2024-028086'; '[JVN] タイトル'='Linux の Linux Kernel における有効期限後のメモリの解放の欠如に関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/9/26 15:08'; '[JVN] CVE'='CVE-2024-35956'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2024-028086'; '影響有無'='修正版配布待ち'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-153.163'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-002140'; '[JVN] タイトル'='OpenBSD の OpenSSH 等複数ベンダの製品におけるリソースの枯渇に関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.9'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:N/AC:H/PR:N/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 11:24'; '[JVN] CVE'='CVE-2025-26466'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-002140'; '影響有無'='影響なし'; '対象パッケージ／ソフトウェア'=''; '現行バージョン／更新版バージョン'=''; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2024-004335'; '[JVN] タイトル'='lighttpd における解放済みメモリ使用 (use-after-free) の脆弱性'; '[JVN] CVSS v3 深刻度'='N/A'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='N/A'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 11:26'; '[JVN] CVE'='CVE-2018-25103'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2024-004335'; '影響有無'='影響なし(範囲外)'; '対象パッケージ／ソフトウェア'='lighttpd'; '現行バージョン／更新版バージョン'=''; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022400'; '[JVN] タイトル'='Ruijie Networks 製 AP180 シリーズにおける OS コマンドインジェクションの脆弱性'; '[JVN] CVSS v3 深刻度'='7.2'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:N/AC:L/PR:H/UI:N/S:U/C:H/I:H/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 11:42'; '[JVN] CVE'='CVE-2025-68459'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022400'; '影響有無'='影響なし'; '対象パッケージ／ソフトウェア'=''; '現行バージョン／更新版バージョン'=''; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022474'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における競合状態に関する脆弱性'; '[JVN] CVSS v3 深刻度'='7'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:H/PR:L/UI:N/S:U/C:H/I:H/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:30'; '[JVN] CVE'='CVE-2025-38108'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022474'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174 / 5.15.0-156.166'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2024-029632'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:31'; '[JVN] CVE'='CVE-2024-36017'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2024-029632'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174 / 5.15.0-118.128'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2023-030306'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:32'; '[JVN] CVE'='CVE-2023-52683'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2023-030306'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174 / 5.15.0-102.112'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022476'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における初期化されていないリソースの使用に関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:34'; '[JVN] CVE'='CVE-2025-37961'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022476'; '影響有無'='修正版配布待ち'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022479'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における NULL ポインタデリファレンスに関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:38'; '[JVN] CVE'='CVE-2025-37972'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022479'; '影響有無'='影響なし(範囲外)'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2024-029635'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:40'; '[JVN] CVE'='CVE-2024-35950'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2024-029635'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174 / 5.15.0-116.126'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022484'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における NULL ポインタデリファレンスに関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:44'; '[JVN] CVE'='CVE-2025-38364'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022484'; '影響有無'='影響なし(範囲外)'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022485'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:46'; '[JVN] CVE'='CVE-2025-38097'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022485'; '影響有無'='修正版配布待ち'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022486'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における初期化されていないリソースの使用に関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:47'; '[JVN] CVE'='CVE-2025-38382'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022486'; '影響有無'='影響なし(範囲外)'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022489'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における例外的な状態のチェックに関する脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:52'; '[JVN] CVE'='CVE-2025-38334'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022489'; '影響有無'='修正版配布待ち'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022490'; '[JVN] タイトル'='アップルの macOS における脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:H/I:N/A:N'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:54'; '[JVN] CVE'='CVE-2025-43416'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022490'; '影響有無'='影響なし'; '対象パッケージ／ソフトウェア'=''; '現行バージョン／更新版バージョン'=''; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022492'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における境界外読み取りに関する脆弱性'; '[JVN] CVSS v3 深刻度'='7.1'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:H/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 17:58'; '[JVN] CVE'='CVE-2025-38342'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022492'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174 / 5.15.0-156.166'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022494'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における境界外書き込みに関する脆弱性'; '[JVN] CVSS v3 深刻度'='7.8'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:H/I:H/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 18:00'; '[JVN] CVE'='CVE-2025-38348'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022494'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174 / 5.15.0-156.166'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022495'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 18:01'; '[JVN] CVE'='CVE-2025-37959'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022495'; '影響有無'='修正版配布待ち'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174'; 'P14'=''},
    [PSCustomObject]@{'[JVN] 区分'='Update'; '[JVN] 番号'='JVNDB-2025-022508'; '[JVN] タイトル'='Linux の Linux Kernel 等複数ベンダの製品における脆弱性'; '[JVN] CVSS v3 深刻度'='5.5'; '[JVN] CVSS v2 深刻度'='N/A'; '[JVN] CVSS v3 URL'='https://jvndb.jvn.jp/cvss/ja/v3.html#CVSS:3.0/AV:L/AC:L/PR:L/UI:N/S:U/C:N/I:N/A:H'; '[JVN] CVSS v2 URL'='N/A'; '[JVN] 更新日'='2025/12/19 18:28'; '[JVN] CVE'='CVE-2025-37964'; '[Gauge] URL'='https://vns.vital-service.com/member/search?id=JVNDB-2025-022508'; '影響有無'='更新版適用済'; '対象パッケージ／ソフトウェア'='Linux Kernel'; '現行バージョン／更新版バージョン'='5.15.0-164.174 / 5.15.0-144.157'; 'P14'=''}
)

# Manually define the correct, multi-line headers as they appear in the template.
$correctHeaders = @(
    "[JVN]`n区分", "[JVN]`n番号", "[JVN]`nタイトル", "[JVN]`nCVSS v3`n深刻度", "[JVN]`nCVSS v2`n深刻度",
    "[JVN]`nCVSS v3 URL", "[JVN]`nCVSS v2 URL", "[JVN]`n更新日", "[JVN]`nCVE", "[Gauge]`nURL",
    "影響有無", "対象パッケージ／ソフトウェア", "現行バージョン／`n更新版バージョン", ""
)

# Define a unique output path to prevent file lock issues.
$fontMSPGothic = 'ＭＳ Ｐゴシック'
$fontYuGothic = '游ゴシック'
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$outputPath = ".\vns_report_$timestamp.xlsx"
$fileInfo = New-Object System.IO.FileInfo $outputPath
if ($fileInfo.Exists) { Remove-Item $outputPath }

# --- 2. Manual Excel Creation ---
Write-Host "Step 2: Creating a new Excel package..."
$pkg = New-Object OfficeOpenXml.ExcelPackage $fileInfo
$ws = $pkg.Workbook.Worksheets.Add('vns_確認')

# --- 3. Populate Data and Headers ---
Write-Host "Step 3: Populating headers and data..."
# Write headers cell by cell.
for ($i = 0; $i -lt $correctHeaders.Count; $i++) { $ws.Cells[1, ($i+1)].Value = $correctHeaders[$i] }

# Write data cell by cell for maximum stability.
$csvHeaders = $fullData[0].psobject.Properties.Name
for ($row = 0; $row -lt $fullData.Count; $row++) {
    for ($col = 0; $col -lt $csvHeaders.Count; $col++) {
        $ws.Cells[($row+2), ($col+1)].Value = $fullData[$row].($csvHeaders[$col])
    }
}

# --- 4. Apply All Formatting ---
Write-Host "Step 4: Applying formatting (columns, rows, fonts, borders)..."
$dataRowCount = $fullData.Count
$totalRange = "A1:N$($dataRowCount + 1)"
$ws.Cells["A1:N1"].AutoFilter = $true
$ws.Column(1).Width = 7.75; $ws.Column(2).Width = 19.33; $ws.Column(3).Width = 99.25; $ws.Column(4).Width = 11; $ws.Column(5).Width = 10.08; $ws.Column(6).Width = 33.16; $ws.Column(7).Width = 22.83; $ws.Column(11).Width = 29; $ws.Column(12).Width = 30; $ws.Column(13).Width = 30;
$ws.Row(1).Height = 39; $ws.Row(1).CustomHeight = $true
for ($i = 2; $i -le ($dataRowCount + 1); $i++) { $ws.Row($i).CustomHeight = $true; $ws.Row($i).Height = 18 }
$headerRange = $ws.Cells["A1:N1"]
$headerRange.Style.Font.Name = $fontMSPGothic
$headerRange.Style.Font.Bold = $true
$headerRange.Style.WrapText = $true
$headerRange.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#FFFFFF'))
$headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$headerRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml('#2D7DCE'))
$dataRange = $ws.Cells["A2:N$($dataRowCount + 1)"]; $dataRange.Style.Font.Name = $fontYuGothic; $dataRange.Style.Font.Size = 11;
$ws.Cells[$totalRange].Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$ws.Cells[$totalRange].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$ws.Cells[$totalRange].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$ws.Cells[$totalRange].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$ws.Cells["A1:N1"].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium

# --- 5. Conditional Formatting ---
Write-Host "Step 5: Applying conditional formatting..."
$cfRange = $ws.Cells["A2:N$($dataRowCount + 1)"] # Apply to the entire row for columns A to N

# Rule 1: Highlight rows where K column is "修正版配布待ち" with color #F1A983
$cfRule1 = $ws.ConditionalFormatting.AddExpression($cfRange)
$cfRule1.Formula = '$K2="修正版配布待ち"'
$cfRule1.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cfRule1.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#F1A983')

# Rule 2: Highlight rows where K column is "修正版適用待ち" with color #F1A983
$cfRule2 = $ws.ConditionalFormatting.AddExpression($cfRange)
$cfRule2.Formula = '$K2="修正版適用待ち"'
$cfRule2.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cfRule2.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#F1A983')

# Rule 3: Highlight rows where K column is "更新版適用済" with color #92D050
$cfRule3 = $ws.ConditionalFormatting.AddExpression($cfRange)
$cfRule3.Formula = '$K2="更新版適用済"'
$cfRule3.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$cfRule3.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#92D050')

# --- 5.5. Add Footer Legend ---
Write-Host "Step 5.5: Adding footer legend..."

# Legend data based on vns-sample.xlsx, range A32:M38
$legendItems = @(
    @{ Row = 0; Description = '脆弱性が修正された更新版パッケージが適用済みであるもの'; Status = '更新版適用済'; Color = '#92D050' },
    @{ Row = 1; Description = '脆弱性が存在するものの、修正パッチがまだ配布されておらず、配布を待っている状態のもの。'; Status = '修正版配布待ち'; Color = '#F1A983' },
    @{ Row = 2; Description = '修正パッチは配布されているが、まだ適用ができていないもの。（再起動を行わないと適用ができないものなど）'; Status = '修正版適用待ち'; Color = '#F1A983' },
    @{ Row = 3; Description = '脆弱性が情報が存在するパッケージであるものの、該当パッケージをインストールしていないため、影響が発生しないもの。'; Status = '影響なし（未インストール）'; Color = $null },
    @{ Row = 4; Description = '脆弱性が情報が存在するパッケージであるものの、該当OSには影響がないことを公式にて確認されているもの。'; Status = '影響なし（範囲外）'; Color = $null },
    @{ Row = 5; Description = '該当OSでは利用されないパッケージ、ソフトウェアであるため、とくに確認対象とはならないもの。'; Status = '影響なし'; Color = $null }
)

$TopRow = $dataRowCount + 4 + ($legendItems | Measure-Object -Property Row -Minimum | Select-Object -ExpandProperty Minimum)
$BottomRow = $dataRowCount + 4 + ($legendItems | Measure-Object -Property Row -Maximum | Select-Object -ExpandProperty Maximum)
$TitleRow = $TopRow - 1
$range="A$($TitleRow)"
$ws.Cells[$range].Value = "凡例"
$ws.Cells[$range].Style.Font.Bold = $true
$ws.Cells[$range].Style.Font.Name = $fontMSPGothic
$ws.Cells[$range].Style.Font.Size = 11
$ws.Row($TitleRow).CustomHeight = $true;
$ws.Row($TitleRow).Height = 13

foreach ($item in $legendItems) {
    $row = $dataRowCount + 4 + $item.Row

    $ws.Row($TopRow).CustomHeight = $true;
    $ws.Row($TopRow).Height = 13

    # Merge cells D:J for the description, set value and style for proper display
    $range="A$($row):M$($row)"
    $rangeLeft="A$($row):C$($row)"
    $rangeCenter="D$($row):J$($row)"
    $rangeRight="K$($row):M$($row)"
    $ws.Cells[$rangeCenter].Merge = $true
    $ws.Cells[$rangeLeft].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)
    $ws.Cells[$rangeCenter].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)
    $ws.Cells[$rangeRight].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)
    $ws.Cells[$rangeRight].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

    $ws.Cells[$rangeCenter].Value = $item.Description
    $ws.Cells[$rangeCenter].Style.WrapText = $true
    $ws.Cells[$rangeCenter].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top

    # Add the status text to column K
    $ws.Cells["K$row"].Value = $item.Status
    #$ws.Cells[$rangeCenter].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin)

    $ws.Cells[$range].Style.Font.Name = $fontMSPGothic
    $ws.Cells[$range].Style.Font.Size = 11
    $ws.Cells[$range].Style.Font.Bold = $false
    # Set the background color for the status cell
    if ($item.Color) {
        $ws.Cells[$range].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        # For regular cell styling, SetColor is often available and reliable.
        $ws.Cells[$range].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($item.Color))
    }

}

$range="A$($TopRow):M$($BottomRow)"
$ws.Cells[$range].Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Medium)

# --- 6. Finalization ---
Write-Host "Step 6: Saving and closing the file..."
$pkg.Save()
$pkg.Dispose()

Write-Host "--------------------------------------------------"
Write-Host "Success! The definitive script has been created as 'generate_excel.ps1'."
Write-Host "The generated Excel file is: $outputPath"
Write-Host "--------------------------------------------------"
