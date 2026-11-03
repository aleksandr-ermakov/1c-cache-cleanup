<#
.SYNOPSIS
    Скрипт выборочной очистки кэша 1С.

.DESCRIPTION
    Скрипт очищает кэш информационных баз 1С под Windows.
    Перед очисткой показывает базы, размер их кэша и меню действий.
    Кэш можно удалить полностью, только программный или только пользовательский —
    у всех найденных баз или у одной выбранной.
    Осиротевшие папки кэша скрипт только показывает.

.PARAMETER IbasesPath
    Путь к файлу ibases.v8i с личным списком баз.
    По умолчанию: %APPDATA%\1C\1CEStart\ibases.v8i.

.PARAMETER StartCfgPath
    Путь к файлу 1cestart.cfg, в котором перечислены общие списки баз.
    По умолчанию: %APPDATA%\1C\1CEStart\1cestart.cfg.

.PARAMETER UserCachePath
    Каталог пользовательского кэша баз (папки с именами-GUID).
    По умолчанию: %APPDATA%\1C\1cv8.

.PARAMETER ProgramCachePath
    Каталог программного кэша баз (папки с именами-GUID).
    По умолчанию: %LOCALAPPDATA%\1C\1cv8.

.PARAMETER ShowOrphans
    Показывает осиротевшие папки кэша подробно: по идентификаторам с размером.

.PARAMETER Help
    Выводит эту справку и завершает работу. То же самое делают ключи -h и --help.

.EXAMPLE
    .\1c-cache-cleanup.ps1

.EXAMPLE
    .\1c-cache-cleanup.ps1 -ShowOrphans

.EXAMPLE
    .\1c-cache-cleanup.ps1 -h
#>

[CmdletBinding()]
param(
    [string] $IbasesPath = (Join-Path -Path $env:APPDATA -ChildPath '1C\1CEStart\ibases.v8i'),
    [string] $StartCfgPath = (Join-Path -Path $env:APPDATA -ChildPath '1C\1CEStart\1cestart.cfg'),
    [string] $UserCachePath = (Join-Path -Path $env:APPDATA -ChildPath '1C\1cv8'),
    [string] $ProgramCachePath = (Join-Path -Path $env:LOCALAPPDATA -ChildPath '1C\1cv8'),
    [switch] $ShowOrphans,
    [Alias('h')]
    [switch] $Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-1CTextFile {
    <#
    .SYNOPSIS
        Читает текстовый файл 1С с учётом его кодировки.

    .DESCRIPTION
        Списки баз (v8i) 1С хранит в UTF-8, а 1cestart.cfg — в UTF-16LE.
        Кодировка определяется по BOM; файл без BOM считается UTF-8,
        а если байты не являются корректным UTF-8 — читается как ANSI (cp1251).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $utf8 = New-Object System.Text.UTF8Encoding($false, $true)

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return $utf8.GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
    }

    try {
        return $utf8.GetString($bytes)
    }
    catch [System.Text.DecoderFallbackException] {
        return [System.Text.Encoding]::GetEncoding(1251).GetString($bytes)
    }
}

function Get-1CIBases {
    <#
    .SYNOPSIS
        Читает список баз 1С из файла ibases.v8i.

    .OUTPUTS
        Объекты с полями Name (наименование), Id (GUID) и Connect (строка подключения).
        Секции-папки (без строки Connect) и базы без идентификатора в результат не попадают.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Не найден файл со списком баз: $Path"
    }

    $fullPath = (Resolve-Path -LiteralPath $Path).ProviderPath
    $content = Read-1CTextFile -Path $fullPath

    # Файл состоит из секций: заголовок [Наименование] и строки "Ключ=Значение".
    $sections = New-Object 'System.Collections.Generic.List[object]'
    $section = $null

    foreach ($line in ($content -split '\r?\n')) {
        $line = $line.Trim()
        if ($line -eq '') { continue }

        if ($line -match '^\[(?<name>.+)\]$') {
            if ($null -ne $section) { $sections.Add($section) }
            $section = @{
                Name       = $Matches['name'].Trim()
                Properties = @{}
            }
            continue
        }

        if ($null -eq $section) { continue }
        if ($line -match '^(?<key>[^=]+?)\s*=\s*(?<value>.*)$') {
            # Ключи в файле не зависят от регистра.
            $section.Properties[$Matches['key'].Trim().ToLowerInvariant()] = $Matches['value'].Trim()
        }
    }
    if ($null -ne $section) { $sections.Add($section) }

    foreach ($item in $sections) {
        # Папки дерева баз не имеют строки Connect, поэтому в результат не попадают.
        $connect = $item.Properties['connect']
        if ([string]::IsNullOrWhiteSpace($connect)) { continue }

        # Базы без идентификатора (GUID) молча пропускаем.
        $id = [guid]::Empty
        if (-not [guid]::TryParse([string]$item.Properties['id'], [ref]$id)) { continue }

        [pscustomobject]@{
            Name    = $item.Name
            Id      = $id
            Connect = $connect
        }
    }
}

function Get-1CCommonBasesListPaths {
    <#
    .SYNOPSIS
        Возвращает пути общих списков баз из файла 1cestart.cfg.

    .DESCRIPTION
        Параметр CommonInfoBases может встречаться в файле несколько раз,
        а внутри одной строки пути могут быть перечислены через точку с запятой.
        Пути могут быть локальными и сетевыми (UNC). Недоступность самих списков
        здесь не проверяется — это делает вызывающий код.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }

    try {
        $content = Read-1CTextFile -Path (Resolve-Path -LiteralPath $Path).ProviderPath
    }
    catch {
        # Файл настроек недоступен — считаем, что общих списков нет.
        return
    }

    foreach ($line in ($content -split '\r?\n')) {
        $line = $line.Trim()
        if ($line -eq '' -or $line.StartsWith(';')) { continue }

        if ($line -match '^CommonInfoBases\s*=\s*(?<value>.*)$') {
            foreach ($entry in ($Matches['value'] -split ';')) {
                $listPath = $entry.Trim().Trim('"').Trim()
                if ($listPath -ne '') { $listPath }
            }
        }
    }
}

function Get-1CFolderSizeBytes {
    <#
    .SYNOPSIS
        Считает суммарный размер файлов в каталоге (рекурсивно).

    .DESCRIPTION
        Недоступные файлы и подкаталоги пропускаются. Для отсутствующего каталога
        возвращается 0.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return [long] 0 }

    [long] $size = 0
    Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
        ForEach-Object { $size += $_.Length }

    return $size
}

function Get-1CCacheFolders {
    <#
    .SYNOPSIS
        Перечисляет папки кэша, чьё имя является GUID.

    .DESCRIPTION
        Служебные каталоги (logs, tmplts, conf и т.п.) отбрасываются,
        так как их имена не являются идентификаторами баз.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return }

    foreach ($item in @(Get-ChildItem -LiteralPath $Path -Force -Directory -ErrorAction SilentlyContinue)) {
        $id = [guid]::Empty
        if (-not [guid]::TryParse($item.Name, [ref]$id)) { continue }

        [pscustomobject]@{
            Id   = $id
            Name = $item.Name
            Path = $item.FullName
        }
    }
}

function Get-1CState {
    <#
    .SYNOPSIS
        Собирает список баз с размерами кэша и список осиротевших папок кэша.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $IbasesPath,
        [Parameter(Mandatory = $true)]
        [string] $StartCfgPath,
        [Parameter(Mandatory = $true)]
        [string] $UserCachePath,
        [Parameter(Mandatory = $true)]
        [string] $ProgramCachePath
    )

    # Базы из личного списка пользователя, затем из общих списков.
    # Одна и та же база может встречаться в нескольких списках — оставляем первое вхождение.
    $bases = New-Object 'System.Collections.Generic.List[object]'
    $knownIds = New-Object 'System.Collections.Generic.HashSet[guid]'

    foreach ($base in @(Get-1CIBases -Path $IbasesPath)) {
        if ($knownIds.Add($base.Id)) { $bases.Add($base) }
    }

    foreach ($listPath in @(Get-1CCommonBasesListPaths -Path $StartCfgPath)) {
        try {
            foreach ($base in @(Get-1CIBases -Path $listPath)) {
                if ($knownIds.Add($base.Id)) { $bases.Add($base) }
            }
        }
        catch {
            # Общий список недоступен (сеть, права, файл удалён) — молча пропускаем.
        }
    }

    # Сортируем по наименованию (с учётом текущей локали) и считаем размеры кэша.
    $bases = @(
        foreach ($base in ($bases | Sort-Object -Property Name)) {
            $userFolder = Join-Path -Path $UserCachePath -ChildPath ([string] $base.Id)
            $programFolder = Join-Path -Path $ProgramCachePath -ChildPath ([string] $base.Id)

            [pscustomobject]@{
                Name             = $base.Name
                Id               = $base.Id
                Connect          = $base.Connect
                UserCacheSize    = Get-1CFolderSizeBytes -Path $userFolder
                ProgramCacheSize = Get-1CFolderSizeBytes -Path $programFolder
                HasUserCache     = Test-Path -LiteralPath $userFolder -PathType Container
                HasProgramCache  = Test-Path -LiteralPath $programFolder -PathType Container
            }
        }
    )

    # В список выводим только базы, у которых есть папка кэша (пользовательского или программного).
    $bases = @($bases | Where-Object { $_.HasUserCache -or $_.HasProgramCache })

    # Осиротевшие папки кэша — те, что не соответствуют ни одной найденной базе.
    $orphans = New-Object 'System.Collections.Generic.List[object]'
    foreach ($cache in @(
            [pscustomobject]@{ Kind = 'пользовательский'; Path = $UserCachePath }
            [pscustomobject]@{ Kind = 'программный'; Path = $ProgramCachePath }
        )) {
        foreach ($folder in @(Get-1CCacheFolders -Path $cache.Path)) {
            if ($knownIds.Contains($folder.Id)) { continue }
            $orphans.Add([pscustomobject]@{
                    Id   = $folder.Name
                    Kind = $cache.Kind
                    Size = Get-1CFolderSizeBytes -Path $folder.Path
                })
        }
    }
    $orphans = @($orphans | Sort-Object -Property Id, Kind)

    return [pscustomobject]@{
        Bases   = @($bases)
        Orphans = @($orphans)
    }
}

function Remove-1CCacheFolder {
    <#
    .SYNOPSIS
        Удаляет папку кэша одной базы.

    .DESCRIPTION
        Удаляется только каталог, который лежит прямо в каталоге кэша и имя которого
        совпадает с идентификатором базы. Если папки нет, возвращается $false.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $CacheRoot,
        [Parameter(Mandatory = $true)]
        [guid] $Id
    )

    $path = Join-Path -Path $CacheRoot -ChildPath ([string] $Id)
    if (-not (Test-Path -LiteralPath $path -PathType Container)) { return $false }

    $rootFull = [System.IO.Path]::GetFullPath($CacheRoot).TrimEnd('\', '/')
    $pathFull = [System.IO.Path]::GetFullPath($path)
    $expected = $rootFull + [System.IO.Path]::DirectorySeparatorChar + [string] $Id

    if (-not [string]::Equals($pathFull, $expected, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Неожиданный путь кэша: $pathFull"
    }

    Remove-Item -LiteralPath $pathFull -Recurse -Force
    return $true
}

function Invoke-1CCacheCleanup {
    <#
    .SYNOPSIS
        Очищает кэш указанных баз.

    .PARAMETER Kind
        All — программный и пользовательский, Program — только программный,
        User — только пользовательский.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $Bases,
        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Program', 'User')]
        [string] $Kind,
        [Parameter(Mandatory = $true)]
        [string] $UserCachePath,
        [Parameter(Mandatory = $true)]
        [string] $ProgramCachePath
    )

    $removed = 0
    $failed = 0
    [long] $freed = 0

    foreach ($base in $Bases) {
        $targets = New-Object 'System.Collections.Generic.List[object]'
        if ($Kind -eq 'All' -or $Kind -eq 'User') {
            $targets.Add([pscustomobject]@{ Root = $UserCachePath; Size = $base.UserCacheSize })
        }
        if ($Kind -eq 'All' -or $Kind -eq 'Program') {
            $targets.Add([pscustomobject]@{ Root = $ProgramCachePath; Size = $base.ProgramCacheSize })
        }

        foreach ($target in $targets) {
            try {
                if (Remove-1CCacheFolder -CacheRoot $target.Root -Id $base.Id) {
                    $removed++
                    $freed += $target.Size
                }
            }
            catch {
                $failed++
            }
        }
    }

    return [pscustomobject]@{
        Removed = $removed
        Failed  = $failed
        Freed   = $freed
    }
}

function Format-1CSize {
    <#
    .SYNOPSIS
        Форматирует размер в мегабайтах для вывода.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [long] $Bytes
    )

    return ('{0:N1} MB' -f ($Bytes / 1MB))
}

function Format-1CFolderCount {
    <#
    .SYNOPSIS
        Склоняет слово "папка" для указанного количества.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int] $Count
    )

    $mod10 = $Count % 10
    $mod100 = $Count % 100

    if ($mod10 -eq 1 -and $mod100 -ne 11) { return ('{0} папка' -f $Count) }
    if ($mod10 -ge 2 -and $mod10 -le 4 -and ($mod100 -lt 12 -or $mod100 -gt 14)) { return ('{0} папки' -f $Count) }
    return ('{0} папок' -f $Count)
}

function Read-1CChoice {
    <#
    .SYNOPSIS
        Читает пункт меню и проверяет, что он есть в списке допустимых.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]] $Allowed,
        [string] $Prompt = 'Выбор'
    )

    for ($attempt = 1; $attempt -le 20; $attempt++) {
        $answer = Read-Host $Prompt
        if ($null -eq $answer) { return '0' }

        $answer = $answer.Trim()
        if ($Allowed -contains $answer) { return $answer }

        Write-Host 'Неверный ввод.'
    }

    return '0'
}

function Show-1COrphanDetails {
    <#
    .SYNOPSIS
        Показывает осиротевшие папки кэша подробно: по идентификаторам.
    #>
    [CmdletBinding()]
    param(
        [object[]] $Orphans
    )

    if ($Orphans.Count -eq 0) { return }

    [long] $size = 0
    foreach ($orphan in $Orphans) { $size += $orphan.Size }

    $kindWidth = 0
    $sizeWidth = 0
    foreach ($orphan in $Orphans) {
        if ($orphan.Kind.Length -gt $kindWidth) { $kindWidth = $orphan.Kind.Length }
        $text = Format-1CSize -Bytes $orphan.Size
        if ($text.Length -gt $sizeWidth) { $sizeWidth = $text.Length }
    }

    Write-Host
    Write-Host ('Осиротевшие базы: {0}, {1}.' -f (Format-1CFolderCount -Count $Orphans.Count), (Format-1CSize -Bytes $size))
    Write-Host
    foreach ($orphan in $Orphans) {
        Write-Host ('  {0}  {1}  {2}' -f $orphan.Id, $orphan.Kind.PadRight($kindWidth), (Format-1CSize -Bytes $orphan.Size).PadLeft($sizeWidth))
    }
}

function Get-1CRunningClientNames {
    <#
    .SYNOPSIS
        Возвращает имена запущенных клиентов 1С (1cv8.exe, 1cv8c.exe).
    #>
    [CmdletBinding()]
    param()

    return @(Get-Process -Name '1cv8', '1cv8c' -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty ProcessName -Unique)
}

function Show-1CMainMenu {
    <#
    .SYNOPSIS
        Показывает главное меню.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object] $State
    )

    [long] $programTotal = 0
    [long] $userTotal = 0
    foreach ($base in $State.Bases) {
        $programTotal += $base.ProgramCacheSize
        $userTotal += $base.UserCacheSize
    }

    [long] $orphansSize = 0
    foreach ($orphan in $State.Orphans) { $orphansSize += $orphan.Size }

    # Предупреждаем, если клиент 1С запущен: часть файлов кэша может быть занята.
    $running = @(Get-1CRunningClientNames)
    if ($running.Count -gt 0) {
        Write-Host
        Write-Host ('ВНИМАНИЕ: 1С запущена ({0}). Закройте 1С — иначе часть кэша может не удалиться.' -f ($running -join ', ')) -ForegroundColor Yellow
    }

    Write-Host
    Write-Host ('Базы с кэшем: {0}.' -f $State.Bases.Count)
    if ($State.Orphans.Count -gt 0) {
        Write-Host ('Осиротевшие базы: {0}, {1}.' -f (Format-1CFolderCount -Count $State.Orphans.Count), (Format-1CSize -Bytes $orphansSize))
    }
    Write-Host
    Write-Host ('1. Очистить всё ({0})' -f (Format-1CSize -Bytes ($programTotal + $userTotal)))
    Write-Host ('2. Очистить программный кэш ({0})' -f (Format-1CSize -Bytes $programTotal))
    Write-Host ('3. Очистить пользовательский кэш ({0})' -f (Format-1CSize -Bytes $userTotal))
    Write-Host '4. Выбрать базу...'
    Write-Host '0. Выход.'
    Write-Host
}

function Show-1CBaseList {
    <#
    .SYNOPSIS
        Показывает пронумерованный список баз с размерами обоих кэшей.
    #>
    [CmdletBinding()]
    param(
        [object[]] $Bases
    )

    if ($Bases.Count -eq 0) {
        Write-Host
        Write-Host 'Базы с кэшем не найдены.'
        Write-Host
        Write-Host '0. Назад'
        Write-Host
        return
    }

    $nameWidth = 0
    $programWidth = 'Программный'.Length
    $userWidth = 'Пользовательский'.Length
    foreach ($base in $Bases) {
        if ($base.Name.Length -gt $nameWidth) { $nameWidth = $base.Name.Length }

        $programText = Format-1CSize -Bytes $base.ProgramCacheSize
        if ($programText.Length -gt $programWidth) { $programWidth = $programText.Length }

        $userText = Format-1CSize -Bytes $base.UserCacheSize
        if ($userText.Length -gt $userWidth) { $userWidth = $userText.Length }
    }

    Write-Host
    Write-Host 'Размер кэша указан в МБ: программный и пользовательский.'
    Write-Host
    Write-Host ('      {0}  {1}  {2}' -f 'База'.PadRight($nameWidth), 'Программный'.PadLeft($programWidth), 'Пользовательский'.PadLeft($userWidth))

    for ($i = 0; $i -lt $Bases.Count; $i++) {
        Write-Host ('{0,4}. {1}  {2}  {3}' -f ($i + 1), $Bases[$i].Name.PadRight($nameWidth),
            (Format-1CSize -Bytes $Bases[$i].ProgramCacheSize).PadLeft($programWidth),
            (Format-1CSize -Bytes $Bases[$i].UserCacheSize).PadLeft($userWidth))
    }

    Write-Host
    Write-Host '0. Назад'
    Write-Host
}

function Show-1CBaseMenu {
    <#
    .SYNOPSIS
        Показывает меню действий для одной базы.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object] $Base
    )

    Write-Host
    Write-Host ('База: {0}' -f $Base.Name)
    Write-Host
    Write-Host '1. Очистить всё'
    Write-Host ('2. Очистить программный кэш ({0})' -f (Format-1CSize -Bytes $Base.ProgramCacheSize))
    Write-Host ('3. Очистить пользовательский кэш ({0})' -f (Format-1CSize -Bytes $Base.UserCacheSize))
    Write-Host '0. Назад'
    Write-Host
}

function Show-1CCleanupResult {
    <#
    .SYNOPSIS
        Показывает результат очистки кэша.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object] $Result
    )

    if ($Result.Removed -eq 0 -and $Result.Failed -eq 0) {
        Write-Host 'Нечего очищать.'
    }
    else {
        Write-Host ('Очищено: {0}, освобождено {1}.' -f (Format-1CFolderCount -Count $Result.Removed), (Format-1CSize -Bytes $Result.Freed))
        if ($Result.Failed -gt 0) {
            Write-Host ('Не удалось удалить: {0}.' -f (Format-1CFolderCount -Count $Result.Failed)) -ForegroundColor Yellow
        }
    }
}

function Wait-AnyKey {
    <#
    .SYNOPSIS
        Ждёт нажатия любой клавиши.
    #>
    [CmdletBinding()]
    param(
        [string] $Prompt = 'Нажмите любую клавишу для выхода...'
    )

    if ($Prompt) { Write-Host $Prompt -NoNewline }

    try {
        [void]$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
    catch {
        # Хост без консоли (перенаправленный ввод) — ждём Enter.
        try { [void](Read-Host) } catch { }
    }
    Write-Host
}

if ($Help) {
    try {
        Get-Help -Name $PSCommandPath -Detailed
    }
    catch {
        Write-Host 'Использование: .\1c-cache-cleanup.ps1 [-IbasesPath <путь>] [-StartCfgPath <путь>] [-UserCachePath <путь>] [-ProgramCachePath <путь>] [-ShowOrphans] [-h]'
    }
    exit 0
}

$config = @{
    IbasesPath      = $IbasesPath
    StartCfgPath    = $StartCfgPath
    UserCachePath   = $UserCachePath
    ProgramCachePath = $ProgramCachePath
}

try {
    $state = Get-1CState @config

    if ($ShowOrphans) { Show-1COrphanDetails -Orphans $state.Orphans }

    :main while ($true) {
        if ($state.Bases.Count -eq 0) {
            Write-Host
            Write-Host 'Базы с кэшем не найдены.'
            Write-Host
            break main
        }

        Show-1CMainMenu -State $state
        $choice = Read-1CChoice -Allowed @('0', '1', '2', '3', '4')

        if ($choice -eq '0') { break main }

        if ($choice -eq '4') {
            :baseList while ($true) {
                Show-1CBaseList -Bases $state.Bases

                $allowed = @('0')
                for ($i = 1; $i -le $state.Bases.Count; $i++) { $allowed += [string] $i }
                $answer = Read-1CChoice -Allowed $allowed -Prompt 'Номер базы (0 — назад)'

                if ($answer -eq '0') { break baseList }

                $base = $state.Bases[[int] $answer - 1]

                :baseMenu while ($true) {
                    Show-1CBaseMenu -Base $base
                    $action = Read-1CChoice -Allowed @('0', '1', '2', '3')

                    if ($action -eq '0') { break baseMenu }

                    $kind = switch ($action) {
                        '1' { 'All' }
                        '2' { 'Program' }
                        '3' { 'User' }
                    }

                    Show-1CCleanupResult -Result (Invoke-1CCacheCleanup -Bases @($base) -Kind $kind -UserCachePath $UserCachePath -ProgramCachePath $ProgramCachePath)
                    $state = Get-1CState @config

                    # Список баз мог измениться — возвращаемся к нему.
                    break baseMenu
                }
            }
            continue main
        }

        $kind = switch ($choice) {
            '1' { 'All' }
            '2' { 'Program' }
            '3' { 'User' }
        }

        Show-1CCleanupResult -Result (Invoke-1CCacheCleanup -Bases $state.Bases -Kind $kind -UserCachePath $UserCachePath -ProgramCachePath $ProgramCachePath)
        $state = Get-1CState @config
    }
}
catch {
    Write-Host
    Write-Host ('Ошибка: {0}' -f $_.Exception.Message) -ForegroundColor Red
    Write-Host
    Wait-AnyKey
    exit 1
}
