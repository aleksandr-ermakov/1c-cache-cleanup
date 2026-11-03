<#
.SYNOPSIS
    Скрипт выборочной очистки кэша 1С.

.DESCRIPTION
    Скрипт очищает кэш информационных баз 1С под Windows.
    Без параметров показывает базы, размер их кэша и меню действий.
    Кэш можно удалить полностью, только программный или только пользовательский —
    у всех найденных баз, у одной выбранной или у осиротевших баз.
    С параметрами -ClearAll, -ClearProgramCache, -ClearUserCache и -ClearOrphaned
    очистка выполняется сразу, без меню.

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

.PARAMETER ClearAll
    Удаляет программный и пользовательский кэш всех баз, без меню. Синоним: -a.

.PARAMETER ClearProgramCache
    Удаляет программный кэш всех баз, без меню. Синоним: -p.

.PARAMETER ClearUserCache
    Удаляет пользовательский кэш всех баз, без меню. Синоним: -u.

.PARAMETER ClearOrphaned
    Удаляет весь кэш осиротевших баз, без меню. Синоним: -o.

.PARAMETER Help
    Выводит эту справку и завершает работу. То же самое делает ключ -h.

.EXAMPLE
    .\1c-cache-cleanup.ps1
    Без параметров показывает меню с размерами кэша.

.EXAMPLE
    .\1c-cache-cleanup.ps1 -ClearAll
    Удаляет программный и пользовательский кэш всех баз (синоним -a).

.EXAMPLE
    .\1c-cache-cleanup.ps1 -ClearProgramCache
    Удаляет программный кэш всех баз (синоним -p).

.EXAMPLE
    .\1c-cache-cleanup.ps1 -ClearUserCache
    Удаляет пользовательский кэш всех баз (синоним -u).

.EXAMPLE
    .\1c-cache-cleanup.ps1 -ClearOrphaned
    Удаляет весь кэш осиротевших баз (синоним -o).

.EXAMPLE
    .\1c-cache-cleanup.ps1 -h
    Показывает справку по скрипту.
#>

[CmdletBinding()]
param(
    [string] $IbasesPath = (Join-Path -Path $env:APPDATA -ChildPath '1C\1CEStart\ibases.v8i'),
    [string] $StartCfgPath = (Join-Path -Path $env:APPDATA -ChildPath '1C\1CEStart\1cestart.cfg'),
    [string] $UserCachePath = (Join-Path -Path $env:APPDATA -ChildPath '1C\1cv8'),
    [string] $ProgramCachePath = (Join-Path -Path $env:LOCALAPPDATA -ChildPath '1C\1cv8'),
    [Alias('a')]
    [switch] $ClearAll,
    [Alias('p')]
    [switch] $ClearProgramCache,
    [Alias('u')]
    [switch] $ClearUserCache,
    [Alias('o')]
    [switch] $ClearOrphaned,
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
        Собирает список баз с размерами кэша и список осиротевших баз.
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

    # Осиротевшие базы — папки кэша, для которых нет зарегистрированной базы.
    # Одна база может иметь папки и в пользовательском, и в программном кэше.
    $orphansById = @{}
    foreach ($cache in @(
            [pscustomobject]@{ IsUser = $true; Path = $UserCachePath }
            [pscustomobject]@{ IsUser = $false; Path = $ProgramCachePath }
        )) {
        foreach ($folder in @(Get-1CCacheFolders -Path $cache.Path)) {
            if ($knownIds.Contains($folder.Id)) { continue }

            $key = [string] $folder.Id
            if (-not $orphansById.ContainsKey($key)) {
                $orphansById[$key] = [pscustomobject]@{
                    Name             = $key
                    Id               = $folder.Id
                    UserCacheSize    = [long] 0
                    ProgramCacheSize = [long] 0
                }
            }

            $entry = $orphansById[$key]
            if ($cache.IsUser) { $entry.UserCacheSize = Get-1CFolderSizeBytes -Path $folder.Path }
            else { $entry.ProgramCacheSize = Get-1CFolderSizeBytes -Path $folder.Path }
        }
    }

    # Осиротевшие базы сортируем по общему размеру кэша по убыванию.
    $orphans = @(
        foreach ($entry in $orphansById.Values) {
            [pscustomobject]@{
                Name             = $entry.Name
                Id               = $entry.Id
                UserCacheSize    = $entry.UserCacheSize
                ProgramCacheSize = $entry.ProgramCacheSize
                TotalCacheSize   = $entry.UserCacheSize + $entry.ProgramCacheSize
            }
        }
    )
    $orphans = @($orphans | Sort-Object -Property @{ Expression = { $_.TotalCacheSize }; Descending = $true }, @{ Expression = { $_.Name } })

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
        [AllowEmptyCollection()]
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

    # Округляем вручную: PowerShell 5.1 и 7 по-разному обрабатывают ровно половину.
    $megabytes = [math]::Round($Bytes / 1MB, 1, [System.MidpointRounding]::AwayFromZero)
    return ('{0:N1} MB' -f $megabytes)
}

function Show-1CCacheTotals {
    <#
    .SYNOPSIS
        Показывает общий размер программного и пользовательского кэша.
    #>
    [CmdletBinding()]
    param(
        [object[]] $Bases
    )

    [long] $programTotal = 0
    [long] $userTotal = 0
    foreach ($base in $Bases) {
        $programTotal += $base.ProgramCacheSize
        $userTotal += $base.UserCacheSize
    }

    Write-Host ('Всего кэша: программный {0}, пользовательский {1}.' -f (Format-1CSize -Bytes $programTotal), (Format-1CSize -Bytes $userTotal))
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

function Show-1COrphanList {
    <#
    .SYNOPSIS
        Показывает список осиротевших баз с размерами кэша и общий размер.
    #>
    [CmdletBinding()]
    param(
        [object[]] $Orphans
    )

    if ($Orphans.Count -eq 0) { return }

    $programWidth = 'Программный'.Length
    $userWidth = 'Пользовательский'.Length

    Write-Host
    Write-Host ('Осиротевшие базы: {0}.' -f $Orphans.Count)
    Write-Host 'Для этих баз нет зарегистрированной информационной базы.'
    Write-Host 'Размер кэша указан в МБ: программный и пользовательский.'
    Write-Host
    Write-Host ('  {0}  {1}  {2}' -f 'База'.PadRight(36), 'Программный'.PadLeft($programWidth), 'Пользовательский'.PadLeft($userWidth))
    foreach ($base in $Orphans) {
        Write-Host ('  {0}  {1}  {2}' -f $base.Name, (Format-1CSize -Bytes $base.ProgramCacheSize).PadLeft($programWidth), (Format-1CSize -Bytes $base.UserCacheSize).PadLeft($userWidth))
    }
    Write-Host
    Show-1CCacheTotals -Bases $Orphans
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
        Показывает главное меню и возвращает список допустимых пунктов.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object] $State
    )

    $allowed = @('0')

    [long] $programTotal = 0
    [long] $userTotal = 0
    foreach ($base in $State.Bases) {
        $programTotal += $base.ProgramCacheSize
        $userTotal += $base.UserCacheSize
    }

    [long] $orphansSize = 0
    foreach ($base in $State.Orphans) { $orphansSize += $base.TotalCacheSize }

    # Предупреждаем, если клиент 1С запущен: часть файлов кэша может быть занята.
    $running = @(Get-1CRunningClientNames)
    if ($running.Count -gt 0) {
        Write-Host
        Write-Host ('ВНИМАНИЕ: 1С запущена ({0}). Закройте 1С — иначе часть кэша может не удалиться.' -f ($running -join ', ')) -ForegroundColor Yellow
    }

    Write-Host
    if ($State.Bases.Count -gt 0) {
        Write-Host ('Базы с кэшем: {0}.' -f $State.Bases.Count)
    }
    else {
        Write-Host 'Базы с кэшем не найдены.'
    }
    if ($State.Orphans.Count -gt 0) {
        Write-Host ('Осиротевшие базы: {0}.' -f $State.Orphans.Count)
    }

    Write-Host
    Write-Host 'Выберите действие с кэшем:'
    if ($State.Bases.Count -gt 0) {
        Write-Host ('1. Очистить всё ({0})' -f (Format-1CSize -Bytes ($programTotal + $userTotal)))
        Write-Host ('2. Очистить программный кэш ({0})' -f (Format-1CSize -Bytes $programTotal))
        Write-Host ('3. Очистить пользовательский кэш ({0})' -f (Format-1CSize -Bytes $userTotal))
        Write-Host '4. Выбрать базу...'
        $allowed += @('1', '2', '3', '4')
    }
    if ($State.Orphans.Count -gt 0) {
        Write-Host '5. Осиротевшие базы...'
        $allowed += '5'
    }
    Write-Host '0. Выход.'
    Write-Host

    return $allowed
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
    Show-1CCacheTotals -Bases $Bases
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
        return
    }

    Write-Host ('Очищено: {0}, освобождено {1}.' -f (Format-1CFolderCount -Count $Result.Removed), (Format-1CSize -Bytes $Result.Freed))

    if ($Result.Failed -gt 0) {
        Write-Host ('Не удалось удалить: {0}.' -f (Format-1CFolderCount -Count $Result.Failed)) -ForegroundColor Yellow
    }
}

function Show-1CCleanupResultLine {
    <#
    .SYNOPSIS
        Показывает результат очистки кэша строкой с заголовком.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Title,
        [Parameter(Mandatory = $true)]
        [object] $Result
    )

    if ($Result.Removed -eq 0 -and $Result.Failed -eq 0) {
        Write-Host ('{0}: очищать нечего.' -f $Title)
        return
    }

    Write-Host ('{0}: очищено {1}, освобождено {2}.' -f $Title, (Format-1CFolderCount -Count $Result.Removed), (Format-1CSize -Bytes $Result.Freed))
    if ($Result.Failed -gt 0) {
        Write-Host ('{0}: не удалось удалить {1}.' -f $Title, (Format-1CFolderCount -Count $Result.Failed)) -ForegroundColor Yellow
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
        Write-Host 'Использование: .\1c-cache-cleanup.ps1 [-IbasesPath <путь>] [-StartCfgPath <путь>] [-UserCachePath <путь>] [-ProgramCachePath <путь>] [-ClearAll] [-ClearProgramCache] [-ClearUserCache] [-ClearOrphaned] [-h]'
    }
    exit 0
}

$script:IbasesPath = $IbasesPath
$script:StartCfgPath = $StartCfgPath
$script:UserCachePath = $UserCachePath
$script:ProgramCachePath = $ProgramCachePath
$script:State = $null

function Update-1CState {
    <#
    .SYNOPSIS
        Перечитывает список баз, осиротевшие базы и размеры кэша.
    #>
    $script:State = Get-1CState -IbasesPath $script:IbasesPath -StartCfgPath $script:StartCfgPath -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath
}

function Invoke-1CBaseSection {
    <#
    .SYNOPSIS
        Показывает список баз и меню действий для выбранной базы.
    #>
    while ($true) {
        $items = @($script:State.Bases)
        if ($items.Count -eq 0) { return }

        Show-1CBaseList -Bases $items

        $allowed = @('0')
        for ($i = 1; $i -le $items.Count; $i++) { $allowed += [string] $i }
        $answer = Read-1CChoice -Allowed $allowed -Prompt 'Номер базы (0 — назад)'
        if ($answer -eq '0') { return }

        $base = $items[[int] $answer - 1]

        while ($true) {
            Show-1CBaseMenu -Base $base
            $action = Read-1CChoice -Allowed @('0', '1', '2', '3')
            if ($action -eq '0') { break }

            $kind = switch ($action) {
                '1' { 'All' }
                '2' { 'Program' }
                '3' { 'User' }
            }

            Show-1CCleanupResult -Result (Invoke-1CCacheCleanup -Bases @($base) -Kind $kind -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath)
            Update-1CState

            # Список баз мог измениться — возвращаемся к нему.
            break
        }
    }
}

function Invoke-1COrphanSection {
    <#
    .SYNOPSIS
        Показывает список осиротевших баз и действия с ними.
    #>
    while ($true) {
        $items = @($script:State.Orphans)
        if ($items.Count -eq 0) {
            Write-Host
            Write-Host 'Осиротевшие базы не найдены.'
            Write-Host
            return
        }

        Show-1COrphanList -Orphans $items
        Write-Host
        Write-Host 'Выберите действие с кэшем осиротевших баз:'
        Write-Host '1. Очистить всё'
        Write-Host '2. Очистить программный кэш'
        Write-Host '3. Очистить пользовательский кэш'
        Write-Host '0. Назад'
        Write-Host

        $choice = Read-1CChoice -Allowed @('0', '1', '2', '3')
        if ($choice -eq '0') { return }

        $kind = switch ($choice) {
            '1' { 'All' }
            '2' { 'Program' }
            '3' { 'User' }
        }

        Show-1CCleanupResult -Result (Invoke-1CCacheCleanup -Bases $items -Kind $kind -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath)
        Update-1CState
    }
}

function Invoke-1CMainMenu {
    <#
    .SYNOPSIS
        Показывает главное меню и выполняет выбранные действия.
    #>
    while ($true) {
        if ($script:State.Bases.Count -eq 0 -and $script:State.Orphans.Count -eq 0) {
            Write-Host
            Write-Host 'Базы с кэшем не найдены.'
            Write-Host
            return
        }

        $allowed = Show-1CMainMenu -State $script:State
        $choice = Read-1CChoice -Allowed $allowed
        if ($choice -eq '0') { return }

        switch ($choice) {
            '1' {
                Show-1CCleanupResult -Result (Invoke-1CCacheCleanup -Bases $script:State.Bases -Kind All -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath)
                Update-1CState
            }
            '2' {
                Show-1CCleanupResult -Result (Invoke-1CCacheCleanup -Bases $script:State.Bases -Kind Program -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath)
                Update-1CState
            }
            '3' {
                Show-1CCleanupResult -Result (Invoke-1CCacheCleanup -Bases $script:State.Bases -Kind User -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath)
                Update-1CState
            }
            '4' { Invoke-1CBaseSection }
            '5' { Invoke-1COrphanSection }
        }
    }
}

function Invoke-1CCleanupCommands {
    <#
    .SYNOPSIS
        Выполняет очистку по параметрам командной строки, без меню.

    .OUTPUTS
        Количество папок, которые не удалось удалить.
    #>
    [CmdletBinding()]
    param(
        [switch] $ClearAll,
        [switch] $ClearProgramCache,
        [switch] $ClearUserCache,
        [switch] $ClearOrphaned
    )

    $failed = 0

    # Предупреждаем, если клиент 1С запущен: часть файлов кэша может быть занята.
    $running = @(Get-1CRunningClientNames)
    if ($running.Count -gt 0) {
        Write-Host ('ВНИМАНИЕ: 1С запущена ({0}). Часть кэша может не удалиться.' -f ($running -join ', ')) -ForegroundColor Yellow
    }

    if ($ClearAll) {
        $result = Invoke-1CCacheCleanup -Bases $script:State.Bases -Kind All -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath
        Show-1CCleanupResultLine -Title 'Кэш всех баз' -Result $result
        $failed += $result.Failed
        Update-1CState
    }

    if ($ClearProgramCache) {
        $result = Invoke-1CCacheCleanup -Bases $script:State.Bases -Kind Program -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath
        Show-1CCleanupResultLine -Title 'Программный кэш всех баз' -Result $result
        $failed += $result.Failed
        Update-1CState
    }

    if ($ClearUserCache) {
        $result = Invoke-1CCacheCleanup -Bases $script:State.Bases -Kind User -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath
        Show-1CCleanupResultLine -Title 'Пользовательский кэш всех баз' -Result $result
        $failed += $result.Failed
        Update-1CState
    }

    if ($ClearOrphaned) {
        $result = Invoke-1CCacheCleanup -Bases $script:State.Orphans -Kind All -UserCachePath $script:UserCachePath -ProgramCachePath $script:ProgramCachePath
        Show-1CCleanupResultLine -Title 'Кэш осиротевших баз' -Result $result
        $failed += $result.Failed
        Update-1CState
    }

    return $failed
}

try {
    Update-1CState

    if ($ClearAll -or $ClearProgramCache -or $ClearUserCache -or $ClearOrphaned) {
        $failed = Invoke-1CCleanupCommands -ClearAll:$ClearAll -ClearProgramCache:$ClearProgramCache -ClearUserCache:$ClearUserCache -ClearOrphaned:$ClearOrphaned
        if ($failed -gt 0) { exit 1 }
        exit 0
    }

    Invoke-1CMainMenu
}
catch {
    Write-Host
    Write-Host ('Ошибка: {0}' -f $_.Exception.Message) -ForegroundColor Red
    Write-Host

    # В интерактивном режиме даём прочитать ошибку, в режиме очистки — сразу выходим.
    if (-not ($ClearAll -or $ClearProgramCache -or $ClearUserCache -or $ClearOrphaned)) {
        Wait-AnyKey
    }
    exit 1
}
