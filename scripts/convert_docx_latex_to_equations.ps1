param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,
    [string]$TargetPath,
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$wdAlignParagraphCenter = 1
$initialMathTypeIds = @(Get-Process MathType -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

function Resolve-DefaultOutputPath {
    param(
        [string]$InputPath,
        [string]$Suffix
    )

    $directory = Split-Path -Parent $InputPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $extension = [System.IO.Path]::GetExtension($InputPath)
    $suffixExtension = [System.IO.Path]::GetExtension($Suffix)

    if ([string]::IsNullOrWhiteSpace($suffixExtension)) {
        return (Join-Path $directory ($baseName + $Suffix + $extension))
    }

    return (Join-Path $directory ($baseName + $Suffix))
}

function Assert-FileIsClosed {
    param(
        [string]$Path,
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
    } catch {
        throw "$Label is locked. Close the document in Word and try again: $Path"
    }
}

$SourcePath = (Resolve-Path -LiteralPath $SourcePath).Path
if ([string]::IsNullOrWhiteSpace($TargetPath)) {
    $TargetPath = Resolve-DefaultOutputPath -InputPath $SourcePath -Suffix '_formula.docx'
}
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Resolve-DefaultOutputPath -InputPath $SourcePath -Suffix '_formula_log.json'
}

$TargetPath = [System.IO.Path]::GetFullPath($TargetPath)
$LogPath = [System.IO.Path]::GetFullPath($LogPath)

if ($SourcePath -eq $TargetPath) {
    throw 'TargetPath must be different from SourcePath.'
}

function New-LogEvent {
    param(
        [string]$Kind,
        [int]$ParagraphIndex,
        [string]$Engine,
        [string]$Status,
        [string]$Formula,
        [string]$Message = '',
        [object[]]$Segments = @()
    )

    [pscustomobject]@{
        kind = $Kind
        paragraphIndex = $ParagraphIndex
        engine = $Engine
        status = $Status
        formula = $Formula
        message = $Message
        segments = $Segments
    }
}

function Get-ParagraphBodyText {
    param($Paragraph)
    return ($Paragraph.Range.Text -replace "(\r|\a)+$", '')
}

function Normalize-FormulaText {
    param([string]$Text)
    $clean = $Text -replace '[\u0007\u000b]', ' '
    $clean = $clean -replace '\r|\n', ' '
    $clean = $clean -replace '\s+', ' '
    return $clean.Trim()
}

function Test-HasFormulaSignature {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    return ($Text -match '\\[A-Za-z]+' -or
            $Text -match '_[{A-Za-z0-9]' -or
            $Text -match '\^[{A-Za-z0-9]' -or
            $Text -match '[=<>]' -or
            $Text -match '[{}]')
}

function Test-IsAsciiIdentifierLike {
    param([string]$Text)
    $trimmed = $Text.Trim()
    return ($trimmed -match '^[A-Za-z][A-Za-z0-9_]*(?:\s*/\s*[A-Za-z][A-Za-z0-9_]*)*(?:\s+[\u4E00-\u9FFF]+)?$')
}

function Test-IsDisplayFormulaParagraph {
    param([string]$Text)
    $normalized = Normalize-FormulaText -Text $Text
    if (-not (Test-HasFormulaSignature -Text $normalized)) {
        return $false
    }

    if (Test-IsAsciiIdentifierLike -Text $normalized) {
        return $false
    }

    if ($normalized -match ';' -and $normalized -notmatch '\\|[{}]|\^') {
        return $false
    }

    if ($normalized -notmatch '\\|_|\^|[{}]') {
        return $false
    }

    if ($normalized -notmatch '[\u4E00-\u9FFF]') {
        return $true
    }

    return ($normalized -match '^(\\|[A-Za-z].*[=<>])')
}

function Test-ShouldSkipInlineMatch {
    param([string]$Text)
    $trimmed = $Text.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return $true
    }

    if ($trimmed -match '^(build|load|read|apply|initialize|finalize|perform|grow|remove|calibrate|integrate|mie)_[A-Za-z0-9_]*(?:\s*/\s*(build|load|read|apply|initialize|finalize|perform|grow|remove|calibrate|integrate|mie)_[A-Za-z0-9_]*)*$') {
        return $true
    }

    return (Test-IsAsciiIdentifierLike -Text $trimmed)
}

function Get-InlineFormulaMatches {
    param([string]$Text)

    $patterns = @(
        'ratio\^\\alpha',
        'Q_[A-Za-z0-9{}]+\\times\s*Q_[A-Za-z0-9{}]+(?:\\times\s*Q_[A-Za-z0-9{}]+)*',
        '[A-Za-z](?:_[A-Za-z0-9{}]+)+(?:\([^()\r\n]{0,80}\))?(?:\s*=\s*[^\u4E00-\u9FFF]{1,120})',
        '\\alpha\s*[<>=]\s*1',
        '\\alpha',
        'k_[0-9]+',
        'S_\{[^}]+\}'
    )

    $candidates = New-Object System.Collections.Generic.List[object]
    foreach ($pattern in $patterns) {
        foreach ($match in [regex]::Matches($Text, $pattern)) {
            $value = $match.Value.Trim()
            if (-not (Test-ShouldSkipInlineMatch -Text $value)) {
                $candidates.Add([pscustomobject]@{
                    Start = $match.Index
                    Length = $match.Length
                    End = $match.Index + $match.Length
                    Text = $value
                })
            }
        }
    }

    $ordered = $candidates | Sort-Object @{ Expression = 'Start'; Descending = $false }, @{ Expression = 'Length'; Descending = $true }
    $selected = New-Object System.Collections.Generic.List[object]
    foreach ($candidate in $ordered) {
        $overlap = $false
        foreach ($existing in $selected) {
            if (-not ($candidate.End -le $existing.Start -or $candidate.Start -ge $existing.End)) {
                $overlap = $true
                break
            }
        }

        if (-not $overlap) {
            $selected.Add($candidate)
        }
    }

    return $selected | Sort-Object Start -Descending
}

function Wrap-ForMathType {
    param(
        [string]$Formula,
        [ValidateSet('Inline', 'Display')]
        [string]$Mode = 'Inline'
    )

    $trimmed = $Formula.Trim()
    if ($trimmed -match '^\$.*\$$') {
        return $trimmed
    }

    if ($Mode -eq 'Display') {
        return ('$$' + $trimmed + '$$')
    }

    return ('$' + $trimmed + '$')
}

function Remove-AdjacentLiteralDollars {
    param($InlineShape)

    $document = $InlineShape.Range.Document
    while ($true) {
        $removed = $false

        if ($InlineShape.Range.Start -gt 0) {
            $leading = $document.Range($InlineShape.Range.Start - 1, $InlineShape.Range.Start)
            if ($leading.Text -eq '$') {
                $leading.Text = ''
                $removed = $true
            }
        }

        $trailing = $document.Range($InlineShape.Range.End, $InlineShape.Range.End + 1)
        if ($trailing.Text -eq '$') {
            $trailing.Text = ''
            $removed = $true
        }

        if (-not $removed) {
            break
        }
    }
}

function Find-InlineShapeNearStart {
    param(
        $Document,
        [int]$Start,
        [int]$MaxStart
    )

    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $shape = $Document.InlineShapes.Item($i)
        if ($shape.Range.Start -ge $Start -and $shape.Range.Start -le $MaxStart) {
            return $shape
        }
    }

    return $null
}

function Test-HasCjkTextMacro {
    param([string]$Formula)
    return ($Formula -match '\\text\{[^{}]*[\u4E00-\u9FFF][^{}]*\}')
}

function Normalize-MathTypeMathSegment {
    param([string]$Text)
    $math = Normalize-FormulaText -Text $Text
    $math = $math -replace '\\ ', ' '
    $math = $math -replace '\s+', ' '
    $math = $math -replace '\\+$', ''
    return $math.Trim()
}

function Merge-AdjacentSegments {
    param([object[]]$Segments)

    $merged = New-Object System.Collections.Generic.List[object]
    foreach ($segment in $Segments) {
        if ([string]::IsNullOrWhiteSpace($segment.Text)) {
            continue
        }

        $kind = $segment.Kind
        $text = $segment.Text
        if ($kind -eq 'math' -and $text.Trim() -match '^[,;:]+$') {
            $kind = 'text'
        }

        if ($merged.Count -gt 0 -and $merged[$merged.Count - 1].Kind -eq $kind) {
            $merged[$merged.Count - 1].Text += $text
        } else {
            $merged.Add([pscustomobject]@{
                Kind = $kind
                Text = $text
            })
        }
    }

    return $merged
}

function Get-CjkFormulaSegments {
    param([string]$Formula)

    $normalized = Normalize-FormulaText -Text $Formula
    $tokenPattern = '\\text\{([^{}]*)\}|\\qquad|\\quad|\\,'
    $rawSegments = New-Object System.Collections.Generic.List[object]
    $cursor = 0

    foreach ($match in [regex]::Matches($normalized, $tokenPattern)) {
        if ($match.Index -gt $cursor) {
            $beforeText = $normalized.Substring($cursor, $match.Index - $cursor)
            $beforeText = Normalize-MathTypeMathSegment -Text $beforeText
            if (-not [string]::IsNullOrWhiteSpace($beforeText)) {
                $rawSegments.Add([pscustomobject]@{
                    Kind = 'math'
                    Text = $beforeText
                })
            }
        }

        if ($match.Groups[1].Success) {
            $rawSegments.Add([pscustomobject]@{
                Kind = 'text'
                Text = $match.Groups[1].Value
            })
        } else {
            $rawSegments.Add([pscustomobject]@{
                Kind = 'text'
                Text = ' '
            })
        }

        $cursor = $match.Index + $match.Length
    }

    if ($cursor -lt $normalized.Length) {
        $tailText = Normalize-MathTypeMathSegment -Text $normalized.Substring($cursor)
        if (-not [string]::IsNullOrWhiteSpace($tailText)) {
            $rawSegments.Add([pscustomobject]@{
                Kind = 'math'
                Text = $tailText
            })
        }
    }

    return Merge-AdjacentSegments -Segments $rawSegments.ToArray()
}

function Replace-Fractions {
    param([string]$Text)
    $current = $Text
    $pattern = '\\frac\{([^{}]+)\}\{([^{}]+)\}'
    while ($current -match $pattern) {
        $current = [regex]::Replace(
            $current,
            $pattern,
            {
                param($m)
                return '((' + $m.Groups[1].Value + ')/(' + $m.Groups[2].Value + '))'
            }
        )
    }

    return $current
}

function Convert-LatexToWordLinear {
    param([string]$Formula)

    $linear = Normalize-FormulaText -Text $Formula
    $linear = $linear.Trim('$')
    $linear = Replace-Fractions -Text $linear

    do {
        $previous = $linear
        $linear = [regex]::Replace($linear, '\\(?:mathrm|mathbf|mathcal|mathit|operatorname|text)\{([^{}]*)\}', '$1')
    } while ($linear -ne $previous)

    $linear = $linear -replace '\\begin\{cases\}', ''
    $linear = $linear -replace '\\end\{cases\}', ''
    $linear = $linear -replace '\\left', ''
    $linear = $linear -replace '\\right', ''
    $linear = $linear -replace '\\qquad', ' '
    $linear = $linear -replace '\\quad', ' '
    $linear = $linear -replace '\\,', ' '
    $linear = $linear -replace '\\!', ''
    $linear = $linear -replace '\\cdot', '*'
    $linear = $linear -replace '\\times', ' x '
    $linear = $linear -replace '\\geq?', '>='
    $linear = $linear -replace '\\leq?', '<='
    $linear = $linear -replace '\\neq', '!='
    $linear = $linear -replace '\\Rightarrow', '=>'
    $linear = $linear -replace '\\to', '->'
    $linear = $linear -replace '\\Re', 'Re'
    $linear = $linear -replace '\\Im', 'Im'
    $linear = [regex]::Replace($linear, '\\sqrt\{([^{}]+)\}', 'sqrt($1)')
    $linear = $linear -replace '\\\\', '; '
    $linear = $linear -replace '&', ' '
    $linear = $linear -replace '\{', '('
    $linear = $linear -replace '\}', ')'
    $linear = $linear -replace '\\', ''
    $linear = $linear -replace '\s+', ' '
    return $linear.Trim()
}

function Try-ConvertRangeWithMathType {
    param(
        $Word,
        $Range,
        [string]$Formula,
        [ValidateSet('Inline', 'Display')]
        [string]$Mode = 'Inline'
    )

    $document = $Range.Document
    $wrapped = Wrap-ForMathType -Formula $Formula -Mode $Mode
    $start = $Range.Start
    $beforeInlineShapes = $document.InlineShapes.Count

    $Range.Text = $wrapped
    $workingRange = $document.Range($start, $start + $wrapped.Length)
    $workingRange.Select()

    $callbackArg = $null
    try {
        $null = $Word.Run('MTCommand_OnTexToggle', [ref]$callbackArg)
    } catch {
        return [pscustomobject]@{
            Success = $false
            Message = $_.Exception.Message
            Text = $wrapped
            RangeStart = $start
            RangeEnd = $start + $wrapped.Length
        }
    }

    $elapsedMs = 0
    while ($elapsedMs -lt 6000) {
        Start-Sleep -Milliseconds 300
        $elapsedMs += 300
        if ($document.InlineShapes.Count -gt $beforeInlineShapes) {
            Start-Sleep -Milliseconds 800
            $newInlineShape = Find-InlineShapeNearStart -Document $document -Start $start -MaxStart ($start + $wrapped.Length + 2)
            if ($null -eq $newInlineShape) {
                $newInlineShape = $document.InlineShapes.Item($document.InlineShapes.Count)
            }
            if ($Mode -eq 'Display' -and $wrapped.StartsWith('$$')) {
                Remove-AdjacentLiteralDollars -InlineShape $newInlineShape
            }

            return [pscustomobject]@{
                Success = $true
                Message = 'MathType toggle succeeded.'
                Text = $wrapped
                RangeStart = $newInlineShape.Range.Start
                RangeEnd = $newInlineShape.Range.End
            }
        }
    }

    return [pscustomobject]@{
        Success = $false
        Message = 'MathType did not create an inline shape.'
        Text = $wrapped
        RangeStart = $start
        RangeEnd = $start + $wrapped.Length
    }
}

function Try-ConvertRangeWithWordEquation {
    param(
        $Range,
        [string]$Formula
    )

    $document = $Range.Document
    $linear = Convert-LatexToWordLinear -Formula $Formula
    if ([string]::IsNullOrWhiteSpace($linear)) {
        return [pscustomobject]@{
            Success = $false
            Message = 'Fallback linear text is empty.'
            Text = $linear
            RangeStart = $Range.Start
            RangeEnd = $Range.End
        }
    }

    $beforeOMaths = $document.OMaths.Count
    $start = $Range.Start
    $Range.Text = $linear
    $workingRange = $document.Range($start, $start + $linear.Length)

    try {
        $null = $document.OMaths.Add($workingRange)
        $document.OMaths.BuildUp()
    } catch {
        return [pscustomobject]@{
            Success = $false
            Message = $_.Exception.Message
            Text = $linear
            RangeStart = $start
            RangeEnd = $start + $linear.Length
        }
    }

    if ($document.OMaths.Count -gt $beforeOMaths) {
        $newOMath = $document.OMaths.Item($document.OMaths.Count)
        return [pscustomobject]@{
            Success = $true
            Message = 'Word OMath fallback succeeded.'
            Text = $linear
            RangeStart = $newOMath.Range.Start
            RangeEnd = $newOMath.Range.End
        }
    }

    return [pscustomobject]@{
        Success = $false
        Message = 'Word OMath fallback did not create a new equation.'
        Text = $linear
        RangeStart = $start
        RangeEnd = $start + $linear.Length
    }
}

function Try-ConvertCjkDisplayFormula {
    param(
        $Word,
        $Paragraph,
        [string]$Formula
    )

    $document = $Paragraph.Range.Document
    $segments = @(Get-CjkFormulaSegments -Formula $Formula)
    $start = $Paragraph.Range.Start
    $bodyRange = $document.Range($start, $Paragraph.Range.End - 1)
    $bodyRange.Text = ''
    $cursor = $document.Range($start, $start)
    $segmentLog = New-Object System.Collections.Generic.List[object]

    foreach ($segment in $segments) {
        if ($segment.Kind -eq 'text') {
            if (-not [string]::IsNullOrWhiteSpace($segment.Text)) {
                $cursor.Text = $segment.Text
                $segmentLog.Add([pscustomobject]@{
                    kind = 'text'
                    text = $segment.Text
                    engine = 'WordText'
                })
                $cursor.SetRange($cursor.End, $cursor.End)
            }
            continue
        }

        $segmentStart = $cursor.Start
        $wrapped = Wrap-ForMathType -Formula $segment.Text -Mode Inline
        $cursor.Text = $wrapped
        $segmentRange = $document.Range($segmentStart, $segmentStart + $wrapped.Length)

        $mathResult = Try-ConvertRangeWithMathType -Word $Word -Range $segmentRange -Formula $segment.Text -Mode Inline
        if ($mathResult.Success) {
            $segmentLog.Add([pscustomobject]@{
                kind = 'math'
                text = $segment.Text
                engine = 'MathType'
            })
            $cursor.SetRange($mathResult.RangeEnd, $mathResult.RangeEnd)
            continue
        }

        $segmentRange = $document.Range($segmentStart, $segmentStart + $wrapped.Length)
        $wordResult = Try-ConvertRangeWithWordEquation -Range $segmentRange -Formula $segment.Text
        if ($wordResult.Success) {
            $segmentLog.Add([pscustomobject]@{
                kind = 'math'
                text = $segment.Text
                engine = 'Word'
            })
            $cursor.SetRange($wordResult.RangeEnd, $wordResult.RangeEnd)
            continue
        }

        $restoreRange = $document.Range($start, $Paragraph.Range.End - 1)
        $restoreRange.Text = $Formula
        return [pscustomobject]@{
            Success = $false
            Message = 'CJK split conversion failed. MathType=' + $mathResult.Message + ' | Word=' + $wordResult.Message
            Segments = $segmentLog.ToArray()
        }
    }

    return [pscustomobject]@{
        Success = $true
        Message = 'CJK split conversion succeeded.'
        Segments = $segmentLog.ToArray()
    }
}

if (-not (Test-Path -LiteralPath $SourcePath)) {
    throw "Source DOCX not found: $SourcePath"
}

Assert-FileIsClosed -Path $TargetPath -Label 'Target document'

if (Test-Path -LiteralPath $TargetPath) {
    Remove-Item -LiteralPath $TargetPath -Force
}

if (Test-Path -LiteralPath $LogPath) {
    Remove-Item -LiteralPath $LogPath -Force
}

Copy-Item -LiteralPath $SourcePath -Destination $TargetPath -Force

$summary = [ordered]@{
    source = $SourcePath
    target = $TargetPath
    startedAt = (Get-Date).ToString('s')
    mathTypeAddIn = $false
    totalParagraphs = 0
    displayConverted = 0
    inlineConverted = 0
    mathTypeConverted = 0
    wordConverted = 0
    skipped = 0
    failed = 0
}

$events = New-Object System.Collections.Generic.List[object]
$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($TargetPath, $false, $false)

    foreach ($addin in $word.AddIns) {
        if ($addin.Name -eq 'MathType Commands 2016.dotm' -and $addin.Installed) {
            $summary.mathTypeAddIn = $true
        }
    }

    $summary.totalParagraphs = $document.Paragraphs.Count

    for ($index = $document.Paragraphs.Count; $index -ge 1; $index--) {
        $paragraph = $document.Paragraphs.Item($index)
        $body = Get-ParagraphBodyText -Paragraph $paragraph
        if ([string]::IsNullOrWhiteSpace($body)) {
            continue
        }

        if ($paragraph.Range.OMaths.Count -gt 0 -or $paragraph.Range.InlineShapes.Count -gt 0) {
            $summary.skipped++
            $events.Add((New-LogEvent -Kind 'paragraph' -ParagraphIndex $index -Engine 'none' -Status 'skipped-existing-object' -Formula (Normalize-FormulaText -Text $body)))
            continue
        }

        if (-not (Test-HasFormulaSignature -Text $body)) {
            continue
        }

        if (Test-IsDisplayFormulaParagraph -Text $body) {
            $formula = Normalize-FormulaText -Text $body
            if (Test-IsAsciiIdentifierLike -Text $formula) {
                $summary.skipped++
                $events.Add((New-LogEvent -Kind 'display' -ParagraphIndex $index -Engine 'none' -Status 'skipped-identifier' -Formula $formula))
                continue
            }

            if (Test-HasCjkTextMacro -Formula $formula) {
                $cjkResult = Try-ConvertCjkDisplayFormula -Word $word -Paragraph $paragraph -Formula $formula
                if ($cjkResult.Success) {
                    $paragraph.Alignment = $wdAlignParagraphCenter
                    $summary.displayConverted++
                    $summary.mathTypeConverted++
                    $events.Add((New-LogEvent -Kind 'display' -ParagraphIndex $index -Engine 'Mixed' -Status 'display-mathtype-cjk-split' -Formula $formula -Message $cjkResult.Message -Segments $cjkResult.Segments))
                } else {
                    $summary.failed++
                    $events.Add((New-LogEvent -Kind 'display' -ParagraphIndex $index -Engine 'none' -Status 'failed' -Formula $formula -Message $cjkResult.Message -Segments $cjkResult.Segments))
                }
                continue
            }

            $range = $document.Range($paragraph.Range.Start, $paragraph.Range.Start + $body.Length)
            $result = $null
            if ($summary.mathTypeAddIn) {
                $result = Try-ConvertRangeWithMathType -Word $word -Range $range -Formula $formula -Mode Display
            }

            if ($null -ne $result -and $result.Success) {
                $paragraph.Alignment = $wdAlignParagraphCenter
                $summary.displayConverted++
                $summary.mathTypeConverted++
                $events.Add((New-LogEvent -Kind 'display' -ParagraphIndex $index -Engine 'MathType' -Status 'display-mathtype' -Formula $formula -Message $result.Message))
                continue
            }

            $mathTypeMessage = if ($null -ne $result) { $result.Message } else { 'MathType not attempted.' }
            $currentEnd = if ($null -ne $result) { $result.RangeEnd } else { $paragraph.Range.Start + $formula.Length }
            $range = $document.Range($paragraph.Range.Start, $currentEnd)
            $fallback = Try-ConvertRangeWithWordEquation -Range $range -Formula $formula
            if ($fallback.Success) {
                $paragraph.Alignment = $wdAlignParagraphCenter
                $summary.displayConverted++
                $summary.wordConverted++
                $events.Add((New-LogEvent -Kind 'display' -ParagraphIndex $index -Engine 'Word' -Status 'word-fallback' -Formula $formula -Message $fallback.Message))
            } else {
                $restoreRange = $document.Range($paragraph.Range.Start, $paragraph.Range.End - 1)
                $restoreRange.Text = $formula
                $summary.failed++
                $events.Add((New-LogEvent -Kind 'display' -ParagraphIndex $index -Engine 'none' -Status 'failed' -Formula $formula -Message ($fallback.Message + ' | MathType=' + $mathTypeMessage)))
            }

            continue
        }

        $matches = @(Get-InlineFormulaMatches -Text $body)
        if ($matches.Count -eq 0) {
            continue
        }

        foreach ($match in $matches) {
            $formula = $match.Text.Trim()
            if (Test-ShouldSkipInlineMatch -Text $formula) {
                $summary.skipped++
                $events.Add((New-LogEvent -Kind 'inline' -ParagraphIndex $index -Engine 'none' -Status 'skipped-identifier' -Formula $formula))
                continue
            }

            $start = $paragraph.Range.Start + $match.Start
            $range = $document.Range($start, $start + $match.Length)
            $result = $null
            if ($summary.mathTypeAddIn) {
                $result = Try-ConvertRangeWithMathType -Word $word -Range $range -Formula $formula -Mode Inline
            }

            if ($null -ne $result -and $result.Success) {
                $summary.inlineConverted++
                $summary.mathTypeConverted++
                $events.Add((New-LogEvent -Kind 'inline' -ParagraphIndex $index -Engine 'MathType' -Status 'inline-mathtype' -Formula $formula -Message $result.Message))
                continue
            }

            $mathTypeMessage = if ($null -ne $result) { $result.Message } else { 'MathType not attempted.' }
            $currentEnd = if ($null -ne $result) { $result.RangeEnd } else { $start + $formula.Length }
            $range = $document.Range($start, $currentEnd)
            $fallback = Try-ConvertRangeWithWordEquation -Range $range -Formula $formula
            if ($fallback.Success) {
                $summary.inlineConverted++
                $summary.wordConverted++
                $events.Add((New-LogEvent -Kind 'inline' -ParagraphIndex $index -Engine 'Word' -Status 'word-fallback' -Formula $formula -Message $fallback.Message))
            } else {
                $range.Text = $formula
                $summary.failed++
                $events.Add((New-LogEvent -Kind 'inline' -ParagraphIndex $index -Engine 'none' -Status 'failed' -Formula $formula -Message ($fallback.Message + ' | MathType=' + $mathTypeMessage)))
            }
        }
    }

    $document.Save()
} finally {
    if ($document -ne $null) {
        $document.Close([ref]$false) | Out-Null
    }

    if ($word -ne $null) {
        $word.Quit() | Out-Null
    }
}

$newMathTypeProcesses = @(Get-Process MathType -ErrorAction SilentlyContinue | Where-Object { $initialMathTypeIds -notcontains $_.Id })
foreach ($process in $newMathTypeProcesses) {
    try {
        $process.CloseMainWindow() | Out-Null
        Start-Sleep -Milliseconds 200
        if (-not $process.HasExited) {
            $process | Stop-Process -Force
        }
    } catch {
    }
}

$summary.finishedAt = (Get-Date).ToString('s')
$output = [ordered]@{
    summary = $summary
    events = $events
}

$output | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $LogPath -Encoding UTF8
$summary
