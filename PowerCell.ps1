<#
.SYNOPSIS
    PowerCell - Streamlined WPF Excel Spreadsheet Editor with Authentic Sort & Filter
.DESCRIPTION
    A self-contained Windows WPF GUI spreadsheet application with an authentic Excel
    Sort & Filter dropdown menu, column header click sorting, live filter bar,
    multi-currency formatting, text positioning, and reactive formulas.
.EXAMPLE
    .\PowerCell.ps1
.EXAMPLE
    .\PowerCell.ps1 sample_data.csv
#>

param(
    [Parameter(Position=0, Mandatory=$false)]
    [string]$FilePath
)

# Load WPF Assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Data, System.Drawing

# ============================================================================
# SECTION 1: ENGINE & DATA MODEL
# ============================================================================

function Convert-ColIndexToName([int]$colIndex) {
    if ($colIndex -lt 1) { return "A" }
    $name = ""
    while ($colIndex -gt 0) {
        $rem = ($colIndex - 1) % 26
        $name = [char](65 + $rem) + $name
        $colIndex = [math]::Floor(($colIndex - 1) / 26)
    }
    return $name
}

function Convert-ColNameToIndex([string]$colName) {
    if ([string]::IsNullOrWhiteSpace($colName)) { return 1 }
    $colName = $colName.ToUpper()
    $result = 0
    for ($i = 0; $i -lt $colName.Length; $i++) {
        $char = $colName[$i]
        if ($char -ge 'A' -and $char -le 'Z') {
            $result = $result * 26 + ([int]$char - [int][char]'A' + 1)
        }
    }
    return [math]::Max(1, $result)
}

function Convert-CellNameToCoords([string]$cellRef) {
    if ([string]::IsNullOrWhiteSpace($cellRef)) { return @{ Col = 1; Row = 1 } }
    $cellRef = $cellRef.Trim().ToUpper()
    if ($cellRef -match '^([A-Z]+)([0-9]+)$') {
        $colName = $Matches[1]
        $rowNum = [int]$Matches[2]
        return @{
            Col = Convert-ColNameToIndex $colName
            Row = $rowNum
        }
    }
    return @{ Col = 1; Row = 1 }
}

function Convert-CoordsToCellName([int]$col, [int]$row) {
    $colName = Convert-ColIndexToName $col
    return "$colName$row"
}

function Parse-CellRange([string]$rangeStr) {
    $rangeStr = $rangeStr.Trim().ToUpper()
    if ($rangeStr.Contains(':')) {
        $parts = $rangeStr -split ':'
        $c1 = Convert-CellNameToCoords $parts[0]
        $c2 = Convert-CellNameToCoords $parts[1]
        return @{
            MinCol = [math]::Min($c1.Col, $c2.Col)
            MaxCol = [math]::Max($c1.Col, $c2.Col)
            MinRow = [math]::Min($c1.Row, $c2.Row)
            MaxRow = [math]::Max($c1.Row, $c2.Row)
        }
    } else {
        $c = Convert-CellNameToCoords $rangeStr
        return @{
            MinCol = $c.Col
            MaxCol = $c.Col
            MinRow = $c.Row
            MaxRow = $c.Row
        }
    }
}

class CellState {
    [string]$RawInput = ""
    [string]$Formula = ""
    [string]$EvaluatedValue = ""
    [bool]$IsFormula = $false
    [bool]$IsBold = $false
    [bool]$IsItalic = $false
    [string]$Align = "Left"
    [string]$VAlign = "Center"
    [string]$NumberFormat = "General"
}

class PowerCellEngine {
    [hashtable]$Cells = @{}
    [int]$MaxRow = 50
    [int]$MaxCol = 26
    [string]$FilePath = ""
    [bool]$IsModified = $false
    [bool]$IsZebraTable = $false

    PowerCellEngine() {
        $this.Cells = @{}
    }

    [string] GetCellKey([int]$col, [int]$row) {
        return "$col,$row"
    }

    [CellState] GetCell([int]$col, [int]$row) {
        $key = $this.GetCellKey($col, $row)
        if ($this.Cells.ContainsKey($key)) {
            return $this.Cells[$key]
        }
        return [CellState]::new()
    }

    [string] GetFormattedValue([int]$col, [int]$row) {
        $cell = $this.GetCell($col, $row)
        $valStr = if ($cell.IsFormula) { $cell.EvaluatedValue } else { $cell.RawInput }
        if ([string]::IsNullOrEmpty($valStr)) { return "" }

        [double]$num = 0
        if ([double]::TryParse($valStr, [ref]$num)) {
            switch -wildcard ($cell.NumberFormat) {
                "Currency_EUR" { return "€" + $num.ToString("N2") }
                "Currency_GBP" { return "£" + $num.ToString("N2") }
                "Currency_JPY" { return "¥" + $num.ToString("N0") }
                "Currency*"    { return "`$" + $num.ToString("N2") }
                "Percent"     { return ($num * 100).ToString("F1") + "%" }
            }
        }
        return $valStr
    }

    [string] GetDisplayValue([int]$col, [int]$row) {
        return $this.GetFormattedValue($col, $row)
    }

    [string] GetRawValue([int]$col, [int]$row) {
        $cell = $this.GetCell($col, $row)
        return $cell.RawInput
    }

    SetCell([int]$col, [int]$row, [string]$inputStr) {
        $key = $this.GetCellKey($col, $row)
        $existing = $this.GetCell($col, $row)

        if ([string]::IsNullOrEmpty($inputStr)) {
            if ($this.Cells.ContainsKey($key)) {
                $this.Cells.Remove($key)
                $this.IsModified = $true
            }
            $this.RecalculateAll()
            return
        }

        $cell = [CellState]::new()
        $cell.IsBold = $existing.IsBold
        $cell.IsItalic = $existing.IsItalic
        $cell.Align = $existing.Align
        $cell.VAlign = $existing.VAlign
        $cell.NumberFormat = $existing.NumberFormat
        $cell.RawInput = $inputStr

        if ($inputStr.StartsWith("=")) {
            $cell.IsFormula = $true
            $cell.Formula = $inputStr
        } else {
            $cell.IsFormula = $false
            $cell.EvaluatedValue = $inputStr
        }

        $this.Cells[$key] = $cell
        if ($col -gt $this.MaxCol) { $this.MaxCol = $col }
        if ($row -gt $this.MaxRow) { $this.MaxRow = $row }
        $this.IsModified = $true

        $this.RecalculateAll()
    }

    SetCellFormat([int]$col, [int]$row, [string]$propName, $propValue) {
        $cell = $this.GetCell($col, $row)
        switch ($propName) {
            "Bold" { $cell.IsBold = [bool]$propValue }
            "Italic" { $cell.IsItalic = [bool]$propValue }
            "Align" { $cell.Align = [string]$propValue }
            "VAlign" { $cell.VAlign = [string]$propValue }
            "NumberFormat" { $cell.NumberFormat = [string]$propValue }
        }
        $key = $this.GetCellKey($col, $row)
        $this.Cells[$key] = $cell
        $this.IsModified = $true
    }

    [double[]] GetNumericValuesFromRange([string]$rangeStr) {
        $range = Parse-CellRange $rangeStr
        $values = @()
        for ($r = $range.MinRow; $r -le $range.MaxRow; $r++) {
            for ($c = $range.MinCol; $c -le $range.MaxCol; $c++) {
                $disp = $this.GetDisplayValue($c, $r)
                [double]$val = 0
                if ([double]::TryParse($disp, [ref]$val)) {
                    $values += $val
                }
            }
        }
        return $values
    }

    [string] EvaluateFormula([string]$formulaStr) {
        if (-not $formulaStr.StartsWith("=")) { return $formulaStr }
        $expr = $formulaStr.Substring(1).Trim()

        if ($expr -match '^(SUM|AVG|AVERAGE|COUNT|MIN|MAX)\(([A-Z0-9:]+)\)$') {
            $fn = $Matches[1].ToUpper()
            $arg = $Matches[2]
            $vals = $this.GetNumericValuesFromRange($arg)

            if ($vals.Count -eq 0) {
                if ($fn -eq 'COUNT') { return "0" }
                return "0"
            }

            switch ($fn) {
                'SUM' {
                    $sum = ($vals | Measure-Object -Sum).Sum
                    return [string]$sum
                }
                { $_ -in 'AVG','AVERAGE' } {
                    $avg = ($vals | Measure-Object -Average).Average
                    return [string][math]::Round($avg, 4)
                }
                'COUNT' {
                    return [string]$vals.Count
                }
                'MIN' {
                    $min = ($vals | Measure-Object -Minimum).Minimum
                    return [string]$min
                }
                'MAX' {
                    $max = ($vals | Measure-Object -Maximum).Maximum
                    return [string]$max
                }
            }
        }

        $evalExpr = [regex]::Replace($expr, '(?<=\b)[A-Z]+[0-9]+\b', {
            param($match)
            $refStr = $match.Value
            $coords = Convert-CellNameToCoords $refStr
            $disp = $this.GetDisplayValue($coords.Col, $coords.Row)
            $cleanDisp = $disp.Replace('$', '').Replace('€', '').Replace('£', '').Replace('¥', '').Replace('%', '').Trim()
            [double]$num = 0
            if ([double]::TryParse($cleanDisp, [ref]$num)) {
                return $num.ToString([System.Globalization.CultureInfo]::InvariantCulture)
            } else {
                return "0"
            }
        })

        if ($evalExpr -match '^[0-9\.\+\-\*\/\%\(\)\s]+$') {
            try {
                $result = Invoke-Expression $evalExpr
                if ($result -is [ValueType]) {
                    return [string][math]::Round([double]$result, 4)
                }
                return [string]$result
            } catch {
                return "#ERR!"
            }
        }

        return "#VALUE!"
    }

    RecalculateAll() {
        for ($pass = 0; $pass -lt 2; $pass++) {
            foreach ($key in @($this.Cells.Keys)) {
                $cell = $this.Cells[$key]
                if ($cell.IsFormula) {
                    $cell.EvaluatedValue = $this.EvaluateFormula($cell.Formula)
                }
            }
        }
    }

    SortByColumn([int]$targetCol, [bool]$ascending = $true) {
        $rowObjs = @()
        for ($r = 2; $r -le $this.MaxRow; $r++) {
            $sortVal = $this.GetDisplayValue($targetCol, $r)
            [double]$numVal = 0
            $isNum = [double]::TryParse($sortVal, [ref]$numVal)
            $rowObjs += [PSCustomObject]@{
                RowIndex = $r
                SortVal  = if ($isNum) { $numVal } else { $sortVal }
            }
        }

        if ($ascending) {
            $sortedRows = $rowObjs | Sort-Object SortVal
        } else {
            $sortedRows = $rowObjs | Sort-Object SortVal -Descending
        }

        $newCells = @{}
        foreach ($key in $this.Cells.Keys) {
            $parts = $key -split ','
            $c = [int]$parts[0]
            $r = [int]$parts[1]
            if ($r -eq 1) {
                $newCells[$key] = $this.Cells[$key]
            }
        }

        $destRow = 2
        foreach ($item in $sortedRows) {
            $srcRow = $item.RowIndex
            for ($c = 1; $c -le $this.MaxCol; $c++) {
                $srcKey = "$c,$srcRow"
                if ($this.Cells.ContainsKey($srcKey)) {
                    $destKey = "$c,$destRow"
                    $newCells[$destKey] = $this.Cells[$srcKey]
                }
            }
            $destRow++
        }

        $this.Cells = $newCells
        $this.IsModified = $true
        $this.RecalculateAll()
    }

    InsertRow([int]$targetRow) {
        $newCells = @{}
        foreach ($key in $this.Cells.Keys) {
            $parts = $key -split ','
            $c = [int]$parts[0]
            $r = [int]$parts[1]

            if ($r -ge $targetRow) {
                $newKey = "$c,$($r + 1)"
                $newCells[$newKey] = $this.Cells[$key]
            } else {
                $newCells[$key] = $this.Cells[$key]
            }
        }
        $this.Cells = $newCells
        $this.MaxRow++
        $this.IsModified = $true
        $this.RecalculateAll()
    }

    DeleteRow([int]$targetRow) {
        $newCells = @{}
        foreach ($key in $this.Cells.Keys) {
            $parts = $key -split ','
            $c = [int]$parts[0]
            $r = [int]$parts[1]

            if ($r -eq $targetRow) {
                continue
            } elseif ($r -gt $targetRow) {
                $newKey = "$c,$($r - 1)"
                $newCells[$newKey] = $this.Cells[$key]
            } else {
                $newCells[$key] = $this.Cells[$key]
            }
        }
        $this.Cells = $newCells
        if ($this.MaxRow -gt 1) { $this.MaxRow-- }
        $this.IsModified = $true
        $this.RecalculateAll()
    }
}

# ============================================================================
# SECTION 2: FORMATS
# ============================================================================

function Import-PowerCellFile([PowerCellEngine]$engine, [string]$filePath) {
    if (-not (Test-Path -Path $filePath)) {
        throw "File not found: $filePath"
    }

    $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
    $engine.FilePath = [System.IO.Path]::GetFullPath($filePath)

    if ($ext -in '.csv', '.tsv', '.txt') {
        $sep = if ($ext -eq '.tsv') { "`t" } else { "," }
        Import-PowerCellCsv -engine $engine -filePath $filePath -delimiter $sep
    } elseif ($ext -in '.xlsx', '.xls') {
        Import-PowerCellExcel -engine $engine -filePath $filePath
    } else {
        Import-PowerCellCsv -engine $engine -filePath $filePath -delimiter ","
    }

    $engine.IsModified = $false
}

function Export-PowerCellFile([PowerCellEngine]$engine, [string]$filePath) {
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        $filePath = $engine.FilePath
    }
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        throw "No file path specified for saving."
    }

    $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
    $engine.FilePath = [System.IO.Path]::GetFullPath($filePath)

    if ($ext -in '.csv', '.tsv', '.txt') {
        $sep = if ($ext -eq '.tsv') { "`t" } else { "," }
        Export-PowerCellCsv -engine $engine -filePath $filePath -delimiter $sep
    } elseif ($ext -in '.xlsx', '.xls') {
        Export-PowerCellExcel -engine $engine -filePath $filePath
    } else {
        Export-PowerCellCsv -engine $engine -filePath $filePath -delimiter ","
    }

    $engine.IsModified = $false
}

function Import-PowerCellCsv([PowerCellEngine]$engine, [string]$filePath, [string]$delimiter) {
    $lines = Get-Content -Path $filePath -Encoding UTF8
    $row = 1
    $pattern = [regex]::Escape($delimiter) + '(?=(?:[^"]*"[^"]*")*[^"]*$)'
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { $row++; continue }

        $cols = [regex]::Split($line, $pattern)
        $col = 1
        foreach ($cVal in $cols) {
            $trimmed = $cVal.Trim()
            if ($trimmed.StartsWith('"') -and $trimmed.EndsWith('"') -and $trimmed.Length -ge 2) {
                $trimmed = $trimmed.Substring(1, $trimmed.Length - 2).Replace('""', '"')
            }
            if (-not [string]::IsNullOrEmpty($trimmed)) {
                $engine.SetCell($col, $row, $trimmed)
            }
            $col++
        }
        $row++
    }
}

function Export-PowerCellCsv([PowerCellEngine]$engine, [string]$filePath, [string]$delimiter) {
    $maxR = $engine.MaxRow
    $maxC = $engine.MaxCol

    for ($r = $maxR; $r -ge 1; $r--) {
        $hasData = $false
        for ($c = 1; $c -le $maxC; $c++) {
            if (-not [string]::IsNullOrEmpty($engine.GetRawValue($c, $r))) { $hasData = $true; break }
        }
        if ($hasData) { $maxR = $r; break }
    }

    for ($c = $maxC; $c -ge 1; $c--) {
        $hasData = $false
        for ($r = 1; $r -le $maxR; $r++) {
            if (-not [string]::IsNullOrEmpty($engine.GetRawValue($c, $r))) { $hasData = $true; break }
        }
        if ($hasData) { $maxC = $c; break }
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    for ($r = 1; $r -le $maxR; $r++) {
        $rowVals = [System.Collections.Generic.List[string]]::new()
        for ($c = 1; $c -le $maxC; $c++) {
            $val = $engine.GetRawValue($c, $r)
            if ($val.Contains($delimiter) -or $val.Contains('"') -or $val.Contains("`n")) {
                $escaped = $val.Replace('"', '""')
                $val = '"' + $escaped + '"'
            }
            $rowVals.Add($val)
        }
        $lines.Add(($rowVals -join $delimiter))
    }

    [System.IO.File]::WriteAllLines($filePath, $lines, [System.Text.Encoding]::UTF8)
}

function Import-PowerCellExcel([PowerCellEngine]$engine, [string]$filePath) {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($filePath)
        $sheet = $workbook.Sheets.Item(1)
        $usedRange = $sheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        for ($r = 1; $r -le $rowCount; $r++) {
            for ($c = 1; $c -le $colCount; $c++) {
                $cell = $usedRange.Cells.Item($r, $c)
                $formula = $cell.Formula
                $val = $cell.Value2
                if ($formula -and $formula.StartsWith("=")) {
                    $engine.SetCell($c, $r, [string]$formula)
                } elseif ($null -ne $val) {
                    $engine.SetCell($c, $r, [string]$val)
                }
            }
        }

        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        throw "Failed to open Excel file via COM: $_. Save as CSV."
    }
}

function Export-PowerCellExcel([PowerCellEngine]$engine, [string]$filePath) {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Sheets.Item(1)

        for ($r = 1; $r -le $engine.MaxRow; $r++) {
            for ($c = 1; $c -le $engine.MaxCol; $c++) {
                $raw = $engine.GetRawValue($c, $r)
                if (-not [string]::IsNullOrEmpty($raw)) {
                    $sheet.Cells.Item($r, $c).Value2 = $raw
                }
            }
        }

        $workbook.SaveAs($filePath)
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        $csvPath = [System.IO.Path]::ChangeExtension($filePath, ".csv")
        Export-PowerCellCsv -engine $engine -filePath $csvPath -delimiter ","
        throw "Excel (COM) is not installed on this system.`nSpreadsheet has been exported as CSV to:`n$csvPath"
    }
}

# ============================================================================
# SECTION 3: STREAMLINED & VIRTUALIZED WPF INTERFACE
# ============================================================================

$engine = [PowerCellEngine]::new()

if (-not [string]::IsNullOrWhiteSpace($FilePath)) {
    if (Test-Path $FilePath) {
        Import-PowerCellFile -engine $engine -filePath $FilePath
    } else {
        $engine.FilePath = [System.IO.Path]::GetFullPath($FilePath)
    }
}

# Authentic Office Ribbon Layout
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Excel - PowerCell Spreadsheet" Height="800" Width="1360"
        WindowStartupLocation="CenterScreen" Background="#262626" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="Button" x:Key="RibbonTabBtn">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Foreground" Value="#E1E1E1"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#333333"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="Button" x:Key="OfficeToolBtn">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Margin" Value="1"/>
            <Setter Property="Foreground" Value="#F3F3F3"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#383838"/>
                    <Setter Property="BorderBrush" Value="#555555"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ToggleButton" x:Key="OfficeToggleBtn">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Margin" Value="1"/>
            <Setter Property="Foreground" Value="#F3F3F3"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsChecked" Value="True">
                    <Setter Property="Background" Value="#107C41"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="BorderBrush" Value="#555555"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Title Bar -->
            <RowDefinition Height="Auto"/> <!-- Main Excel Ribbon Bar -->
            <RowDefinition Height="Auto"/> <!-- Formula Bar -->
            <RowDefinition Height="*"/>    <!-- Data Grid -->
            <RowDefinition Height="Auto"/> <!-- Status Bar -->
        </Grid.RowDefinitions>

        <!-- 1. Office Green Title Bar -->
        <Border Grid.Row="0" Background="#107C41" Padding="10,6">
            <DockPanel LastChildFill="True">
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Left" VerticalAlignment="Center">
                    <TextBlock Text="📗" FontSize="15" Margin="0,0,8,0" VerticalAlignment="Center"/>
                    <TextBlock Text="Excel - PowerCell" Foreground="#FFFFFF" FontWeight="Bold" FontSize="14" Margin="0,0,15,0"/>
                    <TextBlock Name="txtFileName" Text="Untitled.csv" Foreground="#DFF6DD" FontSize="12" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" HorizontalAlignment="Right">
                    <Button Name="btnQuickSave" Content="💾 Save" Style="{StaticResource RibbonTabBtn}"/>
                    <Button Name="btnQuickOpen" Content="📂 Open" Style="{StaticResource RibbonTabBtn}"/>
                    <Button Name="btnNewSheet" Content="📄 New" Style="{StaticResource RibbonTabBtn}"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- 2. Streamlined Office Ribbon Toolbar -->
        <Border Grid.Row="1" Background="#1F1F1F" BorderBrush="#333333" BorderThickness="0,0,0,1" Padding="4,6">
            <StackPanel Orientation="Horizontal" Height="78">
                
                <!-- Group 1: Clipboard -->
                <Border BorderBrush="#333333" BorderThickness="0,0,1,0" Padding="6,0,8,0" Margin="0,0,4,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <StackPanel Grid.Row="0" Orientation="Horizontal" VerticalAlignment="Center">
                            <Button Name="btnSave" Content="💾 Save" Style="{StaticResource OfficeToolBtn}" Padding="8,10"/>
                            <StackPanel Orientation="Vertical" VerticalAlignment="Center" Margin="4,0">
                                <Button Name="btnOpen" Content="📂 Open" Style="{StaticResource OfficeToolBtn}"/>
                                <Button Name="btnExportExcel" Content="📊 Export XLSX" Style="{StaticResource OfficeToolBtn}"/>
                            </StackPanel>
                        </StackPanel>
                        <TextBlock Grid.Row="1" Text="Clipboard" Foreground="#A1A1A1" FontSize="11" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                    </Grid>
                </Border>

                <!-- Group 2: Font -->
                <Border BorderBrush="#333333" BorderThickness="0,0,1,0" Padding="6,0,8,0" Margin="0,0,4,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <StackPanel Grid.Row="0" VerticalAlignment="Center">
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                                <ComboBox Width="100" SelectedIndex="0" Margin="0,0,4,0">
                                    <ComboBoxItem Content="Aptos Narrow"/>
                                    <ComboBoxItem Content="Segoe UI"/>
                                    <ComboBoxItem Content="Calibri"/>
                                    <ComboBoxItem Content="Arial"/>
                                </ComboBox>
                                <ComboBox Width="45" SelectedIndex="1">
                                    <ComboBoxItem Content="10"/>
                                    <ComboBoxItem Content="11"/>
                                    <ComboBoxItem Content="12"/>
                                    <ComboBoxItem Content="14"/>
                                </ComboBox>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal">
                                <ToggleButton Name="btnBold" Content="B" Style="{StaticResource OfficeToggleBtn}" ToolTip="Bold"/>
                                <ToggleButton Name="btnItalic" Content="I" Style="{StaticResource OfficeToggleBtn}" FontStyle="Italic" ToolTip="Italic"/>
                                <Button Content="U" Style="{StaticResource OfficeToolBtn}" FontWeight="Bold"/>
                                <Button Name="btnToggleZebra" Content="🎨 Table Style" Style="{StaticResource OfficeToolBtn}"/>
                            </StackPanel>
                        </StackPanel>
                        <TextBlock Grid.Row="1" Text="Font" Foreground="#A1A1A1" FontSize="11" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                    </Grid>
                </Border>

                <!-- Group 3: Alignment -->
                <Border BorderBrush="#333333" BorderThickness="0,0,1,0" Padding="6,0,8,0" Margin="0,0,4,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <StackPanel Grid.Row="0" VerticalAlignment="Center">
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                                <Button Name="btnAlignTop" Content="≡ Top" Style="{StaticResource OfficeToolBtn}"/>
                                <Button Name="btnAlignMid" Content="= Mid" Style="{StaticResource OfficeToolBtn}"/>
                                <Button Name="btnAlignBot" Content="_ Bot" Style="{StaticResource OfficeToolBtn}"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal">
                                <Button Name="btnAlignLeft" Content="⯇ Left" Style="{StaticResource OfficeToolBtn}"/>
                                <Button Name="btnAlignCenter" Content="⯌ Center" Style="{StaticResource OfficeToolBtn}"/>
                                <Button Name="btnAlignRight" Content="Right ⯈" Style="{StaticResource OfficeToolBtn}"/>
                            </StackPanel>
                        </StackPanel>
                        <TextBlock Grid.Row="1" Text="Alignment" Foreground="#A1A1A1" FontSize="11" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                    </Grid>
                </Border>

                <!-- Group 4: Number & Multi-Currencies -->
                <Border BorderBrush="#333333" BorderThickness="0,0,1,0" Padding="6,0,8,0" Margin="0,0,4,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <StackPanel Grid.Row="0" VerticalAlignment="Center">
                            <ComboBox Name="cmbNumFormat" Width="135" Margin="0,0,0,4" SelectedIndex="0">
                                <ComboBoxItem Content="General"/>
                                <ComboBoxItem Content="Currency ($ USD)"/>
                                <ComboBoxItem Content="Currency (€ EUR)"/>
                                <ComboBoxItem Content="Currency (£ GBP)"/>
                                <ComboBoxItem Content="Currency (¥ JPY)"/>
                                <ComboBoxItem Content="Percent (%)"/>
                            </ComboBox>
                            <StackPanel Orientation="Horizontal">
                                <Button Name="btnCurrUSD" Content="$" Style="{StaticResource OfficeToolBtn}" ToolTip="USD ($)"/>
                                <Button Name="btnCurrEUR" Content="€" Style="{StaticResource OfficeToolBtn}" ToolTip="EUR (€)"/>
                                <Button Name="btnCurrGBP" Content="£" Style="{StaticResource OfficeToolBtn}" ToolTip="GBP (£)"/>
                                <Button Name="btnCurrJPY" Content="¥" Style="{StaticResource OfficeToolBtn}" ToolTip="JPY (¥)"/>
                                <Button Name="btnPercent" Content="%" Style="{StaticResource OfficeToolBtn}" ToolTip="Percent"/>
                            </StackPanel>
                        </StackPanel>
                        <TextBlock Grid.Row="1" Text="Number &amp; Currencies" Foreground="#A1A1A1" FontSize="11" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                    </Grid>
                </Border>

                <!-- Group 5: Cells -->
                <Border BorderBrush="#333333" BorderThickness="0,0,1,0" Padding="6,0,8,0" Margin="0,0,4,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <StackPanel Grid.Row="0" Orientation="Horizontal" VerticalAlignment="Center">
                            <Button Name="btnInsertRow" Content="➕ Insert Row" Style="{StaticResource OfficeToolBtn}"/>
                            <Button Name="btnDeleteRow" Content="➖ Delete Row" Style="{StaticResource OfficeToolBtn}"/>
                        </StackPanel>
                        <TextBlock Grid.Row="1" Text="Cells" Foreground="#A1A1A1" FontSize="11" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                    </Grid>
                </Border>

                <!-- Group 6: Editing & Authentic Sort/Filter -->
                <Border BorderBrush="#333333" BorderThickness="0,0,1,0" Padding="6,0,8,0" Margin="0,0,4,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <StackPanel Grid.Row="0" VerticalAlignment="Center">
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                                <Button Name="btnSum" Content="∑ AutoSum" Style="{StaticResource OfficeToolBtn}"/>
                                <Button Name="btnAvg" Content="x̄ Average" Style="{StaticResource OfficeToolBtn}"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal">
                                <Button Name="btnSortFilterMenu" Content="AZ 🌁 Sort &amp; Filter ▾" Style="{StaticResource OfficeToolBtn}">
                                    <Button.ContextMenu>
                                        <ContextMenu Background="#252525" Foreground="#FFFFFF" BorderBrush="#555555">
                                            <MenuItem Name="menuSortAsc" Header="Sort A to Z" Foreground="#FFFFFF"/>
                                            <MenuItem Name="menuSortDesc" Header="Sort Z to A" Foreground="#FFFFFF"/>
                                            <Separator Background="#444444"/>
                                            <MenuItem Name="menuClearFilter" Header="Clear Filter" Foreground="#FFFFFF"/>
                                        </ContextMenu>
                                    </Button.ContextMenu>
                                </Button>
                                <TextBox Name="txtFilter" Width="65" Padding="2,1" Margin="2,0" VerticalAlignment="Center" ToolTip="Search filter rows"/>
                                <Button Name="btnClearFilter" Content="✖" Style="{StaticResource OfficeToolBtn}" ToolTip="Clear search filter"/>
                            </StackPanel>
                        </StackPanel>
                        <TextBlock Grid.Row="1" Text="Editing" Foreground="#A1A1A1" FontSize="11" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                    </Grid>
                </Border>

            </StackPanel>
        </Border>

        <!-- 3. Formula Bar -->
        <Border Grid.Row="2" Background="#1F1F1F" BorderBrush="#333333" BorderThickness="0,0,0,1" Padding="6,4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="110"/> <!-- Name Box -->
                    <ColumnDefinition Width="Auto"/> <!-- fx icon -->
                    <ColumnDefinition Width="*"/>   <!-- Formula Input -->
                </Grid.ColumnDefinitions>
                
                <TextBox Name="txtNameBox" Grid.Column="0" Text="A1" IsReadOnly="True" 
                         TextAlignment="Center" FontWeight="Bold" Background="#2D2D2D" Foreground="#FFFFFF" BorderBrush="#555555" Padding="2,3"/>
                
                <TextBlock Grid.Column="1" Text=" fx " FontWeight="Bold" FontStyle="Italic" Foreground="#107C41" 
                           FontSize="14" VerticalAlignment="Center" Margin="8,0"/>
                
                <TextBox Name="txtFormula" Grid.Column="2" Padding="4,3" Background="#2D2D2D" Foreground="#FFFFFF" BorderBrush="#555555" FontSize="13"/>
            </Grid>
        </Border>

        <!-- 4. High-Performance Virtualized Data Grid -->
        <DataGrid Name="gridSpreadsheet" Grid.Row="3" AutoGenerateColumns="False" 
                  CanUserAddRows="False" CanUserDeleteRows="False" CanUserSortColumns="True" GridLinesVisibility="All"
                  HorizontalGridLinesBrush="#333333" VerticalGridLinesBrush="#333333"
                  HeadersVisibility="All" RowHeaderWidth="50" RowHeight="24"
                  Background="#121212" Foreground="#FFFFFF" SelectionMode="Extended" SelectionUnit="Cell"
                  EnableRowVirtualization="True" EnableColumnVirtualization="True"
                  VirtualizingStackPanel.IsVirtualizing="True" VirtualizingStackPanel.VirtualizationMode="Recycling"
                  FontSize="12">
            <DataGrid.Resources>
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background" Value="#252525"/>
                    <Setter Property="Foreground" Value="#E1E1E1"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="HorizontalContentAlignment" Value="Center"/>
                    <Setter Property="BorderBrush" Value="#3B3B3B"/>
                    <Setter Property="BorderThickness" Value="0,0,1,1"/>
                    <Setter Property="Padding" Value="4"/>
                    <Setter Property="Cursor" Value="Hand"/>
                </Style>
                <Style TargetType="DataGridRowHeader">
                    <Setter Property="Background" Value="#252525"/>
                    <Setter Property="Foreground" Value="#E1E1E1"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="HorizontalContentAlignment" Value="Center"/>
                    <Setter Property="BorderBrush" Value="#3B3B3B"/>
                    <Setter Property="BorderThickness" Value="0,0,1,1"/>
                </Style>
                <Style TargetType="DataGridCell">
                    <Setter Property="Background" Value="#181818"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                    <Style.Triggers>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter Property="Background" Value="#2D3748"/>
                            <Setter Property="BorderBrush" Value="#107C41"/>
                            <Setter Property="BorderThickness" Value="2"/>
                            <Setter Property="Foreground" Value="#FFFFFF"/>
                        </Trigger>
                    </Style.Triggers>
                </Style>
            </DataGrid.Resources>
        </DataGrid>

        <!-- 5. Status Bar -->
        <Border Grid.Row="4" Background="#107C41" Padding="10,4">
            <DockPanel LastChildFill="True">
                <TextBlock Name="txtStatus" Text="Ready" Foreground="#FFFFFF" FontSize="12" DockPanel.Dock="Left"/>
                <TextBlock Name="txtCalcSummary" Text="Sum: 0  |  Average: 0  |  Count: 0" Foreground="#DFF6DD" FontSize="12" DockPanel.Dock="Right" HorizontalAlignment="Right"/>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@

# Read XAML
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Get UI Control References
$gridSpreadsheet    = $window.FindName("gridSpreadsheet")
$txtNameBox         = $window.FindName("txtNameBox")
$txtFormula         = $window.FindName("txtFormula")
$txtFileName        = $window.FindName("txtFileName")
$txtStatus          = $window.FindName("txtStatus")
$txtCalcSummary     = $window.FindName("txtCalcSummary")

$btnSave            = $window.FindName("btnSave")
$btnOpen            = $window.FindName("btnOpen")
$btnExportExcel     = $window.FindName("btnExportExcel")
$btnBold            = $window.FindName("btnBold")
$btnItalic          = $window.FindName("btnItalic")

$btnAlignTop        = $window.FindName("btnAlignTop")
$btnAlignMid        = $window.FindName("btnAlignMid")
$btnAlignBot        = $window.FindName("btnAlignBot")
$btnAlignLeft       = $window.FindName("btnAlignLeft")
$btnAlignCenter     = $window.FindName("btnAlignCenter")
$btnAlignRight      = $window.FindName("btnAlignRight")

$cmbNumFormat       = $window.FindName("cmbNumFormat")

$btnCurrUSD         = $window.FindName("btnCurrUSD")
$btnCurrEUR         = $window.FindName("btnCurrEUR")
$btnCurrGBP         = $window.FindName("btnCurrGBP")
$btnCurrJPY         = $window.FindName("btnCurrJPY")
$btnPercent         = $window.FindName("btnPercent")

$btnSortFilterMenu  = $window.FindName("btnSortFilterMenu")
$menuSortAsc        = $window.FindName("menuSortAsc")
$menuSortDesc       = $window.FindName("menuSortDesc")
$menuClearFilter    = $window.FindName("menuClearFilter")

$txtFilter          = $window.FindName("txtFilter")
$btnClearFilter     = $window.FindName("btnClearFilter")
$btnToggleZebra     = $window.FindName("btnToggleZebra")

$btnInsertRow       = $window.FindName("btnInsertRow")
$btnDeleteRow       = $window.FindName("btnDeleteRow")
$btnSum             = $window.FindName("btnSum")
$btnAvg             = $window.FindName("btnAvg")
$btnQuickSave       = $window.FindName("btnQuickSave")
$btnQuickOpen       = $window.FindName("btnQuickOpen")
$btnNewSheet        = $window.FindName("btnNewSheet")

# DataTable & Columns Setup
$table = [System.Data.DataTable]::new("Spreadsheet")
$maxC = [math]::Max(20, $engine.MaxCol)
$maxR = [math]::Max(50, $engine.MaxRow)

for ($c = 1; $c -le $maxC; $c++) {
    $colName = Convert-ColIndexToName $c
    [void]$table.Columns.Add($colName, [string])

    $dgCol = [System.Windows.Controls.DataGridTextColumn]::new()
    $dgCol.Header = $colName
    $dgCol.Binding = [System.Windows.Data.Binding]::new($colName)
    $dgCol.Width = 100
    $gridSpreadsheet.Columns.Add($dgCol)
}

for ($r = 1; $r -le $maxR; $r++) {
    [void]$table.Rows.Add($table.NewRow())
}

$gridSpreadsheet.ItemsSource = $table.DefaultView

# Fast Data Refresh Function
function Refresh-GridUI {
    $filterText = $txtFilter.Text.Trim().ToLower()

    for ($r = 1; $r -le $maxR; $r++) {
        $row = $table.Rows[$r - 1]
        for ($c = 1; $c -le $maxC; $c++) {
            $colName = Convert-ColIndexToName $c
            $val = $engine.GetDisplayValue($c, $r)
            if ($row[$colName] -ne $val) {
                $row[$colName] = $val
            }
        }
    }

    if ($engine.FilePath) {
        $txtFileName.Text = [System.IO.Path]::GetFileName($engine.FilePath)
    } else {
        $txtFileName.Text = "Untitled.csv"
    }
}

# Attach Row Header Numbers (1..N) & Apply Zebra Styling
$gridSpreadsheet.add_LoadingRow({
    param($sender, $e)
    $e.Row.Header = ($e.Row.GetIndex() + 1).ToString()
    if ($engine.IsZebraTable -and ($e.Row.GetIndex() % 2 -eq 1)) {
        $e.Row.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#202020")
    } else {
        $e.Row.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#181818")
    }
})

# Selected Cells Changed Handler
$gridSpreadsheet.add_SelectedCellsChanged({
    $selectedCount = $gridSpreadsheet.SelectedCells.Count
    if ($selectedCount -gt 0) {
        if ($selectedCount -eq 1) {
            $cellInfo = $gridSpreadsheet.SelectedCells[0]
            if ($null -ne $cellInfo -and $null -ne $cellInfo.Item -and $null -ne $cellInfo.Column) {
                $rowIndex = $gridSpreadsheet.Items.IndexOf($cellInfo.Item) + 1
                $colIndex = $gridSpreadsheet.Columns.IndexOf($cellInfo.Column) + 1

                if ($rowIndex -gt 0 -and $colIndex -gt 0) {
                    $cellRef = Convert-CoordsToCellName $colIndex $rowIndex
                    $txtNameBox.Text = $cellRef
                    $cell = $engine.GetCell($colIndex, $rowIndex)
                    $txtFormula.Text = $cell.RawInput

                    $btnBold.IsChecked = $cell.IsBold
                    $btnItalic.IsChecked = $cell.IsItalic
                    switch ($cell.NumberFormat) {
                        "Currency_USD" { $cmbNumFormat.SelectedIndex = 1 }
                        "Currency_EUR" { $cmbNumFormat.SelectedIndex = 2 }
                        "Currency_GBP" { $cmbNumFormat.SelectedIndex = 3 }
                        "Currency_JPY" { $cmbNumFormat.SelectedIndex = 4 }
                        "Percent"      { $cmbNumFormat.SelectedIndex = 5 }
                        default        { $cmbNumFormat.SelectedIndex = 0 }
                    }

                    $dispVal = $engine.GetDisplayValue($colIndex, $rowIndex)
                    if (-not [string]::IsNullOrEmpty($dispVal)) {
                        $cleanVal = $dispVal.Replace('$', '').Replace('€', '').Replace('£', '').Replace('¥', '').Replace('%', '').Trim()
                        [double]$num = 0
                        if ([double]::TryParse($cleanVal, [ref]$num)) {
                            $txtCalcSummary.Text = "Sum: $num  |  Average: $num  |  Count: 1"
                        } else {
                            $txtCalcSummary.Text = "Ready"
                        }
                    } else {
                        $txtCalcSummary.Text = "Ready"
                    }
                }
            }
        } else {
            [double]$sum = 0
            [int]$numCount = 0
            $minR = 999999; $maxR = 0; $minC = 999999; $maxC = 0

            foreach ($cellInfo in $gridSpreadsheet.SelectedCells) {
                if ($null -ne $cellInfo -and $null -ne $cellInfo.Item -and $null -ne $cellInfo.Column) {
                    $r = $gridSpreadsheet.Items.IndexOf($cellInfo.Item) + 1
                    $c = $gridSpreadsheet.Columns.IndexOf($cellInfo.Column) + 1
                    if ($r -gt 0 -and $c -gt 0) {
                        if ($r -lt $minR) { $minR = $r }; if ($r -gt $maxR) { $maxR = $r }
                        if ($c -lt $minC) { $minC = $c }; if ($c -gt $maxC) { $maxC = $c }

                        $dispVal = $engine.GetDisplayValue($c, $r)
                        if (-not [string]::IsNullOrEmpty($dispVal)) {
                            $cleanVal = $dispVal.Replace('$', '').Replace('€', '').Replace('£', '').Replace('¥', '').Replace('%', '').Trim()
                            [double]$num = 0
                            if ([double]::TryParse($cleanVal, [ref]$num)) {
                                $sum += $num
                                $numCount++
                            }
                        }
                    }
                }
            }

            if ($minR -le $maxR -and $minC -le $maxC) {
                $txtNameBox.Text = "$(Convert-CoordsToCellName $minC $minR):$(Convert-CoordsToCellName $maxC $maxR)"
            } else {
                $txtNameBox.Text = "$selectedCount Cells"
            }
            $txtFormula.Text = ""

            if ($numCount -gt 0) {
                $avg = [math]::Round(($sum / $numCount), 4)
                $txtCalcSummary.Text = "Sum: $sum  |  Average: $avg  |  Count: $numCount"
            } else {
                $txtCalcSummary.Text = "Selected: $selectedCount Cells"
            }
        }
    }
})

# Apply formatting & text alignment to all selected cells
function Apply-FormatToSelectedCells([string]$propName, $propValue) {
    if ($gridSpreadsheet.SelectedCells.Count -gt 0) {
        foreach ($cellInfo in $gridSpreadsheet.SelectedCells) {
            if ($null -ne $cellInfo -and $null -ne $cellInfo.Item -and $null -ne $cellInfo.Column) {
                $r = $gridSpreadsheet.Items.IndexOf($cellInfo.Item) + 1
                $c = $gridSpreadsheet.Columns.IndexOf($cellInfo.Column) + 1
                if ($r -gt 0 -and $c -gt 0) {
                    $engine.SetCellFormat($c, $r, $propName, $propValue)
                }
            }
        }
        
        if ($propName -eq "Align") {
            $colIndex = $gridSpreadsheet.Columns.IndexOf($gridSpreadsheet.SelectedCells[0].Column)
            if ($colIndex -ge 0) {
                $col = $gridSpreadsheet.Columns[$colIndex] -as [System.Windows.Controls.DataGridTextColumn]
                if ($null -ne $col) {
                    $style = [System.Windows.Style]::new([System.Windows.Controls.TextBlock])
                    $alignVal = switch ($propValue) {
                        "Center" { [System.Windows.TextAlignment]::Center }
                        "Right"  { [System.Windows.TextAlignment]::Right }
                        default  { [System.Windows.TextAlignment]::Left }
                    }
                    $style.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.TextBlock]::TextAlignmentProperty, $alignVal))
                    $col.ElementStyle = $style
                }
            }
        }
        
        Refresh-GridUI
    }
}

# Formula Bar Commit
$txtFormula.add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
        if ($txtNameBox.Text) {
            $coords = Convert-CellNameToCoords $txtNameBox.Text
            $engine.SetCell($coords.Col, $coords.Row, $txtFormula.Text)
            Refresh-GridUI
            $txtStatus.Text = "Cell $($txtNameBox.Text) updated."
        }
    }
})

# Direct Cell Edit
$gridSpreadsheet.add_CellEditEnding({
    param($sender, $e)
    $rowIndex = $e.Row.GetIndex() + 1
    $colIndex = $e.Column.DisplayIndex + 1
    $editingElement = $e.EditingElement -as [System.Windows.Controls.TextBox]
    if ($null -ne $editingElement) {
        $newValue = $editingElement.Text
        $engine.SetCell($colIndex, $rowIndex, $newValue)
        $window.Dispatcher.InvokeAsync([Action]{ Refresh-GridUI })
    }
})

# Font & Text Formatting Handlers
$btnBold.add_Click({ Apply-FormatToSelectedCells "Bold" $btnBold.IsChecked })
$btnItalic.add_Click({ Apply-FormatToSelectedCells "Italic" $btnItalic.IsChecked })

# Text Alignment Positioning Handlers
$btnAlignTop.add_Click({ Apply-FormatToSelectedCells "VAlign" "Top" })
$btnAlignMid.add_Click({ Apply-FormatToSelectedCells "VAlign" "Center" })
$btnAlignBot.add_Click({ Apply-FormatToSelectedCells "VAlign" "Bottom" })
$btnAlignLeft.add_Click({ Apply-FormatToSelectedCells "Align" "Left" })
$btnAlignCenter.add_Click({ Apply-FormatToSelectedCells "Align" "Center" })
$btnAlignRight.add_Click({ Apply-FormatToSelectedCells "Align" "Right" })

# Number Format Combo Handler
$cmbNumFormat.add_SelectionChanged({
    if ($cmbNumFormat.SelectedItem) {
        $fmt = switch ($cmbNumFormat.SelectedIndex) {
            1 { "Currency_USD" }
            2 { "Currency_EUR" }
            3 { "Currency_GBP" }
            4 { "Currency_JPY" }
            5 { "Percent" }
            default { "General" }
        }
        Apply-FormatToSelectedCells "NumberFormat" $fmt
    }
})

# Quick Currency Buttons
$btnCurrUSD.add_Click({ Apply-FormatToSelectedCells "NumberFormat" "Currency_USD" })
$btnCurrEUR.add_Click({ Apply-FormatToSelectedCells "NumberFormat" "Currency_EUR" })
$btnCurrGBP.add_Click({ Apply-FormatToSelectedCells "NumberFormat" "Currency_GBP" })
$btnCurrJPY.add_Click({ Apply-FormatToSelectedCells "NumberFormat" "Currency_JPY" })
$btnPercent.add_Click({ Apply-FormatToSelectedCells "NumberFormat" "Percent" })

# Authentic Sort & Filter Ribbon Menu
$btnSortFilterMenu.add_Click({
    $btnSortFilterMenu.ContextMenu.IsOpen = $true
})

$menuSortAsc.add_Click({
    if ($txtNameBox.Text) {
        $coords = Convert-CellNameToCoords $txtNameBox.Text
        $engine.SortByColumn($coords.Col, $true)
        Refresh-GridUI
        $colName = Convert-ColIndexToName $coords.Col
        $txtStatus.Text = "Sorted column $colName Ascending (A to Z)."
    }
})

$menuSortDesc.add_Click({
    if ($txtNameBox.Text) {
        $coords = Convert-CellNameToCoords $txtNameBox.Text
        $engine.SortByColumn($coords.Col, $false)
        Refresh-GridUI
        $colName = Convert-ColIndexToName $coords.Col
        $txtStatus.Text = "Sorted column $colName Descending (Z to A)."
    }
})

$menuClearFilter.add_Click({
    $txtFilter.Text = ""
    Refresh-GridUI
    $txtStatus.Text = "Filter cleared."
})

# Column Header Click Sorting Handler
$gridSpreadsheet.add_Sorting({
    param($sender, $e)
    $e.Handled = $true
    $colIndex = $gridSpreadsheet.Columns.IndexOf($e.Column) + 1
    if ($colIndex -gt 0) {
        $sortAscending = ($e.Column.SortDirection -ne [System.ComponentModel.ListSortDirection]::Ascending)
        $engine.SortByColumn($colIndex, $sortAscending)
        $e.Column.SortDirection = if ($sortAscending) { [System.ComponentModel.ListSortDirection]::Ascending } else { [System.ComponentModel.ListSortDirection]::Descending }
        Refresh-GridUI
        $colName = Convert-ColIndexToName $colIndex
        $txtStatus.Text = "Sorted column $colName $(if ($sortAscending) { 'A to Z' } else { 'Z to A' })."
    }
})

$txtFilter.add_TextChanged({
    Refresh-GridUI
    if ($txtFilter.Text) {
        $txtStatus.Text = "Filtering by: '$($txtFilter.Text)'"
    } else {
        $txtStatus.Text = "Filter cleared."
    }
})

$btnClearFilter.add_Click({
    $txtFilter.Text = ""
    Refresh-GridUI
    $txtStatus.Text = "Filter cleared."
})

$btnToggleZebra.add_Click({
    $engine.IsZebraTable = -not $engine.IsZebraTable
    Refresh-GridUI
    $txtStatus.Text = if ($engine.IsZebraTable) { "Table Zebra Style Applied." } else { "Standard Table Style." }
})

# File Actions
$btnSave.add_Click({
    if ([string]::IsNullOrWhiteSpace($engine.FilePath)) {
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Filter = "CSV Files (*.csv)|*.csv|Excel Workbooks (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        if ($dialog.ShowDialog()) {
            $engine.FilePath = $dialog.FileName
        } else { return }
    }
    try {
        Export-PowerCellFile -engine $engine -filePath $engine.FilePath
        $txtStatus.Text = "Saved successfully to $($engine.FilePath)"
        $txtFileName.Text = [System.IO.Path]::GetFileName($engine.FilePath)
    } catch {
        [System.Windows.MessageBox]::Show("Save Error: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$btnQuickSave.add_Click({ $btnSave.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })

$btnExportExcel.add_Click({
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = "Excel Workbooks (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dialog.DefaultExt = ".xlsx"
    $dialog.FileName = if ($engine.FilePath) { [System.IO.Path]::ChangeExtension([System.IO.Path]::GetFileName($engine.FilePath), ".xlsx") } else { "ExportedSpreadsheet.xlsx" }
    
    if ($dialog.ShowDialog()) {
        try {
            Export-PowerCellFile -engine $engine -filePath $dialog.FileName
            $txtStatus.Text = "Exported successfully to $($dialog.FileName)"
            $txtFileName.Text = [System.IO.Path]::GetFileName($dialog.FileName)
            [System.Windows.MessageBox]::Show("Spreadsheet successfully exported to:`n$($dialog.FileName)", "Export Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            [System.Windows.MessageBox]::Show("Export Error: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

$btnOpen.add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = "Spreadsheets (*.csv;*.xlsx;*.tsv)|*.csv;*.xlsx;*.tsv|All Files (*.*)|*.*"
    if ($dialog.ShowDialog()) {
        try {
            $newEngine = [PowerCellEngine]::new()
            Import-PowerCellFile -engine $newEngine -filePath $dialog.FileName
            $engine.Cells = $newEngine.Cells
            $engine.FilePath = $newEngine.FilePath
            $engine.MaxRow = $newEngine.MaxRow
            $engine.MaxCol = $newEngine.MaxCol
            Refresh-GridUI
            $txtStatus.Text = "Opened file $($dialog.FileName)"
        } catch {
            [System.Windows.MessageBox]::Show("Open Error: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

$btnQuickOpen.add_Click({ $btnOpen.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)) })

$btnNewSheet.add_Click({
    $engine.Cells.Clear()
    $engine.FilePath = ""
    $engine.MaxRow = 50
    $engine.MaxCol = 26
    Refresh-GridUI
    $txtStatus.Text = "New Spreadsheet Created."
})

$btnInsertRow.add_Click({
    if ($txtNameBox.Text) {
        $coords = Convert-CellNameToCoords $txtNameBox.Text
        $engine.InsertRow($coords.Row)
        Refresh-GridUI
        $txtStatus.Text = "Inserted row at $($coords.Row)."
    }
})

$btnDeleteRow.add_Click({
    if ($txtNameBox.Text) {
        $coords = Convert-CellNameToCoords $txtNameBox.Text
        $engine.DeleteRow($coords.Row)
        Refresh-GridUI
        $txtStatus.Text = "Deleted row $($coords.Row)."
    }
})

$btnSum.add_Click({
    if ($txtNameBox.Text) {
        $coords = Convert-CellNameToCoords $txtNameBox.Text
        $targetRow = $coords.Row
        $colName = Convert-ColIndexToName $coords.Col
        if ($targetRow -gt 1) {
            $startRef = "$colName`1"
            $endRef = "$colName$($targetRow - 1)"
            $txtFormula.Text = "=SUM($($startRef):$($endRef))"
            $engine.SetCell($coords.Col, $targetRow, $txtFormula.Text)
            Refresh-GridUI
        }
    }
})

$btnAvg.add_Click({
    if ($txtNameBox.Text) {
        $coords = Convert-CellNameToCoords $txtNameBox.Text
        $targetRow = $coords.Row
        $colName = Convert-ColIndexToName $coords.Col
        if ($targetRow -gt 1) {
            $startRef = "$colName`1"
            $endRef = "$colName$($targetRow - 1)"
            $txtFormula.Text = "=AVG($($startRef):$($endRef))"
            $engine.SetCell($coords.Col, $targetRow, $txtFormula.Text)
            Refresh-GridUI
        }
    }
})

# Initial Data Grid Population
Refresh-GridUI

# Show WPF Window
[void]$window.ShowDialog()
