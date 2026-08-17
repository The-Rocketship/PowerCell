<#
.SYNOPSIS
    PowerCell - Native WPF Excel Spreadsheet Editor in PowerShell
.DESCRIPTION
    A self-contained Windows WPF GUI spreadsheet application for editing, viewing,
    evaluating formulas, and managing CSV, TSV, and XLSX files with an authentic Excel interface.
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
}

class PowerCellEngine {
    [hashtable]$Cells = @{}
    [int]$MaxRow = 50
    [int]$MaxCol = 26
    [string]$FilePath = ""
    [bool]$IsModified = $false

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

    [string] GetDisplayValue([int]$col, [int]$row) {
        $cell = $this.GetCell($col, $row)
        if ($cell.IsFormula) {
            return $cell.EvaluatedValue
        }
        return $cell.RawInput
    }

    [string] GetRawValue([int]$col, [int]$row) {
        $cell = $this.GetCell($col, $row)
        return $cell.RawInput
    }

    SetCell([int]$col, [int]$row, [string]$inputStr) {
        $key = $this.GetCellKey($col, $row)
        if ([string]::IsNullOrEmpty($inputStr)) {
            if ($this.Cells.ContainsKey($key)) {
                $this.Cells.Remove($key)
                $this.IsModified = $true
            }
            $this.RecalculateAll()
            return
        }

        $cell = [CellState]::new()
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
            [double]$num = 0
            if ([double]::TryParse($disp, [ref]$num)) {
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
        for ($pass = 0; $pass -lt 3; $pass++) {
            foreach ($key in @($this.Cells.Keys)) {
                $cell = $this.Cells[$key]
                if ($cell.IsFormula) {
                    $cell.EvaluatedValue = $this.EvaluateFormula($cell.Formula)
                }
            }
        }
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
        throw "Failed to open Excel file via COM: $_. For best results without Excel installed, save as CSV."
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
        throw "Excel export failed (Excel not installed). Saved spreadsheet as CSV to: $csvPath"
    }
}

# ============================================================================
# SECTION 3: WPF EXCEL GUI IMPLEMENTATION
# ============================================================================

$engine = [PowerCellEngine]::new()

if (-not [string]::IsNullOrWhiteSpace($FilePath)) {
    if (Test-Path $FilePath) {
        Import-PowerCellFile -engine $engine -filePath $FilePath
    } else {
        $engine.FilePath = [System.IO.Path]::GetFullPath($FilePath)
    }
}

# Define XAML Layout for Excel GUI Interface
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PowerCell Excel - Spreadsheet Editor" Height="700" Width="1100"
        WindowStartupLocation="CenterScreen" Background="#F3F3F3" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="Button" x:Key="RibbonBtn">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#168D4C"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="Button" x:Key="ActionBtn">
            <Setter Property="Background" Value="#E1DFDD"/>
            <Setter Property="BorderBrush" Value="#C8C6C4"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,4"/>
            <Setter Property="Margin" Value="2,0"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#D0CECC"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Title & Quick Header -->
            <RowDefinition Height="Auto"/> <!-- Ribbon Toolbar -->
            <RowDefinition Height="Auto"/> <!-- Formula Bar -->
            <RowDefinition Height="*"/>    <!-- Data Grid -->
            <RowDefinition Height="Auto"/> <!-- Status Bar -->
        </Grid.RowDefinitions>

        <!-- 1. Header Bar -->
        <Border Grid.Row="0" Background="#107C41" Padding="12,8">
            <DockPanel LastChildFill="True">
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Left" VerticalAlignment="Center">
                    <TextBlock Text="📊 PowerCell Excel" Foreground="#FFFFFF" FontWeight="Bold" FontSize="16" Margin="0,0,15,0"/>
                    <TextBlock Name="txtFileName" Text="Untitled.csv" Foreground="#DFF6DD" FontSize="13" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" HorizontalAlignment="Right">
                    <Button Name="btnQuickSave" Content="💾 Save" Style="{StaticResource RibbonBtn}"/>
                    <Button Name="btnQuickOpen" Content="📂 Open" Style="{StaticResource RibbonBtn}"/>
                    <Button Name="btnNewSheet" Content="📄 New" Style="{StaticResource RibbonBtn}"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- 2. Ribbon Toolbar -->
        <Border Grid.Row="1" Background="#F3F2F1" BorderBrush="#E1DFDD" BorderThickness="0,0,0,1" Padding="6,6">
            <WrapPanel Orientation="Horizontal">
                <!-- Group 1: File Actions -->
                <StackPanel Orientation="Horizontal" Margin="0,0,12,0">
                    <Button Name="btnSave" Content="💾 Save (Ctrl+S)" Style="{StaticResource ActionBtn}"/>
                    <Button Name="btnOpen" Content="📂 Open File" Style="{StaticResource ActionBtn}"/>
                    <Button Name="btnExportExcel" Content="📊 Export XLSX" Style="{StaticResource ActionBtn}"/>
                </StackPanel>
                <Rectangle Width="1" Height="24" Fill="#C8C6C4" Margin="4,0,8,0"/>
                <!-- Group 2: Insert / Delete -->
                <StackPanel Orientation="Horizontal" Margin="0,0,12,0">
                    <Button Name="btnInsertRow" Content="➕ Add Row" Style="{StaticResource ActionBtn}"/>
                    <Button Name="btnDeleteRow" Content="➖ Delete Row" Style="{StaticResource ActionBtn}"/>
                    <Button Name="btnInsertCol" Content="➕ Add Column" Style="{StaticResource ActionBtn}"/>
                </StackPanel>
                <Rectangle Width="1" Height="24" Fill="#C8C6C4" Margin="4,0,8,0"/>
                <!-- Group 3: Formulas -->
                <StackPanel Orientation="Horizontal">
                    <Button Name="btnSum" Content="∑ AutoSum" Style="{StaticResource ActionBtn}"/>
                    <Button Name="btnAvg" Content="x̄ Average" Style="{StaticResource ActionBtn}"/>
                    <Button Name="btnClearGrid" Content="🧹 Clear All" Style="{StaticResource ActionBtn}"/>
                </StackPanel>
            </WrapPanel>
        </Border>

        <!-- 3. Formula Bar -->
        <Border Grid.Row="2" Background="#FFFFFF" BorderBrush="#E1DFDD" BorderThickness="0,0,0,1" Padding="6,4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="70"/> <!-- Name Box -->
                    <ColumnDefinition Width="Auto"/> <!-- fx icon -->
                    <ColumnDefinition Width="*"/>   <!-- Formula Input -->
                </Grid.ColumnDefinitions>
                
                <TextBox Name="txtNameBox" Grid.Column="0" Text="A1" IsReadOnly="True" 
                         TextAlignment="Center" FontWeight="Bold" Background="#F3F2F1" BorderBrush="#C8C6C4" Padding="2,3"/>
                
                <TextBlock Grid.Column="1" Text=" fx " FontWeight="Bold" FontStyle="Italic" Foreground="#107C41" 
                           FontSize="14" VerticalAlignment="Center" Margin="6,0"/>
                
                <TextBox Name="txtFormula" Grid.Column="2" Padding="4,3" BorderBrush="#C8C6C4" FontSize="13"/>
            </Grid>
        </Border>

        <!-- 4. Main Data Grid -->
        <DataGrid Name="gridSpreadsheet" Grid.Row="3" AutoGenerateColumns="False" 
                  CanUserAddRows="False" CanUserDeleteRows="False" GridLinesVisibility="All"
                  HorizontalGridLinesBrush="#E1DFDD" VerticalGridLinesBrush="#E1DFDD"
                  HeadersVisibility="All" RowHeaderWidth="50" RowHeight="24"
                  Background="#FFFFFF" SelectionMode="Single" SelectionUnit="Cell"
                  FontSize="12">
            <DataGrid.Resources>
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background" Value="#F3F2F1"/>
                    <Setter Property="Foreground" Value="#323130"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="HorizontalContentAlignment" Value="Center"/>
                    <Setter Property="BorderBrush" Value="#C8C6C4"/>
                    <Setter Property="BorderThickness" Value="0,0,1,1"/>
                    <Setter Property="Padding" Value="4"/>
                </Style>
                <Style TargetType="DataGridRowHeader">
                    <Setter Property="Background" Value="#F3F2F1"/>
                    <Setter Property="Foreground" Value="#323130"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="HorizontalContentAlignment" Value="Center"/>
                    <Setter Property="BorderBrush" Value="#C8C6C4"/>
                    <Setter Property="BorderThickness" Value="0,0,1,1"/>
                </Style>
                <Style TargetType="DataGridCell">
                    <Style.Triggers>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter Property="Background" Value="#E1DFDD"/>
                            <Setter Property="BorderBrush" Value="#107C41"/>
                            <Setter Property="BorderThickness" Value="2"/>
                            <Setter Property="Foreground" Value="#000000"/>
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
$gridSpreadsheet = $window.FindName("gridSpreadsheet")
$txtNameBox      = $window.FindName("txtNameBox")
$txtFormula      = $window.FindName("txtFormula")
$txtFileName     = $window.FindName("txtFileName")
$txtStatus       = $window.FindName("txtStatus")
$txtCalcSummary  = $window.FindName("txtCalcSummary")

$btnSave         = $window.FindName("btnSave")
$btnOpen         = $window.FindName("btnOpen")
$btnExportExcel  = $window.FindName("btnExportExcel")
$btnInsertRow    = $window.FindName("btnInsertRow")
$btnDeleteRow    = $window.FindName("btnDeleteRow")
$btnInsertCol    = $window.FindName("btnInsertCol")
$btnSum          = $window.FindName("btnSum")
$btnAvg          = $window.FindName("btnAvg")
$btnClearGrid    = $window.FindName("btnClearGrid")
$btnQuickSave    = $window.FindName("btnQuickSave")
$btnQuickOpen    = $window.FindName("btnQuickOpen")
$btnNewSheet     = $window.FindName("btnNewSheet")

# DataTable Binding Setup
$table = [System.Data.DataTable]::new("Spreadsheet")

# Refresh DataGrid UI from PowerCellEngine
function Refresh-GridUI {
    $table.Rows.Clear()
    $table.Columns.Clear()
    $gridSpreadsheet.Columns.Clear()

    # Column 0 is Row Index placeholder
    $maxC = [math]::Max(20, $engine.MaxCol)
    $maxR = [math]::Max(40, $engine.MaxRow)

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
        $row = $table.NewRow()
        for ($c = 1; $c -le $maxC; $c++) {
            $colName = Convert-ColIndexToName $c
            $row[$colName] = $engine.GetDisplayValue($c, $r)
        }
        $table.Rows.Add($row)
    }

    $gridSpreadsheet.ItemsSource = $table.DefaultView
    if ($engine.FilePath) {
        $txtFileName.Text = [System.IO.Path]::GetFileName($engine.FilePath)
    } else {
        $txtFileName.Text = "Untitled.csv"
    }
}

# Attach Row Header Numbers (1..N)
$gridSpreadsheet.add_LoadingRow({
    param($sender, $e)
    $e.Row.Header = ($e.Row.GetIndex() + 1).ToString()
})

# Event: Selected Cell Changed
$gridSpreadsheet.add_SelectedCellsChanged({
    if ($gridSpreadsheet.SelectedCells.Count -gt 0) {
        $cellInfo = $gridSpreadsheet.SelectedCells[0]
        $rowIndex = $gridSpreadsheet.Items.IndexOf($cellInfo.Item) + 1
        $colIndex = $gridSpreadsheet.Columns.IndexOf($cellInfo.Column) + 1

        if ($rowIndex -gt 0 -and $colIndex -gt 0) {
            $cellRef = Convert-CoordsToCellName $colIndex $rowIndex
            $txtNameBox.Text = $cellRef
            $rawVal = $engine.GetRawValue($colIndex, $rowIndex)
            $txtFormula.Text = $rawVal

            # Calculate Selection Summary
            $dispVal = $engine.GetDisplayValue($colIndex, $rowIndex)
            [double]$num = 0
            if ([double]::TryParse($dispVal, [ref]$num)) {
                $txtCalcSummary.Text = "Sum: $num  |  Average: $num  |  Count: 1"
            } else {
                $txtCalcSummary.Text = "Ready"
            }
        }
    }
})

# Event: Formula Bar Text Change / Commit
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

# Event: Direct Cell Edit in DataGrid
$gridSpreadsheet.add_CellEditEnding({
    param($sender, $e)
    $rowIndex = $e.Row.GetIndex() + 1
    $colIndex = $e.Column.DisplayIndex + 1
    $editingElement = $e.EditingElement -as [System.Windows.Controls.TextBox]
    if ($null -ne $editingElement) {
        $newValue = $editingElement.Text
        $engine.SetCell($colIndex, $rowIndex, $newValue)
        # Deferred UI Refresh after edit completion
        $window.Dispatcher.InvokeAsync([Action]{ Refresh-GridUI })
    }
})

# Button Actions
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

$btnClearGrid.add_Click({
    $engine.Cells.Clear()
    Refresh-GridUI
    $txtStatus.Text = "Grid cleared."
})

# Initial Data Grid Population
Refresh-GridUI

# Show WPF Window
[void]$window.ShowDialog()
