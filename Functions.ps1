function Set-Term
{
	param ([String]$ExcelRange)
	switch ($ExcelRange.ToUpper())
	{
		'В КОНЦЕ СРОКА'{ [String][Char]69 }
		'ЕЖЕМЕСЯЧНО' { [String][Char]77 }
		default { $null }
	}
}
		
function Set-TermRange
{
	param ($Term, $Type)
	[Array]$Result = [RegEx]::Split($Term, [Char]45)
	switch ($Type)
	{
		'Begin' { $Result[0] }
		'End' { $Result[1] }
	}
}
		
function Get-Values
{
	param ([String]$ExcelRange)
	[Array]$Values = [RegEx]::Split($ExcelRange, [Char]59) | ? { $_ }
	return $Values
}
		
function Set-SumRangeFromValues
{
	param ($Range, $Type, $Multiple)
	[Array]$Result = [RegEx]::Split($Range, [Char]45)
	switch ($Type)
	{
		'Min' {
			if ([Int]$Result[0] -eq 0) { '{0}.00' -f $([Int]$Result[0] * $Multiple) }
			else { '{0}.01' -f $([Int]$Result[0] * $Multiple) }
		}
		'Max' {
			if ([Int]$Result[1] -eq 0) { $null }
			else { '{0}.00' -f $([Int]$Result[1] * $Multiple) }
		}
	}
}
		
function Set-SumRangeFromText
{
	param ([String]$ExcelRange)
	$FixedString = [RegEx]::Replace($ExcelRange.ToUpper(), '\W+', '')
	if ($FixedString.Contains('СВЫШЕ') -and $FixedString.Contains('ДО'))
	{
		[Array]$Ranges = [RegEx]::Split($FixedString, 'ДО')
		$From = [RegEx]::Matches($Ranges[0], '\d+')
		$To = [RegEx]::Matches($Ranges[1], '\d+')
		return ("{0}-{1}") -f $($From | Select -Expand Value), $($To | Select -Expand Value)
	}
	else
	{
		$Value = [RegEx]::Matches($FixedString, '\d+')
		if ($FixedString.Contains('СВЫШЕ'))
		{
			return ("{0}-0") -f $($Value | Select -Expand Value)
		}
		elseif ($FixedString.Contains('ДО'))
		{
			return ("0-{0}") -f $($Value | Select -Expand Value)
		}
	}
}
		
function ParseXlsx-KUAPRemains
{
	param (
		[String]$File,
		[String]$FileType,
		[Collections.ArrayList]$CurrencyList = @('USD', 'RUB')
	)
	$objExcel = New-Object OfficeOpenXml.ExcelPackage $File
	$WorkBook = $objExcel.Workbook
	$Sheets = $Workbook.Worksheets | Select-Object -ExpandProperty Name
	# Pattern
	$ColumnCheckOffest = 3
			
	$MultiDimensionalArray = New-Object System.Collections.Specialized.OrderedDictionary
	$MultiDimensionalArray["ЮЛ"] = New-Object System.Collections.Specialized.OrderedDictionary
	$MultiDimensionalArray["ФЛ"] = New-Object System.Collections.Specialized.OrderedDictionary

	if ($script:FileSheetHolder[$FileType].NamePattern.Count -gt 0)
	{
		New-LogRecord -EventType 'DEBUG' -Event $('Указано {0} шаблонов имен листов для файла с типом сделок "{1}."' -f $script:FileSheetHolder[$FileType].NamePattern.Count, $FileType)
		$SheetsFiltered = New-Object System.Collections.ArrayList
		foreach ($sfSheet in [String[]]$Sheets)
		{
			$si = 0
			foreach ($sfPattern in [String[]]$script:FileSheetHolder[$FileType].NamePattern)
			{
				if ($sfSheet -like $sfPattern)
				{
					$si++
					if ($si -eq 1)
					{
						[Void]$SheetsFiltered.Add($sfSheet)
						New-LogRecord -EventType 'DEBUG' -Event $('Шаблон "{0}" совпадает с листом "{1}."' -f $sfPattern, $sfSheet)
					}
					else
					{
						New-LogRecord -EventType 'ERROR' -Event $('Шаблон "{0}" совпадает более чем с одним листом. Совпадает с "{1}."' -f $sfPattern, $sfSheet)
						break
					}
				}
			}
		}
		[String[]]$Sheets = $SheetsFiltered
	}
			
	foreach ($Sheet in [String[]]$Sheets)
	{
		New-LogRecord -EventType 'DEBUG' -Event $('Обрабатывается файл "{0}" ({1}). Лист "{2}".' -f $File, $FileType, $Sheet)
		$CurrentProduct = $script:FileSheetHolder[$FileType] | ? { $Sheet -like $_.NamePattern } | Select -ExpandProperty ProductCode
		$ProductClient = $script:FileSheetHolder[$FileType] | ? { $Sheet -like $_.NamePattern } | Select -ExpandProperty ClientType
				
		$MultiDimensionalArray[$ProductClient][$CurrentProduct] = New-Object System.Collections.Specialized.OrderedDictionary
		foreach ($SC in [Collections.ArrayList]$CurrencyList)
		{
			$MultiDimensionalArray[$ProductClient][$CurrentProduct][$SC] = New-Object System.Collections.Specialized.OrderedDictionary
		}
				
		$CurrentSheet = $WorkBook.Worksheets[$Sheet]
				
		$EndRow = $CurrentSheet.Dimension.Rows
		$EndColumn = $CurrentSheet.Dimension.Columns
				
		$SeparatorAt = New-Object System.Collections.Specialized.OrderedDictionary
		$SeparatorNum = 1
				
		#Check if page start column
        $StartRow = 5
		$WorkStartColumn = 1
		$CheckTreshhold = 10
		$BreakTreshhold = 3
				
		$i = 1
		do
		{
			if ([String]::IsNullOrEmpty($CurrentSheet.GetValue($i, $WorkStartColumn)))
			{
				if ($BreakTreshhold -ge $WorkStartColumn)
				{
					if ($i -ge $CheckTreshhold)
					{
						$i = 1
						$WorkStartColumn++
					}
				}
				else { New-LogRecord -EventType 'ERROR' -Event $('Ошибка в формате файла "{0}".' -f $File); break }
				$i++
			}
			else { break }
		}
		while ($true)
				
		# определяем количество секций
		$c = 1
		for ($r = $StartRow; $r -lt $EndRow; $r++)
		{
			$Data = [String]$CurrentSheet.GetValue($r, $c)
			if ($c -eq 1)
			{
				if ([String]::IsNullOrEmpty($Data))
				{
					$IsEmptyRow = 0
					$c .. $ColumnCheckOffest | % { if ([String]::IsNullOrEmpty([String]$CurrentSheet.GetValue($r, $c + $_))) { $IsEmptyRow++ } }
					if ($IsEmptyRow -eq $ColumnCheckOffest) { $SeparatorAt.Add($SeparatorNum, $r); $SeparatorNum++ }
				}
			}
		}
				
		# всего секций
		$SectionList = New-Object System.Collections.Specialized.OrderedDictionary
				
		1..($SeparatorAt.Count + 1) | % {
			$SeparatorAtItem = $_
			if ($SeparatorAtItem -eq 1)
			{
				$StartRange = $StartRow
				$EndRange = $SeparatorAt.Item($SeparatorAtItem - 1) - 1
			}
			else
			{
				$StartRange = $SeparatorAt.Item($SeparatorAtItem - 2) + 1
				if ($SeparatorAt.Contains($SeparatorAtItem + 1))
				{
					$EndRange = $SeparatorAt.Item($SeparatorAtItem) - 1
				}
				else
				{
					$EndRange = $EndRow
				}
			}
			$SectionRange = $StartRange .. $EndRange
			$SectionList.Add($SeparatorAtItem, $SectionRange)
		}
				
		$PreviousCell = $null
		foreach ($Section in $SectionList.Keys)
		{
			foreach ($VerticalCell in $SectionList.Item($Section - 1))
			{
				$CurrentCell = [String]$CurrentSheet.GetValue($VerticalCell, $WorkStartColumn)
						
				if ($CurrentCell.Trim() -eq 'Валюта')
				{
					$SumRangePositions = New-Object System.Collections.Specialized.OrderedDictionary
					$TermDescriptionCol = $WorkStartColumn + 2
					foreach ($TDI in @($TermDescriptionCol .. $EndColumn))
					{
						$CellData = $CurrentSheet.GetValue($VerticalCell + 2, $TDI)
						if (-not [String]::IsNullOrEmpty($CellData))
						{
							$FixedCellRange = [RegEx]::Replace([RegEx]::Replace($CellData, [Char]47, '-'), 'дней', '')
							if ([Int][RegEx]::Split($FixedCellRange.Trim(), '-')[0] -gt 30)
							{
								$SumRangePositions[$FixedCellRange.Trim()] = $TDI
							}
						}
					}
				}
						
				if (-not [System.String]::IsNullOrEmpty($CurrentCell)) { $PreviousCell = $CurrentCell.Trim() }
						
				if ($CurrentCell.Trim() -ne 'Валюта' -and $PreviousCell -ne 'Валюта')
				{
							
					$SumRange = $WorkStartColumn + 1 #sum range from col num 2 of table
							
					$PercentPeriod = switch ($CurrentProduct) { 'DNS'{ 'E'; break }'DFO'{ 'M'; break } } #only these percent types exists for these products
							
					$RangeCell = Set-SumRangeFromText $CurrentSheet.GetValue($VerticalCell, $SumRange)
							
					$CurrencyDefinition = $script:Currencies[$PreviousCell]
							
					if ($CurrencyList.Contains($CurrencyDefinition))
					{
						if (-not $MultiDimensionalArray[$ProductClient][$CurrentProduct].Contains($CurrencyDefinition))
						{
							$MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition] = New-Object System.Collections.Specialized.OrderedDictionary
						}
						foreach ($SRPK in $SumRangePositions.Keys)
						{
							if (-not $MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition].Contains($SRPK))
							{
								$MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition][$SRPK] = New-Object System.Collections.Specialized.OrderedDictionary
							}
							$MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition][$SRPK][$RangeCell] = New-Object System.Collections.Specialized.OrderedDictionary
							$MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition][$SRPK][$RangeCell][$PercentPeriod] = [String]$CurrentSheet.GetValue($VerticalCell, $SumRangePositions[$SRPK])
						}
					}
					$CurrencyDefinition = $null
							
				}
			}
			$PreviousCell = $null
		}
	}
	Remove-Variable -Name objExcel;
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();
	$MultiDimensionalArray
}

function ParseXlsx-KUAPRemains_NEWFORMAT
{
	param (
		[String]$File,
		[String]$FileType,
		[Collections.ArrayList]$CurrencyList = @('USD', 'RUB')
	)
	$objExcel = New-Object OfficeOpenXml.ExcelPackage $File
	$WorkBook = $objExcel.Workbook
	$Sheets = $Workbook.Worksheets | Select-Object -ExpandProperty Name
	# Pattern
	[Int]$StartRow = 6
    [Int]$StartColumn = 4
	[Int]$ColumnCheckOffest = 3
			
	$MultiDimensionalArray = New-Object System.Collections.Specialized.OrderedDictionary
	$MultiDimensionalArray["ЮЛ"] = New-Object System.Collections.Specialized.OrderedDictionary
	$MultiDimensionalArray["ФЛ"] = New-Object System.Collections.Specialized.OrderedDictionary

	if ($script:FileSheetHolder[$FileType].NamePattern.Count -gt 0)
	{
		New-LogRecord -EventType 'DEBUG' -Event $('Указано {0} шаблонов имен листов для файла с типом сделок "{1}".' -f $script:FileSheetHolder[$FileType].NamePattern.Count, $FileType)
		$SheetsFiltered = New-Object System.Collections.ArrayList
		foreach ($sfSheet in [String[]]$Sheets)
		{
			$si = 0
			foreach ($sfPattern in [String[]]$script:FileSheetHolder[$FileType].NamePattern)
			{
				if ($sfSheet -like $sfPattern)
				{
					$si++
					if ($si -eq 1)
					{
						[Void]$SheetsFiltered.Add($sfSheet)
						New-LogRecord -EventType 'DEBUG' -Event $('Шаблон "{0}" совпадает с листом "{1}".' -f $sfPattern, $sfSheet)
					}
					else
					{
						New-LogRecord -EventType 'ERROR' -Event $('Шаблон "{0}" совпадает более чем с одним листом. Совпадает с "{1}".' -f $sfPattern, $sfSheet)
						break
					}
				}
			}
		}
		[String[]]$Sheets = $SheetsFiltered
	}
			
	foreach ($Sheet in [String[]]$Sheets)
	{
		New-LogRecord -EventType 'DEBUG' -Event $('Обрабатывается файл "{0}" ({1}). Лист "{2}".' -f $File, $FileType, $Sheet)
		$CurrentProduct = $script:FileSheetHolder[$FileType] | ? { $Sheet -like $_.NamePattern } | Select -ExpandProperty ProductCode
		$ProductClient = $script:FileSheetHolder[$FileType] | ? { $Sheet -like $_.NamePattern } | Select -ExpandProperty ClientType
				
		$MultiDimensionalArray[$ProductClient][$CurrentProduct] = New-Object System.Collections.Specialized.OrderedDictionary
		foreach ($SC in [Collections.ArrayList]$CurrencyList)
		{
			$MultiDimensionalArray[$ProductClient][$CurrentProduct][$SC] = New-Object System.Collections.Specialized.OrderedDictionary
		}
				
		$CurrentSheet = $WorkBook.Worksheets[$Sheet]
				
		$EndRow = $CurrentSheet.Dimension.Rows
		$EndColumn = $CurrentSheet.Dimension.Columns
				
		$SeparatorAt = New-Object System.Collections.Specialized.OrderedDictionary
		$SeparatorNum = 1
				
		# определяем количество секций
		$c = 4
		for ($r = $StartRow; $r -lt $EndRow; $r++)
		{
			$Data = [String]$CurrentSheet.GetValue($r, $c)
			if ($c -eq 1)
			{
				if ([String]::IsNullOrEmpty($Data))
				{
					$IsEmptyRow = 0
					$c .. $ColumnCheckOffest | % { if ([String]::IsNullOrEmpty([String]$CurrentSheet.GetValue($r, $c + $_))) { $IsEmptyRow++ } }
					if ($IsEmptyRow -eq $ColumnCheckOffest) { $SeparatorAt.Add($SeparatorNum, $r); $SeparatorNum++ }
				}
			}
		}
				
		# всего секций
		$SectionList = New-Object System.Collections.Specialized.OrderedDictionary
        
		if ($SeparatorAt.Count -gt 0){		
		    1..($SeparatorAt.Count + 1) | % {
			    $SeparatorAtItem = $_
			    if ($SeparatorAtItem -eq 1)
			    {
				    $StartRange = $StartRow
				    $EndRange = $SeparatorAt.Item($SeparatorAtItem - 1) - 1
			    }
			    else
			    {
				    $StartRange = $SeparatorAt.Item($SeparatorAtItem - 2) + 1
				    if ($SeparatorAt.Contains($SeparatorAtItem + 1))
				    {
					    $EndRange = $SeparatorAt.Item($SeparatorAtItem) - 1
				    }
				    else
				    {
					    $EndRange = $EndRow
				    }
			    }
			    $SectionRange = $StartRange .. $EndRange
			    $SectionList.Add($SeparatorAtItem, $SectionRange)
		    }
		}
        else{
            $SectionList.Add(1,@($StartRow..$EndRow))
        }		

		$PreviousCell = $null

		foreach ($Section in $SectionList.Keys)
		{
			foreach ($VerticalCell in $SectionList.Item($Section - 1))
			{
				$CurrentCell = [String]$CurrentSheet.GetValue($VerticalCell, $StartColumn)		
				if ($CurrentCell.Trim() -eq 'Валюта')
				{
					$SumRangePositions = New-Object System.Collections.Specialized.OrderedDictionary
					$TermDescriptionCol = $StartColumn + 2
					foreach ($TDI in @($TermDescriptionCol .. $EndColumn))
					{
						$CellData = $CurrentSheet.GetValue($VerticalCell + 2, $TDI)
						if (-not [String]::IsNullOrEmpty($CellData))
						{
							$FixedCellRange = [RegEx]::Replace([RegEx]::Replace($CellData, [Char]47, '-'), 'дней', '')
							if ([Int][RegEx]::Split($FixedCellRange.Trim(), '-')[0] -gt 30)
							{
								$SumRangePositions[$FixedCellRange.Trim()] = $TDI
							}
						}
					}
				}
						
				if (-not [System.String]::IsNullOrEmpty($CurrentCell)) { $PreviousCell = $CurrentCell.Trim() }
						
				if ($CurrentCell.Trim() -ne 'Валюта' -and $PreviousCell -ne 'Валюта')
				{
							
					$SumRange = $StartColumn + 1 #sum range from col num 2 of table
							
					$PercentPeriod = switch ($CurrentProduct) { 'DNS'{ 'E'; break }'DFO'{ 'M'; break } } #only these percent types exists for these products
							
					$RangeCell = Set-SumRangeFromText $CurrentSheet.GetValue($VerticalCell, $SumRange)
							
					$CurrencyDefinition = $script:Currencies[$PreviousCell]
							
					if ($CurrencyList.Contains($CurrencyDefinition))
					{
						if (-not $MultiDimensionalArray[$ProductClient][$CurrentProduct].Contains($CurrencyDefinition))
						{
							$MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition] = New-Object System.Collections.Specialized.OrderedDictionary
						}
						foreach ($SRPK in $SumRangePositions.Keys)
						{
							if (-not $MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition].Contains($SRPK))
							{
								$MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition][$SRPK] = New-Object System.Collections.Specialized.OrderedDictionary
							}
							$MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition][$SRPK][$RangeCell] = New-Object System.Collections.Specialized.OrderedDictionary
							$MultiDimensionalArray[$ProductClient][$CurrentProduct][$CurrencyDefinition][$SRPK][$RangeCell][$PercentPeriod] = [String]$CurrentSheet.GetValue($VerticalCell, $SumRangePositions[$SRPK])
						}
					}
					$CurrencyDefinition = $null
							
				}
			}
			$PreviousCell = $null
		}
	}
	Remove-Variable -Name objExcel;
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();
	$MultiDimensionalArray
}
		
function ParseXlsx-KUAPDeposit
{
			param (
				[String]$File,
				[String]$FileType,
				[String]$ClientType
			)
			$objExcel = New-Object OfficeOpenXml.ExcelPackage $File
			$WorkBook = $objExcel.Workbook
			$Sheets = $Workbook.Worksheets | Select-Object -ExpandProperty Name
			$ColumnOffsetByClientType = switch ($ClientType) { 'ЮЛ'{ 3 }'ФЛ'{ 2 } }
			
			$ResultMDObject = New-Object System.Collections.Specialized.OrderedDictionary
			$ResultMDObject[$ClientType] = New-Object System.Collections.Specialized.OrderedDictionary
			
			if ($script:FileSheetHolder[$FileType].NamePattern.Count -gt 0)
			{
				New-LogRecord -EventType 'DEBUG' -Event $('Указано {0} шаблонов имен листов для файла с типом сделок "{1}".' -f $script:FileSheetHolder[$FileType].NamePattern.Count, $FileType)
				$SheetsFiltered = New-Object System.Collections.ArrayList
				foreach ($sfSheet in [String[]]$Sheets)
				{
					$si = 0
					foreach ($sfPattern in [String[]]$script:FileSheetHolder[$FileType].NamePattern)
					{
						if ($sfSheet -like $sfPattern)
						{
							$si++
							if ($si -eq 1)
							{
								[Void]$SheetsFiltered.Add($sfSheet)
								New-LogRecord -EventType 'DEBUG' -Event $('Шаблон "{0}" совпадает с листом "{1}".' -f $sfPattern, $sfSheet)
							}
							else
							{
								New-LogRecord -EventType 'ERROR' -Event $('Шаблон "{0}" совпадает более чем с одним листом. Совпадает с "{1}".' -f $sfPattern, $sfSheet)
								break
							}
						}
					}
				}
				[String[]]$Sheets = $SheetsFiltered
			}
			
			foreach ($Sheet in [String[]]$Sheets)
			{
				New-LogRecord -EventType 'DEBUG' -Event $('Обрабатывается файл "{0}" ({1}). Лист "{2}".' -f $File, $FileType, $Sheet)
				$CurrentProduct = $script:FileSheetHolder[$FileType] | ? { $Sheet -like $_.NamePattern } | Select -ExpandProperty ProductCode
				
				$CurrentSheet = $WorkBook.Worksheets[$Sheet]
				
				$MultiDimensionalArray = New-Object System.Collections.Specialized.OrderedDictionary
				$MultiDimensionalArray[$CurrentProduct] = New-Object System.Collections.Specialized.OrderedDictionary
				$MultiDimensionalArray[$CurrentProduct][$ClientType] = New-Object System.Collections.Specialized.OrderedDictionary
				
				#Check if page start column
				$WorkStartColumn = 1
				$CheckTreshhold = 10
				$BreakTreshhold = 3
				
				$i = 1
				do
				{
					if ([String]::IsNullOrEmpty($CurrentSheet.GetValue($i, $WorkStartColumn)))
					{
						if ($BreakTreshhold -ge $WorkStartColumn)
						{
							if ($i -ge $CheckTreshhold)
							{
								$i = 1
								$WorkStartColumn++
							}
						}
						else
						{
							New-LogRecord -EventType 'ERROR' -Event $('Ошибка в формате файла "{0}".' -f $File); break
						}
						$i++
					}
					else { break }
				}
				while ($true)
				
				#Getting currencies and document logical bounds
				$CurrentCell = $null
				$PreviousCell = $null
				$i = 1
				do
				{
					$CurrentCell = [String]$CurrentSheet.GetValue($i, $WorkStartColumn)
					if ($CurrentCell.ToUpper() -like '*ДЕПОЗИТ*') { $StartRow = $i + 1 }
					if (-not [String]::IsNullOrEmpty($CurrentCell) -and $CurrentCell.ToUpper() -notlike '*МАКСИМАЛЬНЫЕ*' -and $CurrentCell.ToUpper() -notlike '*ДЕПОЗИТ*' -and $CurrentCell.ToUpper() -notlike 'ВАЛЮТА*' -and $CurrentCell.ToUpper() -ne '* - ВКЛЮЧИТЕЛЬНО')
					{
						$MultiDimensionalArray[$CurrentProduct][$ClientType][$script:Currencies[$CurrentCell]] = New-Object System.Collections.Specialized.OrderedDictionary
					}
					$CRD = $CurrentSheet.Dimension | Select -Expand Rows
					if ($i -ge $CRD)
					{
						$EndRow = $CRD
						do
						{
							$ERFCounter = 0
							foreach ($iERF in @(0..10))
							{
								$CSER = $CurrentSheet.GetValue($EndRow, $WorkStartColumn + $iERF)
								if ([String]::IsNullOrEmpty($CSER) -and $CSER -ne '* - ВКЛЮЧИТЕЛЬНО') { $ERFCounter++ }
							}
							if ($ERFCounter -gt (12 - $ColumnOffsetByClientType)) { $EndRow-- }
							else { break }
						}
						while ($true)
						break
					}
					$i++
				}
				while ($true)
				
				$DatesRangeFrom = $WorkStartColumn + $ColumnOffsetByClientType
				$GoodRangeFrom = $DatesRangeFrom
				$CurrentCell = $null
				$PreviousCell = $null
				do
				{
					$CurrentCell = [String]$CurrentSheet.GetValue($StartRow + 2, $DatesRangeFrom)
					if (-not [String]::IsNullOrEmpty($CurrentCell))
					{
						if ([Int][RegEx]::Split($CurrentCell, '-')[0] -gt 30)
						{
							foreach ($CurrencyKey in $MultiDimensionalArray[$CurrentProduct][$ClientType].Keys)
							{
								$MultiDimensionalArray[$CurrentProduct][$ClientType][$CurrencyKey][$CurrentCell] = New-Object System.Collections.Specialized.OrderedDictionary
								$MultiDimensionalArray[$CurrentProduct][$ClientType][$CurrencyKey][$CurrentCell] = $DatesRangeFrom
							}
						}
						else { $GoodRangeFrom++ }
						$DatesRangeFrom++
					}
					else { break }
				}
				while ($true)
				$DatesRangeTo = $DatesRangeFrom
				
				
				$CurrencyHolder = New-Object System.Collections.Specialized.OrderedDictionary
				$Currencies.Keys | % {
					$CurrencyHolder[$script:Currencies[$_]] = New-Object System.Collections.ArrayList
				}
				
				$CurrentCell = $null
				$PreviousCell = $null
				foreach ($Row in @($StartRow .. $EndRow))
				{
					$CurrentCell = [String]$CurrentSheet.GetValue($Row, $WorkStartColumn)
					if (-not [String]::IsNullOrEmpty($CurrentCell)) { $PreviousCell = $CurrentCell.Trim() }
					if ($CurrentCell.Trim() -ne 'Валюта вклада' -and $PreviousCell -ne 'Валюта вклада')
					{
						[Void]$CurrencyHolder[$script:Currencies[$PreviousCell]].Add($Row)
					}
				}
				
				$CurrentCell = $null
				$PreviousCell = $null
				switch ($ColumnOffsetByClientType)
				{
					3 {
						$SumHolder = New-Object System.Collections.Specialized.OrderedDictionary
						foreach ($Key in @($CurrencyHolder.Keys))
						{
							$SumHolder[$Key] = New-Object System.Collections.Specialized.OrderedDictionary
							foreach ($k in @($CurrencyHolder[$Key]))
							{
								$CurrentCell = [String]$CurrentSheet.GetValue($k, $WorkStartColumn + 1)
								if (-not [String]::IsNullOrEmpty($CurrentCell)) { $PreviousCell = $CurrentCell.Trim() }
								else { $CurrentCell = $PreviousCell }
								$SumRange = Set-SumRangeFromText $CurrentCell
								$SumHolder[$Key][$SumRange] += ([String]$k + [Char]59)
							}
						}
					}
					2   {
						$SumHolder = New-Object System.Collections.Specialized.OrderedDictionary
						foreach ($Key in @($CurrencyHolder.Keys))
						{
							$SumHolder[$Key] = New-Object System.Collections.Specialized.OrderedDictionary
							foreach ($k in @($CurrencyHolder[$Key]))
							{
								$SumRange = '0-0'
								$SumHolder[$Key][$SumRange] += ([String]$k + [Char]59)
							}
						}
					}
				}
				
				$CurrentCell = $null
				$PreviousCell = $null
				$TermHolder = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($sh in @($SumHolder.Keys))
				{
					$TermHolder[$sh] = New-Object System.Collections.Specialized.OrderedDictionary
					foreach ($s in @($SumHolder[$sh].Keys))
					{
						$TermHolder[$sh][$s] = New-Object System.Collections.Specialized.OrderedDictionary
						foreach ($r in @(Get-Values $SumHolder[$sh][$s]))
						{
							$CurrentCell = [String]$CurrentSheet.GetValue($r, $WorkStartColumn + $ColumnOffsetByClientType - 1)
							if (-not [String]::IsNullOrEmpty($CurrentCell)) { $PreviousCell = $CurrentCell.Trim() }
							else { $CurrentCell = $PreviousCell }
							$PayMethod = Set-Term $CurrentCell
							$TermHolder[$sh][$s][$PayMethod] += ([String]$r + [Char]59)
						}
					}
				}
				
				$CSC = New-Object System.Collections.ArrayList
				foreach ($tKey in $TermHolder.Keys) { if ($TermHolder[$tKey].Count -gt 0) { [Void]$CSC.Add($tKey) } }
				
				$ResultMDObject[$ClientType][$CurrentProduct] = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($C in $CSC)
				{
					$ResultMDObject[$ClientType][$CurrentProduct][$C] = New-Object System.Collections.Specialized.OrderedDictionary
					foreach ($rKey in @($MultiDimensionalArray[$CurrentProduct][$ClientType][$C].Keys))
					{
						$ResultMDObject[$ClientType][$CurrentProduct][$C][$rKey] = New-Object System.Collections.Specialized.OrderedDictionary
						foreach ($sKey in @($TermHolder[$C].Keys))
						{
							$ResultMDObject[$ClientType][$CurrentProduct][$C][$rKey][$sKey] = New-Object System.Collections.Specialized.OrderedDictionary
							foreach ($pKey in @($TermHolder[$C][$sKey].Keys))
							{
								$ResultMDObject[$ClientType][$CurrentProduct][$C][$rKey][$sKey][$pKey] = New-Object System.Collections.Specialized.OrderedDictionary
								[Int]$RowCell = [RegEx]::Replace($TermHolder[$C][$sKey][$pKey], [Char]59, '')
								[Int]$ColumnCell = $MultiDimensionalArray[$CurrentProduct][$ClientType][$C][$rKey]
								$ResultMDObject[$ClientType][$CurrentProduct][$C][$rKey][$sKey][$pKey] = [String]$CurrentSheet.GetValue($RowCell, $ColumnCell)
							}
						}
					}
				}
			}
			Remove-Variable -Name objExcel;
			[System.GC]::Collect();
			[System.GC]::WaitForPendingFinalizers();
			$ResultMDObject
		}

function ParseXlsx-KUAPDeposit_NEWFORMAT
{
			param (
				[String]$File,
				[String]$FileType,
				[String]$ClientType
			)
			$objExcel = New-Object OfficeOpenXml.ExcelPackage $File
			$WorkBook = $objExcel.Workbook
			$Sheets = $Workbook.Worksheets | Select-Object -ExpandProperty Name
			$ColumnOffsetByClientType = switch ($ClientType) { 'ЮЛ'{ 3 }'ФЛ'{ 2 } }
			
			$ResultMDObject = New-Object System.Collections.Specialized.OrderedDictionary
			$ResultMDObject[$ClientType] = New-Object System.Collections.Specialized.OrderedDictionary

			if ($script:FileSheetHolder[$FileType].NamePattern.Count -gt 0)
			{
				New-LogRecord -EventType 'DEBUG' -Event $('Указано {0} шаблонов имен листов для файла с типом сделок "{1}".' -f $script:FileSheetHolder[$FileType].NamePattern.Count, $FileType)
				$SheetsFiltered = New-Object System.Collections.ArrayList
				foreach ($sfSheet in [String[]]$Sheets)
				{
					$si = 0
					foreach ($sfPattern in [String[]]$script:FileSheetHolder[$FileType].NamePattern)
					{
						if ($sfSheet -like $sfPattern)
						{
							$si++
							if ($si -eq 1)
							{
								[Void]$SheetsFiltered.Add($sfSheet)
								New-LogRecord -EventType 'DEBUG' -Event $('Шаблон "{0}" совпадает с листом "{1}".' -f $sfPattern, $sfSheet)
							}
							else
							{
								New-LogRecord -EventType 'ERROR' -Event $('Шаблон "{0}" совпадает более чем с одним листом. Совпадает с "{1}".' -f $sfPattern, $sfSheet)
								break
							}
						}
					}
				}
				[String[]]$Sheets = $SheetsFiltered
			}
			
			foreach ($Sheet in [String[]]$Sheets)
			{
				New-LogRecord -EventType 'DEBUG' -Event $('Обрабатывается файл "{0}" ({1}). Лист "{2}".' -f $File, $FileType, $Sheet)
				$CurrentProduct = $script:FileSheetHolder[$FileType] | ? { $Sheet -like $_.NamePattern } | Select -ExpandProperty ProductCode
				
				$CurrentSheet = $WorkBook.Worksheets[$Sheet]

				$MultiDimensionalArray = New-Object System.Collections.Specialized.OrderedDictionary
				$MultiDimensionalArray[$CurrentProduct] = New-Object System.Collections.Specialized.OrderedDictionary
				$MultiDimensionalArray[$CurrentProduct][$ClientType] = New-Object System.Collections.Specialized.OrderedDictionary
				
				#Check if page start column
				$WorkStartColumn = 5
				
				#Getting currencies and document logical bounds
				$CurrentCell = $null
				$PreviousCell = $null
				$i = 1
				do
				{
					$CurrentCell = [String]$CurrentSheet.GetValue($i, $WorkStartColumn)
					if ($CurrentCell.ToUpper() -like '*ДЕПОЗИТ*') { $StartRow = $i + 1 }
					if (-not [String]::IsNullOrEmpty($CurrentCell) -and $CurrentCell.ToUpper() -notlike '*МАКСИМАЛЬНЫЕ*' -and $CurrentCell.ToUpper() -notlike '*ДЕПОЗИТ*' -and $CurrentCell.ToUpper() -notlike 'ВАЛЮТА*' -and $CurrentCell.ToUpper() -ne '* - ВКЛЮЧИТЕЛЬНО')
					{
						$MultiDimensionalArray[$CurrentProduct][$ClientType][$script:Currencies[$CurrentCell]] = New-Object System.Collections.Specialized.OrderedDictionary
					}
					$CRD = $CurrentSheet.Dimension | Select -Expand Rows
					if ($i -ge $CRD)
					{
						$EndRow = $CRD
						do
						{
							$ERFCounter = 0
							foreach ($iERF in @(0..10))
							{
								$CSER = $CurrentSheet.GetValue($EndRow, $WorkStartColumn + $iERF)
								if ([String]::IsNullOrEmpty($CSER) -and $CSER -ne '* - ВКЛЮЧИТЕЛЬНО') { $ERFCounter++ }
							}
							if ($ERFCounter -gt (12 - $ColumnOffsetByClientType)) { $EndRow-- }
							else { break }
						}
						while ($true)
						break
					}
					$i++
				}
				while ($true)
				
				$DatesRangeFrom = $WorkStartColumn + $ColumnOffsetByClientType
				$GoodRangeFrom = $DatesRangeFrom
				$CurrentCell = $null
				$PreviousCell = $null
				do
				{
					$CurrentCell = [String]$CurrentSheet.GetValue($StartRow + 2, $DatesRangeFrom)
					if (-not [String]::IsNullOrEmpty($CurrentCell))
					{
						if ([Int][RegEx]::Split($CurrentCell, '-')[0] -gt 30)
						{
							foreach ($CurrencyKey in $MultiDimensionalArray[$CurrentProduct][$ClientType].Keys)
							{
								$MultiDimensionalArray[$CurrentProduct][$ClientType][$CurrencyKey][$CurrentCell] = New-Object System.Collections.Specialized.OrderedDictionary
								$MultiDimensionalArray[$CurrentProduct][$ClientType][$CurrencyKey][$CurrentCell] = $DatesRangeFrom
							}
						}
						else { $GoodRangeFrom++ }
						$DatesRangeFrom++
					}
					else { break }
				}
				while ($true)
				$DatesRangeTo = $DatesRangeFrom
				
				
				$CurrencyHolder = New-Object System.Collections.Specialized.OrderedDictionary
				$Currencies.Keys | % {
					$CurrencyHolder[$script:Currencies[$_]] = New-Object System.Collections.ArrayList
				}
				
				$CurrentCell = $null
				$PreviousCell = $null
				foreach ($Row in @($StartRow .. $EndRow))
				{
					$CurrentCell = [String]$CurrentSheet.GetValue($Row, $WorkStartColumn)
					if (-not [String]::IsNullOrEmpty($CurrentCell)) { $PreviousCell = $CurrentCell.Trim() }
					if ($CurrentCell.Trim() -ne 'Валюта вклада' -and $PreviousCell -ne 'Валюта вклада')
					{
						[Void]$CurrencyHolder[$script:Currencies[$PreviousCell]].Add($Row)
					}
				}
				
				$CurrentCell = $null
				$PreviousCell = $null
				switch ($ColumnOffsetByClientType)
				{
					3 {
						$SumHolder = New-Object System.Collections.Specialized.OrderedDictionary
						foreach ($Key in @($CurrencyHolder.Keys))
						{
							$SumHolder[$Key] = New-Object System.Collections.Specialized.OrderedDictionary
							foreach ($k in @($CurrencyHolder[$Key]))
							{
								$CurrentCell = [String]$CurrentSheet.GetValue($k, $WorkStartColumn + 1)
								if (-not [String]::IsNullOrEmpty($CurrentCell)) { $PreviousCell = $CurrentCell.Trim() }
								else { $CurrentCell = $PreviousCell }
								$SumRange = Set-SumRangeFromText $CurrentCell
								$SumHolder[$Key][$SumRange] += ([String]$k + [Char]59)
							}
						}
					}
					2   {
						$SumHolder = New-Object System.Collections.Specialized.OrderedDictionary
						foreach ($Key in @($CurrencyHolder.Keys))
						{
							$SumHolder[$Key] = New-Object System.Collections.Specialized.OrderedDictionary
							foreach ($k in @($CurrencyHolder[$Key]))
							{
								$SumRange = '0-0'
								$SumHolder[$Key][$SumRange] += ([String]$k + [Char]59)
							}
						}
					}
				}
				
				$CurrentCell = $null
				$PreviousCell = $null
				$TermHolder = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($sh in @($SumHolder.Keys))
				{
					$TermHolder[$sh] = New-Object System.Collections.Specialized.OrderedDictionary
					foreach ($s in @($SumHolder[$sh].Keys))
					{
						$TermHolder[$sh][$s] = New-Object System.Collections.Specialized.OrderedDictionary
						foreach ($r in @(Get-Values $SumHolder[$sh][$s]))
						{
							$CurrentCell = [String]$CurrentSheet.GetValue($r, $WorkStartColumn + $ColumnOffsetByClientType - 1)
							if (-not [String]::IsNullOrEmpty($CurrentCell)) { $PreviousCell = $CurrentCell.Trim() }
							else { $CurrentCell = $PreviousCell }
							$PayMethod = Set-Term $CurrentCell
							$TermHolder[$sh][$s][$PayMethod] += ([String]$r + [Char]59)
						}
					}
				}
				
				$CSC = New-Object System.Collections.ArrayList
				foreach ($tKey in @($TermHolder.Keys)) { if ($TermHolder[$tKey].Count -gt 0) { [Void]$CSC.Add($tKey) } }
				
				$ResultMDObject[$ClientType][$CurrentProduct] = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($C in @($CSC))
				{
					$ResultMDObject[$ClientType][$CurrentProduct][$C] = New-Object System.Collections.Specialized.OrderedDictionary
					foreach ($rKey in @($MultiDimensionalArray[$CurrentProduct][$ClientType][$C].Keys))
					{
						$ResultMDObject[$ClientType][$CurrentProduct][$C][$rKey] = New-Object System.Collections.Specialized.OrderedDictionary
						foreach ($sKey in @($TermHolder[$C].Keys))
						{
							$ResultMDObject[$ClientType][$CurrentProduct][$C][$rKey][$sKey] = New-Object System.Collections.Specialized.OrderedDictionary
							foreach ($pKey in @($TermHolder[$C][$sKey].Keys))
							{
								$ResultMDObject[$ClientType][$CurrentProduct][$C][$rKey][$sKey][$pKey] = New-Object System.Collections.Specialized.OrderedDictionary
								[Int]$RowCell = [RegEx]::Replace($TermHolder[$C][$sKey][$pKey], [Char]59, '')
								[Int]$ColumnCell = $MultiDimensionalArray[$CurrentProduct][$ClientType][$C][$rKey]
								$ResultMDObject[$ClientType][$CurrentProduct][$C][$rKey][$sKey][$pKey] = [String]$CurrentSheet.GetValue($RowCell, $ColumnCell)
							}
						}
					}
				}
			}
			Remove-Variable -Name objExcel;
			[System.GC]::Collect();
			[System.GC]::WaitForPendingFinalizers();
			$ResultMDObject
		}
