# requires -version 2
<#
.SYNOPSIS
	Script que convierte varios reportes de QAD en formato CSV a formato Excel 2010+
	
.DESCRIPTION
	El script está pensado para mostrar la información de seguridad de los perfiles de usuarios configurados en QAD.Net 2008.1se
	
.INPUTS
	Ninguno
	
.OUTPUTS
	Write-Host para poner en el prompt lo que está realizando
	Write-Exception para las excepciones
	
.NOTES
	Version:	2.1
	Author:		Santiago Platero
	Creation Date: 	04/04/2018
	Purpose/Change:	Fork del script original para cambiar todos los archivos de un directorio
	
.EXAMPLE
	> ./masscsv2xlsx.ps1
#>

#---------------------------------------------------------[Inicializaciones]--------------------------------------------------------

# Inicializaciones de variables, CSV de origen, XLSX de destino, path del logo y el delimitador que usa el CSV de origen
$imgPath = "W:\splatero\logo.png"
$folder = "W:\splatero"
$files = Get-ChildItem $folder *.txt
$delimiter = ";"

#----------------------------------------------------------[Declaraciones]----------------------------------------------------------

# Información del script
$scriptVersion = "2.1"
$scriptName = $MyInvocation.MyCommand.Name

#-----------------------------------------------------------[Ejecución]------------------------------------------------------------

# Función para que aparezca un mensaje en caso de error
Function Write-Exception
{
	Write-Host "[$(Get-Date -format $($dateFormat))] ERROR: $($_.Exception.Message)"
	exit 1
}

# Control de errores
try {
	
	# Creación de objeto COM de Microsoft Excel
	Write-Host "[$(Get-Date -format $($dateFormat))] INFO: Creando planilla de Excel"
	$excel = New-Object -ComObject excel.application 
	
	foreach ($file in $files) {
		$csv = $folder + "\" + $file
		$xlsx = $folder + "\" + $file + ".xlsx"
		$workbook = $excel.Workbooks.Add(1)
		$worksheet = $workbook.worksheets.Item(1)

		# Establecemos algunos tamaños específicos para columnas y filas
		Write-Host "[$(Get-Date -format $($dateFormat))] INFO: Estableciendo cuestiones de formato a la planilla"
		$worksheet.Rows("1").RowHeight=74
		$worksheet.Rows("2").RowHeight=22
		$worksheet.Rows("4:6").RowHeight=28
		$worksheet.Rows("3").RowHeight=14
		$worksheet.Rows("7").RowHeight=14
		$worksheet.Columns("A").ColumnWidth=11
		$worksheet.Columns("B").ColumnWidth=9
		# NEW: aplicamos a toda la columna C la propiedad reduce hasta ajustar para que no se pase del ancho de página
		$worksheet.Range("C:C").ShrinkToFit = $true
		
		$worksheet.Columns("D").ColumnWidth=33
		$worksheet.Columns("E").ColumnWidth=16
	
		# Alineaciones y la fuente que se debe utilizar
		$worksheet.Columns("A:E").HorizontalAlignment = -4108 # ese negativo significa "centro"
		$worksheet.Columns("A:E").VerticalAlignment = -4108
		$worksheet.Columns("A:E").Font.Name = "Trebuchet MS"
		$worksheet.Columns("A:E").Font.Size = 10

		# Generamos el header del documento, combinamos las celdas para darle el formato que tiene que llevar
		Write-Host "[$(Get-Date -format $($dateFormat))] INFO: Generando header del documento, combinando celdas, poniendo algunos bordes"
		$worksheet.Cells.item(1,1) = "Perfiles de Usuario de QAD.Net"
		$worksheet.Cells.item(1,1).Font.Size = 14
		$mergecells1 = $worksheet.Range("A1:D1")
		$mergecells1.Select()
		$mergecells1.MergeCells = $true

		$worksheet.Cells.item(2,1) = "FOGL-IT-MAR-0079/01"
		$mergecells2 = $worksheet.Range("A2:E2")
		$mergecells2.Select()
		$mergecells2.MergeCells = $true

		$worksheet.Cells.item(4,1) = "SISTEMA: QAD.Net"
		$worksheet.Cells.item(4,1).Font.Bold = $true
		$mergecells3 = $worksheet.Range("A4:B4")
		$mergecells3.Select()
		$mergecells3.MergeCells = $true
		$mergecells3.Interior.ColorIndex = 37 # El 38 refiere a un color celeste

		$worksheet.Cells.item(4,3) = "Firma / Aclaracion"
		$worksheet.Cells.item(4,3).Font.Bold = $true
		$mergecells6 = $worksheet.Range("C4:D4")
		$mergecells6.Select()
		$mergecells6.MergeCells = $true
		$mergecells6.Interior.ColorIndex = 15 # El 15 es un gris claro

		$worksheet.Cells.item(4,5) = "Fecha"
		$worksheet.Cells.item(4,5).Font.Bold = $true
		$worksheet.Cells.item(4,5).Borders.LineStyle = 1 # El 1 es para una linea de borde continua
		$worksheet.Cells.item(4,5).Interior.ColorIndex = 15

		$worksheet.Cells.item(5,1) = "Seguridad Informatica"
		$worksheet.Cells.item(5,1).Font.Bold = $true
		$mergecells4 = $worksheet.Range("A5:B5")
		$mergecells4.Select()
		$mergecells4.MergeCells = $true
		$mergecells4.Interior.ColorIndex = 15

		$worksheet.Cells.item(6,1) = "Aplicaciones"
		$worksheet.Cells.item(6,1).Font.Bold = $true
		$mergecells5 = $worksheet.Range("A6:B6")
		$mergecells5.Select()
		$mergecells5.MergeCells = $true
		$mergecells5.Interior.ColorIndex = 15

		$mergecells7 = $worksheet.Range("C5:D5")
		$mergecells7.Select()
		$mergecells7.MergeCells = $true
		$mergecells7 = $worksheet.Range("C6:D6")
		$mergecells7.Select()
		$mergecells7.MergeCells = $true

		# Agregamos la imagen del logo
		Write-Host "[$(Get-Date -format $($dateFormat))] INFO: Agregando imagen del logo de MTV"
		$img = $worksheet.Shapes.AddPicture($imgPath, $false, $true, 352, 11, 82, 50)

		# Agregamos algunos bordes para presentar un poco más prolija la info, el cuadro de firmantes, etc.
		Write-Host "[$(Get-Date -format $($dateFormat))] INFO: Mas bordes..."
		$worksheet.Range("A1:D1").Borders.LineStyle = 1
		$worksheet.Range("E1").Borders.LineStyle = 1
		$worksheet.Range("A2:E2").Borders.LineStyle = 1
		$worksheet.Range("A4:B4").Borders.LineStyle = 1
		$worksheet.Range("C4:D4").Borders.LineStyle = 1
		$worksheet.Range("E4").Borders.LineStyle = 1
		$worksheet.Range("A5:B5").Borders.LineStyle = 1
		$worksheet.Range("C5:D5").Borders.LineStyle = 1
		$worksheet.Range("E5").Borders.LineStyle = 1
		$worksheet.Range("A6:B6").Borders.LineStyle = 1
		$worksheet.Range("C6:D6").Borders.LineStyle = 1
		$worksheet.Range("E6").Borders.LineStyle = 1

		# Armamos la query para planchar el CSV en lo que queda de planilla
		Write-Host "[$(Get-Date -format $($dateFormat))] INFO: Armando query para importar los datos desde el CSV y pegarlos en la planilla"
		$TxtConnector = ("TEXT;" + $csv)
		# Atento al rango, acá le indicamos en que celda va a planchar la info del CSV
		$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("B8")) 
		$query = $worksheet.QueryTables.item($Connector.name)
		$query.TextFileOtherDelimiter = $delimiter
		$query.TextFileParseType  = 1
		$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
		#$query.AdjustColumnWidth = 1

		# Cerramos la query
		$query.Refresh()
		$query.Delete()
	
		Write-Host "[$(Get-Date -format $($dateFormat))] INFO: Estableciendo algunos parametros para la impresion (encabezado, pie, etc.)"
		# El header tiene que aparecer en todas las páginas que se vayan a imprimir, así que:
		$excel.ActiveWorkbook.ActiveSheet.PageSetup.PrintTitleRows = '$1:$3'

		# Lo mismo para el pie de página
		$worksheet.PageSetup.LeftFooter = "PROC-IT-MAR-0209"
		$worksheet.PageSetup.RightFooter = "Pagina &P"
	
		$worksheet.Columns("C").ColumnWidth=15
		
		# Guardamos el archivo generado con formato XLSX y cerramos el objeto COM
		Write-Host "[$(Get-Date -format $($dateFormat))] INFO: Planilla creada exitosamente, guardando y cerrando..."
		$workbook.SaveAs($xlsx,51)
		$excel.Quit()
	}
	exit 0
}
# Impresión en caso de errores
catch
	{
		Write-Exception
	}
# FIN