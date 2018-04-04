# requires -version 2
<#
.SYNOPSIS
	Script que convierte un reporte de QAD en formato CSV a formato Excel 2010+
	
.DESCRIPTION
	El script está pensado para mostrar la información de seguridad de los perfiles de usuarios configurados en QAD.Net 2008.1se
	
.INPUTS
	Por ahora ninguno, en algún momento haré que se pueda poner nombre de archivo de origen, destino, delimitador,
	si querés que te abra el Excel en vez de sólo generar el archivo y que de paso te apantalle y te haga unos mateykos
	
.OUTPUTS
	Write-Host para las excepciones
	
.NOTES
	Version:	1.6
	Author:		Santiago Platero
	Creation Date: 	04/04/2018
	Purpose/Change:	Commit inicial
	
.EXAMPLE
	> ./csv2xlsx.ps1
#>

#---------------------------------------------------------[Inicializaciones]--------------------------------------------------------

# Inicializaciones de variables, CSV de origen, XLSX de destino, path del logo y el delimitador que usa el CSV de origen
# Con paths relativos (¡salve!) no sé porque no funciona, cuestiones de M$ supongo
$csv = "W:\splatero\report.txt"
$xlsx = "W:\splatero\report.xlsx"
$imgPath = "W:\splatero\logo.png"
$delimiter = ";"

#----------------------------------------------------------[Declaraciones]----------------------------------------------------------

# Nada por aquí, nada por allá

#-----------------------------------------------------------[Ejecución]------------------------------------------------------------

# Control de errores
try {

	# Creación de objeto COM de M$ Excel
	$excel = New-Object -ComObject excel.application 
	$workbook = $excel.Workbooks.Add(1)
	$worksheet = $workbook.worksheets.Item(1)

	# Establecemos algunos tamaños específicos para columnas y filas
	$worksheet.Rows("1").RowHeight=74
	$worksheet.Rows("2").RowHeight=22
	$worksheet.Rows("4:6").RowHeight=28
	$worksheet.Rows("3").RowHeight=14
	$worksheet.Rows("7").RowHeight=14
	$worksheet.Columns("A").ColumnWidth=11
	$worksheet.Columns("B").ColumnWidth=9
	$worksheet.Columns("C").ColumnWidth=14
	$worksheet.Columns("D").ColumnWidth=33
	$worksheet.Columns("E").ColumnWidth=16
	
	# Alineaciones (todo centrado, como corresponde) y la fuente que se debe utilizar
	$worksheet.Columns("A:E").HorizontalAlignment = -4108 # ese negativo hermoso significa "centro", yo ya ni sé
	$worksheet.Columns("A:E").VerticalAlignment = -4108
	$worksheet.Columns("A:E").Font.Name = "Trebuchet MS"
	$worksheet.Columns("A:E").Font.Size = 10

	# Generamos el header del documento, combinamos las celdas para darle el formato que tiene que llevar
	$worksheet.Cells.item(1,1) = "Formulario de Perfiles de Usuario de QAD.Net"
	$worksheet.Cells.item(1,1).Font.Size = 14
	$mergecells1 = $worksheet.Range("A1:D1")
	$mergecells1.Select()
	$mergecells1.MergeCells = $true

	$worksheet.Cells.item(2,1) = "FOGL-IT-MAR-XXXX/01"
	$mergecells2 = $worksheet.Range("A2:E2")
	$mergecells2.Select()
	$mergecells2.MergeCells = $true

	$worksheet.Cells.item(4,1) = "SISTEMA: QAD.Net"
	$worksheet.Cells.item(4,1).Font.Bold = $true
	$mergecells3 = $worksheet.Range("A4:B4")
	$mergecells3.Select()
	$mergecells3.MergeCells = $true
	$mergecells3.Interior.ColorIndex = 37 # El 37 no está cargado, ese es el 38... Digo se refiere a un color celestito medio pedorro

	$worksheet.Cells.item(4,3) = "Firma / Aclaracion"
	$worksheet.Cells.item(4,3).Font.Bold = $true
	$mergecells6 = $worksheet.Range("C4:D4")
	$mergecells6.Select()
	$mergecells6.MergeCells = $true
	$mergecells6.Interior.ColorIndex = 15 # El 15 para en la esquina... No, es un gris claro

	$worksheet.Cells.item(4,5) = "Fecha"
	$worksheet.Cells.item(4,5).Font.Bold = $true
	$worksheet.Cells.item(4,5).Borders.LineStyle = 1 # El 1 es para una linea de borde continua, y para el arquero
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

	# Agregamos la imagen del logo, guarda que es un dolor de huevos los últimos cuatro parámetros: posición relativa (horizontal, luego vertical) y tamaño
	$img = $worksheet.Shapes.AddPicture($imgPath, $false, $true, 370, 11, 82, 50)

	# Agregamos algunos bordes para presentar un poco más prolija la info, el cuadro de firmantes, etc.
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

	# Armamos la query para planchar el CSV en lo que queda de planilla, el código lo choree de no sé donde, pero funciona
	$TxtConnector = ("TEXT;" + $csv)
	# Atento al rango, acá le indicamos en que celda va a planchar la info del CSV, quizá alguna versión futura sea un parámetro, si tengo ganas
	$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("B8")) 
	$query = $worksheet.QueryTables.item($Connector.name)
	$query.TextFileOtherDelimiter = $delimiter
	$query.TextFileParseType  = 1
	$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
	$query.AdjustColumnWidth = 1

	# Cerramos la query
	$query.Refresh()
	$query.Delete()

	# El dichoso header tiene que aparecer en todas las páginas que se vayan a imprimir, así que:
	$excel.ActiveWorkbook.ActiveSheet.PageSetup.PrintTitleRows = '$1:$3'

	# Lo mismo para el pie de página
	$worksheet.PageSetup.LeftFooter = "PROC-IT-MAR-XXXX"
	$worksheet.PageSetup.RightFooter = "Pagina &P"
}
# Impresión en caso de errores
catch
	{
		Write-Exception
	}

# Guardamos el archivo generado con formato XLSX y cerramos el objeto COM
$workbook.SaveAs($xlsx,51)
$excel.Quit()

# FIN