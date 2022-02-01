Attribute VB_Name = "Módulo1"
Sub Btn_ObtenerActivos()
' Codigo para llamar a la API
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Url = "https://investpyapi.herokuapp.com/stocks"
objHTTP.Open "GET", Url, False
objHTTP.setRequestHeader "Content-Type", "text/json"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (iPad; U; CPU OS 3_2_1 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Mobile/7B405"
objHTTP.send ("")

' Revisar que exista la hoja Activos
ComprobarHoja "Activos"

'Definir variables
Dim sht As Worksheet
Dim text As String
Dim result() As String
Dim row() As String

Set sht = ThisWorkbook.Worksheets("Activos")
sht.Cells.NumberFormat = "General"

Debug.Print objHTTP.Status
' Revisar si la respuesta fue exitosa
If objHTTP.Status = "200" Then
    ' Obtener el texto de la respuesta, aquí están los datos
    text = objHTTP.responseText
    ' Dividir el texto por salto de línea (\n)
    ' Esta función regresa un arreglo con las filas a insertar
    result = Split(text, "\n")
    For i = LBound(result()) To UBound(result()) - 1
         ' Dividir el texto de la fila por coma, ya que la respuesta está en formato CSV
         row = Split(result(i), ",")
         For j = LBound(row()) To UBound(row())
            ' i = fila y j = columna
            Cells(i + 1, j + 1) = row(j)
         Next j
    Next i
End If
End Sub

Sub Btn_ObtenerPreciosActivos()
    'Definir variables
    Dim inicio As String
    Dim fin As String
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("Activos")
    sht.range("I2").NumberFormat = "dd/mm/yyyy"
    sht.range("J2").NumberFormat = "dd/mm/yyyy"
    inicio = sht.range("I2").text
    fin = sht.range("J2").text
    ComprobarHoja "Historico"
    Dim activoIsin As String
    Dim activoNombre As String
    
    Dim lastRow As Long
    lastRow = sht.Cells(sht.Rows.count, 1).End(xlUp).row

    Dim range As range
    Set range = sht.range(sht.Rows(2), sht.Rows(lastRow)).SpecialCells(xlCellTypeVisible)
    Dim i As Integer
    Dim total As Integer
    total = 0
    For i = 2 To lastRow
        If sht.Cells(i, 1).EntireRow.Hidden Then
        Else
            activoIsin = sht.Cells(i, 5).Value
            activoNombre = sht.Cells(i, 3).Value
            ObtenerPreciosActivo activoIsin, activoNombre, inicio, fin
        End If
    Next
End Sub

Sub ObtenerPreciosActivo(id As String, activo As String, inicio As String, fin As String)
'Código para conectarse a la API para obtener los datos
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Url = "https://investpyapi.herokuapp.com/stocks/" + id + "?from_date=" + inicio + "&to_date=" + fin + "&columns=Date,Close"
objHTTP.Open "GET", Url, False
objHTTP.setRequestHeader "Content-Type", "text/json"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (iPad; U; CPU OS 3_2_1 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Mobile/7B405"
objHTTP.send ("")

'Definir variables
Dim sht As Worksheet
Dim lastColumn As Long
Dim text As String
Dim result() As String
Dim row() As String
Dim tempText As String

Set sht = ThisWorkbook.Worksheets("Historico")
'sht.Cells.NumberFormat = "General"

'Ctrl + Shift + End
lastColumn = sht.Cells(1, sht.Columns.count).End(xlToLeft).Column

If lastColumn > 1 Then lastColumn = lastColumn + 1

If objHTTP.Status = "200" Then
    text = objHTTP.responseText
    
    If Len(text) > 0 Then
        ' Quitar " del inicio
        tempText = Right(text, Len(text) - 1)
        result = Split(tempText, "\n")
        sht.Cells(1, lastColumn) = "Fecha"
        sht.Cells(1, lastColumn + 1) = activo
        For i = LBound(result()) + 1 To UBound(result()) - 1
             row = Split(result(i), ",")
             For j = LBound(row()) To UBound(row())
                'If j Is 0 Then sht.sht.Cells(i + 1, j + LastColumn).NumberFormat = "dd/mm/yyyy"
                sht.Cells(i + 1, j + lastColumn) = row(j)
             Next j
        Next i
    End If
End If
End Sub

Sub Btn_ProcesarActivos()
    'Definir variables
    Dim inicio As String
    Dim fin As String
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("Activos")
    sht.range("I2").NumberFormat = "dd/mm/yyyy"
    sht.range("J2").NumberFormat = "dd/mm/yyyy"
    inicio = sht.range("I2").text
    fin = sht.range("J2").text
    ComprobarHoja "HistoricoProcesado"
    
    ProcesarActivos inicio, fin
End Sub

Function ProcesarActivos(inicio As String, fin As String)
    ' El propósito de esta función es agrupar los valores de los activos por fechas
    ' De esta forma para todo activo se tendrá la misma cantidad de filas
    
    ' 1. Crear diccionario de fechas
    ' Este diccionario tendrá como llave la fecha y como valor tendrá otro diccionario
    ' El diccionario tendrá como llave la columna y como valor el cierre del activo para esa fecha
    ' Ejemplo:
    ' Donde: #C = Número de Columna y V = Valor
    '  Llave-Fecha   #C   V  #C   V  #C  V         #C  V
    ' {"01/01/2016":{ 1: 0.4, 2: 1.2, 3: "", ... , 29: 3 }}
    Dim shtHistoricoProcesado As Worksheet
    Set shtHistoricoProcesado = ThisWorkbook.Worksheets("HistoricoProcesado")
    Dim dicFechas As Object
    Set dicFechas = CreateObject("scripting.dictionary")

    ' 1.1 Procesar los valores de la hoja Historico
    Dim shtHistorico As Worksheet
    Set shtHistorico = ThisWorkbook.Worksheets("Historico")
    
    Dim lastColumn As Long
    Dim lastRow As Long
    
    Dim row As range
    ' Definir rangeActivo como Variant para usarlo como un arreglo de elementos
    Dim rangeActivo As Variant
    Dim rangeFechaActivo As Variant
    
    lastRow = shtHistorico.Cells(shtHistorico.Rows.count, 1).End(xlUp).row
    
    'Ctrl + Shift + End
    lastColumn = shtHistorico.Cells(1, shtHistorico.Columns.count).End(xlToLeft).Column
    
    Dim logNat As Double
    
    ' & sirve para concatener texto
    ' Aquí estamos uniendo A1:A con el número de lastRow, es decir la última fila
    ' Quedando de esta forma: A1:A26 (Ejemplo ilustrativo, puede variar)
    ' shtHistoricoProcesado.range("A1:A" & lastRow) = shtHistorico.range("A1:A" & lastRow).Value
    
    For i = 1 To lastColumn
        If i Mod 2 = 0 Then
            ' Insertar nombre del activo
            shtHistoricoProcesado.Cells(1, (i / 2) + 1) = shtHistorico.Cells(1, i)
            ' Obtener número de filas para la columna del activo
            lastRow = shtHistorico.Cells(shtHistorico.Rows.count, i).End(xlUp).row
            ' Obtener el rango de valores
            rangeFechaActivo = shtHistorico.range(shtHistorico.Cells(2, i - 1).address(0, 0), shtHistorico.Cells(lastRow, i - 1).address(0, 0))
            rangeActivo = shtHistorico.range(shtHistorico.Cells(2, i).address(0, 0), shtHistorico.Cells(lastRow, i).address(0, 0))
            ' UBound nos dice cuántos elementos tiene el rango
            For j = 1 To UBound(rangeActivo) - 1
                ' j controla las filas
                Dim col As Integer
                Dim fechaActivo As String
                col = (i / 2) + 1
                fechaActivo = rangeFechaActivo(j, 1)
                If dicFechas.Exists(fechaActivo) Then
                    dicFechas.Item(fechaActivo).Add col, rangeActivo(j, 1)
                Else
                    Dim dicColumnas As Object
                    Set dicColumnas = CreateObject("scripting.dictionary")
                    dicColumnas.Add col, rangeActivo(j, 1)
                    dicFechas.Add fechaActivo, dicColumnas
                End If
            Next
            
        End If
    Next
    
    Dim varKey As String
    Dim count As Integer
    Dim lastRowFechas As Long
    Dim rowActivoAnterior As Long
    Dim fechaValue As String
    count = 2
    shtHistoricoProcesado.Cells(1, 1) = "Fecha"
    ' Generar fechas
    GenerarFechas "HistoricoProcesado", dicFechas, inicio, fin
    lastRowFechas = shtHistoricoProcesado.Cells(shtHistoricoProcesado.Rows.count, 1).End(xlUp).row
    rangeFechaActivo = shtHistoricoProcesado.range(shtHistoricoProcesado.Cells(2, 1).address(0, 0), shtHistoricoProcesado.Cells(lastRowFechas, 1).address(0, 0))
    For i = 1 To lastRowFechas - 1
    'For Each varKey In dicFechas.Keys()
        varKey = rangeFechaActivo(i, 1)
        If dicFechas.Exists(varKey) Then
            'Dim dateValue As Date
            'dateValue = varKey
            'shtHistoricoProcesado.Cells(count, 1) = dateValue
            fechaValue = varKey
            Dim colKey As Variant
            For Each colKey In dicFechas.Item(varKey).Keys()
                Dim s As String
                ' Insertar el valor en la celda correspondiente
                shtHistoricoProcesado.Cells(count, colKey).Value = dicFechas.Item(varKey)(colKey)
                ' Obtener celda anterior con datos
                s = shtHistoricoProcesado.Cells(count, colKey).End(xlUp).address(0, 0)
                rowActivoAnterior = shtHistoricoProcesado.Cells(count, colKey).End(xlUp).row
                
                If IsEmpty(shtHistoricoProcesado.Cells(count - 1, colKey).Value) Then
                    shtHistoricoProcesado.range(shtHistoricoProcesado.Cells(rowActivoAnterior + 1, colKey), shtHistoricoProcesado.Cells(count - 1, colKey)).Value = dicFechas.Item(varKey)(colKey)
                End If
            Next
        End If
        
        count = count + 1
    Next
    
    ' Rellenar datos finales
    Dim rangeActProcesados As range
    Dim lastColActProcesados As Integer
    Dim addressActivo As String
    Dim colActivo As Integer
    
    colActivo = 2
    lastColActProcesados = shtHistoricoProcesado.Cells(1, 2).End(xlToRight).Column
    Set rangeActProcesados = shtHistoricoProcesado.range(shtHistoricoProcesado.Cells(1, 2).address(0, 0), shtHistoricoProcesado.Cells(1, lastColActProcesados).address(0, 0))
    For Each activo In rangeActProcesados
        addressActivo = activo.End(xlDown).address(0, 0)
        shtHistoricoProcesado.range(shtHistoricoProcesado.Cells(lastRowFechas, colActivo), shtHistoricoProcesado.range(addressActivo).Cells) = shtHistoricoProcesado.range(addressActivo).Value
        colActivo = colActivo + 1
    Next
End Function

Function GenerarFechas(shtName As String, dicFechas As Object, inicio As String, fin As String)
    ' Generar fechas a partir de los valores de inicio y fin comprobando el diccionario de fechas
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(shtName)
    Dim FirstDate As Date
    Dim LastDate As Date
    Dim NextDate As Date
    Dim fechaValue As String
    ' r = variable para llevar el control de las filas
    Dim r As Long, count As Long
    FirstDate = inicio
    LastDate = fin
    ' Empezar en la segunda fila
    r = 2
    count = 2
    Do
        fechaValue = FirstDate
        If dicFechas.Exists(fechaValue) Then
            ' Insertar fecha
            sht.Cells(count, 1).Value = FirstDate
            count = count + 1
        End If

        ' Obtener siguiente fecha
        FirstDate = FirstDate + 1
        ' Siguiente fila
        r = r + 1
    ' Hacer lo anterior hasta que la fecha de inicio sea igual a la final
    Loop Until FirstDate = LastDate
    ' Insertar fecha final
    If dicFechas.Exists(FirstDate) Then
        ' Insertar fecha
        sht.Cells(count, 1) = FirstDate
    End If
End Function


Function ComprobarHoja(hoja As String)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(hoja)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ws.Activate
        Exit Function
    End If
    
    Set ws = Sheets.Add(Before:=Sheets(1))
    ws.Name = hoja

End Function

Sub Btn_Rentabilidad()
ObtenerRentabilidadMedia
'ObtenerEstadisticaDescriptiva
End Sub

Sub Btn_RentabilidadProcesado()
ObtenerRentabilidadMediaProcesado
ObtenerRentabilidadMedia
ObtenerEstadisticaDescriptiva
ObtenerPromedioRentabilidad
End Sub

Function ObtenerRentabilidadMedia()
'Definir variables
Dim sht As Worksheet
Set sht = ThisWorkbook.Worksheets("Historico")
Dim lastColumn As Long
Dim lastRow As Long

Dim row As range
' Definir rangeActivo como Variant para usarlo como un arreglo de elementos
Dim rangeActivo As Variant

lastRow = sht.Cells(sht.Rows.count, 1).End(xlUp).row

'Ctrl + Shift + End
lastColumn = sht.Cells(1, sht.Columns.count).End(xlToLeft).Column

ComprobarHoja "Rentabilidad"
Dim shtRenta As Worksheet
Set shtRenta = ThisWorkbook.Worksheets("Rentabilidad")

Dim logNat As Double

' & sirve para concatener texto
' Aquí estamos uniendo A1:A con el número de lastRow, es decir la última fila
' Quedando de esta forma: A1:A26 (Ejemplo ilustrativo, puede variar)
shtRenta.range("A1:A" & lastRow) = sht.range("A1:A" & lastRow).Value

For i = 2 To lastColumn
    If i Mod 2 = 0 Then
        ' Insertar nombre del activo
        shtRenta.Cells(1, (i / 2) + 1) = sht.Cells(1, i)
        ' Obtener número de filas para la columna del activo
        lastRow = sht.Cells(sht.Rows.count, i).End(xlUp).row
        ' Obtener el rango de valores
        rangeActivo = sht.range(Cells(2, i).address(0, 0), Cells(lastRow, i).address(0, 0))
        ' Aplicar logaritmo natural a los valores de cierre para cada activo, menos el último
        ' UBound nos dice cuántos elementos tiene el rango
        For j = 1 To UBound(rangeActivo) - 1
            ' Calcular logaritmo natural del elemento actual entre el anterior
            ' j controla las filas
            ' Ejemplo:
            ' j = 2 entonces es el actual
            ' j-1 = 1 entonces es el anterior
            logNat = Application.WorksheetFunction.Ln(rangeActivo(j + 1, 1) / rangeActivo(j, 1))
            ' Insertar el valor empezando desde la segunda fila
            shtRenta.Cells(j + 1, (i / 2) + 1) = 100 * logNat
        Next
    End If
Next

End Function

Function ObtenerRentabilidadMediaProcesado()
'Definir variables
Dim sht As Worksheet
Set sht = ThisWorkbook.Worksheets("HistoricoProcesado")
Dim lastColumn As Long
Dim lastRow As Long

Dim row As range
' Definir rangeActivo como Variant para usarlo como un arreglo de elementos
Dim rangeActivo As Variant

lastRow = sht.Cells(sht.Rows.count, 1).End(xlUp).row

'Ctrl + Shift + End
lastColumn = sht.Cells(1, sht.Columns.count).End(xlToLeft).Column

ComprobarHoja "Rentabilidad Media"
Dim shtRenta As Worksheet
Set shtRenta = ThisWorkbook.Worksheets("Rentabilidad Media")

Dim logNat As Double

' & sirve para concatener texto
' Aquí estamos uniendo A1:A con el número de lastRow, es decir la última fila
' Quedando de esta forma: A1:A26 (Ejemplo ilustrativo, puede variar)
shtRenta.range("A1:A" & lastRow) = sht.range("A1:A" & lastRow).Value

For i = 2 To lastColumn
    ' Insertar nombre del activo
    shtRenta.Cells(1, i) = sht.Cells(1, i)
    ' Obtener número de filas para la columna del activo
    lastRow = sht.Cells(2, i).End(xlDown).row
    ' Obtener el rango de valores
    rangeActivo = sht.range(sht.Cells(2, i), sht.Cells(lastRow, i))
    ' Aplicar logaritmo natural a los valores de cierre para cada activo, menos el último
    ' UBound nos dice cuántos elementos tiene el rango
    For j = 1 To UBound(rangeActivo) - 1
        ' Calcular logaritmo natural del elemento actual entre el anterior
        ' j controla las filas
        ' Ejemplo:
        ' j = 1 entonces es el actual
        ' j+1 = 1 entonces es el siguiente
        logNat = Application.WorksheetFunction.Ln(rangeActivo(j + 1, 1) / rangeActivo(j, 1))
        ' Insertar el valor empezando desde la tercera fila
        shtRenta.Cells(j + 2, i) = 100 * logNat
        Next
Next

shtRenta.Rows(2).EntireRow.Delete
End Function


Sub Btn_Estadistica()
ObtenerEstadisticaDescriptiva
End Sub

Function ObtenerEstadisticaDescriptiva()
    'Definir variables
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("Rentabilidad Media")
    Dim lastColumn As Long
    Dim lastRow As Long
    
    Dim rowActivos As range, cellActivo As range
    ' Definir rangeActivo como Variant para usarlo como un arreglo de elementos
    Dim rangeActivo As Variant
    
    lastRow = sht.Cells(sht.Rows.count, 1).End(xlUp).row
    
    'Ctrl + Shift + End
    lastColumn = sht.Cells(1, sht.Columns.count).End(xlToLeft).Column
    
    Dim logNat As Double
    
    ' Obtener los nombres de los activos
    ' Cells(fila, columna)
    Dim lastActivo As range
    Set lastActivo = sht.Cells(1, 2).End(xlToRight)
    Dim l As String
    l = "B1:" & lastActivo.address(0, 0)
    Set rowActivos = sht.range(l)
    sht.range(sht.Cells(1, lastActivo.Column + 3).address(0, 0)).Resize(1, rowActivos.Columns.count) = rowActivos.Cells.Value
    ' Insertar etiquetas
    sht.Cells(2, lastActivo.Column + 2) = "Rentabilidad Media (R)"
    sht.Cells(3, lastActivo.Column + 2) = "Varianza"
    sht.Cells(4, lastActivo.Column + 2) = "Desviación Estandar"
    
    Dim row As Integer, col As Integer
    
    For Each cellActivo In rowActivos.Cells
        Dim tempRange As range
        Dim s As String
        row = cellActivo.row + 1
        col = cellActivo.Column
        Set tempRange = sht.range(sht.Cells(row, col), sht.Cells(row, col).End(xlDown))
        s = sht.Cells(row, col).End(xlDown).address(0, 0)
        sht.Cells(2, lastActivo.Column + col + 1) = Application.WorksheetFunction.Average(tempRange)
        sht.Cells(3, lastActivo.Column + col + 1) = Application.WorksheetFunction.Var_P(tempRange)
        sht.Cells(4, lastActivo.Column + col + 1) = Application.WorksheetFunction.StDev_P(tempRange)
    Next
    
    sht.UsedRange.Columns.AutoFit
End Function


Function ObtenerPromedioRentabilidad()
    ' Esta función obtiene el promedio de las rentabilidades de la hoja Rentabilidad
    ' Y los inserta en la hoja Rentabilidad Media
    'Definir variables
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("Rentabilidad")
    Dim shtMedia As Worksheet
    Set shtMedia = ThisWorkbook.Worksheets("Rentabilidad Media")
    Dim lastColumn As Long
    Dim lastRow As Long
    
    Dim rowActivos As range, cellActivo As range
    ' Definir rangeActivo como Variant para usarlo como un arreglo de elementos
    Dim rangeActivo As Variant
    
    lastRow = sht.Cells(sht.Rows.count, 1).End(xlUp).row
    
    'Ctrl + Shift + End
    lastColumn = sht.Cells(1, sht.Columns.count).End(xlToLeft).Column
    
    Dim logNat As Double
    
    ' Obtener los nombres de los activos
    ' Cells(fila, columna)
    Dim lastActivo As range
    Set lastActivo = sht.Cells(1, 2).End(xlToRight)
    Dim l As String
    l = "B1:" & lastActivo.address(0, 0)
    Set rowActivos = sht.range(l)
    'sht.range(sht.Cells(1, lastActivo.Column + 3).address(0, 0)).Resize(1, rowActivos.Columns.count) = rowActivos.Cells.Value
    ' Insertar etiquetas
    'sht.Cells(2, lastActivo.Column + 2) = "Rentabilidad Media (R)"
    'sht.Cells(3, lastActivo.Column + 2) = "Varianza"
    'sht.Cells(4, lastActivo.Column + 2) = "Desviación Estandar"
    
    Dim row As Integer, col As Integer
    
    For Each cellActivo In rowActivos.Cells
        Dim tempRange As range
        Dim s As String
        row = cellActivo.row + 1
        col = cellActivo.Column
        Set tempRange = sht.range(sht.Cells(row, col), sht.Cells(row, col).End(xlDown))
        s = sht.Cells(row, col).End(xlDown).address(0, 0)
        shtMedia.Cells(2, lastActivo.Column + col + 1) = Application.WorksheetFunction.Average(tempRange)
        'sht.Cells(3, lastActivo.Column + col + 1) = Application.WorksheetFunction.Var_P(tempRange)
        'sht.Cells(4, lastActivo.Column + col + 1) = Application.WorksheetFunction.StDev_P(tempRange)
    Next
    
    sht.UsedRange.Columns.AutoFit
End Function

Sub Btn_ObtenerMatrizVarCov()
ObtenerMatrizVarCov
End Sub
Function ObtenerMatrizVarCov()
    'Definir variables
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("Rentabilidad Media")
    Dim lastColumn As Long
    Dim lastRow As Long
    
    Dim rowActivos As range, cellActivo As range
    ' Definir rangeActivo como Variant para usarlo como un arreglo de elementos
    Dim rangeActivo As Variant
    
    lastRow = sht.Cells(sht.Rows.count, 1).End(xlUp).row
    
    'Ctrl + Shift + End
    lastColumn = sht.Cells(1, sht.Columns.count).End(xlToLeft).Column
    
    Dim logNat As Double
    
    ' Obtener los nombres de los activos
    ' Cells(fila, columna)
    Dim lastActivo As range
    Set lastActivo = sht.Cells(1, 2).End(xlToRight)
    Dim addressActivos As String
    addressActivos = "B1:" & lastActivo.address(0, 0)
    Set rowActivos = sht.range(addressActivos)
    ' Insertar nombres de los activos
    sht.range(sht.Cells(8, lastActivo.Column + 3).address(0, 0)).Resize(1, rowActivos.Columns.count) = rowActivos.Cells.Value
    sht.range(sht.Cells(9, lastActivo.Column + 2).address(0, 0)).Resize(rowActivos.Columns.count, 1) = Application.WorksheetFunction.Transpose(rowActivos.Cells.Value)
    
    Dim row As Integer, col As Integer
    Dim count As Integer
    count = 1
    For Each cellActivo In rowActivos.Cells
        Dim tempRange As range, tempRange2 As range
        Dim s As String
        row = cellActivo.row + 1
        col = cellActivo.Column
        ' Datos del activo 1
        Set tempRange = sht.range(sht.Cells(row, col), sht.Cells(row, col).End(xlDown))
        s = sht.Cells(row, col).End(xlDown).address(0, 0)
        
        ' Insertar covarianza del activo 1
        sht.Cells(8 + count, lastActivo.Column + col + 1) = sht.Cells(3, lastActivo.Column + col + 1)
        ' j controla la columna
        For j = 1 To rowActivos.Cells.count - count
            ' Datos del activo 2
            Set tempRange2 = sht.range(sht.Cells(row, col + j), sht.Cells(row, col + j).End(xlDown))
            s = sht.Cells(row, col + j).End(xlDown).address(0, 0)
            Dim res As Double
            res = Application.WorksheetFunction.Covariance_P(tempRange, tempRange2)
            ' Insertar en fila, incrementar por cada activo 1
            sht.Cells(8 + count, (lastActivo.Column + col) + j + 1) = res
            ' Insertar en columna
            sht.Cells(8 + count + j, (lastActivo.Column + 2) + count) = res
        Next
        
        count = count + 1
        
    Next
    
    ' Insetar matriz de pesos
    Dim addressRentabilidad As String
    Dim addressMatrizPesos As String
    Dim addressMatrizCov As String
    Dim addressTotal As String
    Dim addressRentEsperada As String
    Dim addressRiesgo As String
    
    addressRentabilidad = sht.range(sht.Cells(2, lastActivo.Column + 3), sht.Cells(2, lastActivo.Column + 3 + rowActivos.Columns.count - 1)).address(0, 0)
    addressMatrizPesos = sht.range(sht.Cells(8 + count + 4, lastActivo.Column + 3), sht.Cells(8 + count + 4, lastActivo.Column + 3 + rowActivos.Columns.count - 1)).address(0, 0)
    addressMatrizCov = sht.range(sht.Cells(9, lastActivo.Column + 3), sht.Cells(8 + count - 1, lastActivo.Column + 3 + rowActivos.Columns.count - 1)).address(0, 0)
    
    sht.range(sht.Cells(8 + count + 3, lastActivo.Column + 3).address(0, 0)).Resize(1, rowActivos.Columns.count) = rowActivos.Cells.Value
    sht.range(addressMatrizPesos) = 1 / rowActivos.Cells.count
    sht.Cells(8 + count + 3, lastActivo.Column + 2) = "Total"
    addressTotal = sht.Cells(8 + count + 4, lastActivo.Column + 2).address(0, 0)
    sht.range(addressTotal).Formula2 = "=SUM(" & addressMatrizPesos & ")"

    ' Insertar Función para calcular el Rendimiento Esperado
    sht.Cells(8 + count + 6, lastActivo.Column + 2) = "Rent. Esperada E(R)"
    addressRentEsperada = sht.Cells(8 + count + 7, lastActivo.Column + 2).address(0, 0)
    sht.range(addressRentEsperada).Formula2 = "=MMULT(" & addressRentabilidad & ", TRANSPOSE(" & addressMatrizPesos & "))"
    'sht.Cells(8 + count + 7, lastActivo.Column + 2) = Application.WorksheetFunction.MMult(sht.range(addressRentabilidad), Application.WorksheetFunction.Transpose(sht.range(addressMatrizPesos)))
    
    ' Insertar formula para el Riesgo
    sht.Cells(8 + count + 6, lastActivo.Column + 3) = "Riesgo"
    addressRiesgo = sht.Cells(8 + count + 7, lastActivo.Column + 3).address(0, 0)
    sht.range(addressRiesgo).Formula2 = "=SQRT(MMULT(" & addressMatrizPesos & ", MMULT(" & addressMatrizCov & ", TRANSPOSE(" & addressMatrizPesos & "))))"
    'sht.Cells(8 + count + 7, lastActivo.Column + 3) = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MMult(sht.range(addressMatrizPesos), sht.range(addressMatrizCov)), Application.WorksheetFunction.Transpose(sht.range(addressMatrizPesos)))
    
    ComprobarHoja "Resultados"
    Dim shtRes As Worksheet
    Set shtRes = ThisWorkbook.Worksheets("Resultados")
    
    ' Insertar etiquetas
    shtRes.range("A1") = "Sharp"
    shtRes.range("B1") = "Riesgo"
    shtRes.range("C1") = "Rentabilidad"
    shtRes.range("D1").Resize(1, rowActivos.Columns.count) = rowActivos.Cells.Value
    
    Dim n As Integer
    n = 11
    Dim min As Variant, max As Variant, stepValue As Variant, stepCount As Variant
    ' Minimizar Riesgo
    min = EjecutarSolver(addressRiesgo, 2, addressMatrizPesos, addressTotal, RefMatrizPesos:=addressMatrizPesos, RefRentabilidad:=addressRentEsperada, RefRiesgo:=addressRiesgo)
    ' Sharpe
    shtRes.range("A2").Formula2 = "=(C2-0.000049)/B2"
    
    ' Reiniciar la matriz de pesos
    sht.range(addressMatrizPesos) = 1 / rowActivos.Cells.count
    
    ' Maximizar Rentabilidad
    ' La función ejecutar solver regresa el valor del riesgo
    max = EjecutarSolver(addressRentEsperada, 1, addressMatrizPesos, addressTotal, RefMatrizPesos:=addressMatrizPesos, RefRentabilidad:=addressRentEsperada, RefRiesgo:=addressRiesgo)
    ' Sharpe
    shtRes.range("A2").Formula2 = "=(C2-0.000049)/B2"
    
    ' Reiniciar la matriz de pesos
    sht.range(addressMatrizPesos) = 1 / rowActivos.Cells.count
    
    ' Calcular los saltos entre carteras
    stepValue = (max - min) / (n - 1)
    stepCount = min + stepValue
    
    For i = 1 To n - 2
        ' Ejecutar solver n veces para generar carteras
        ' Los argumentos requeridos deben de ir en el orden definido en la función
        ' Los argumentos opcionales deben ir al final, el orden no importa
        EjecutarSolver addressRiesgo, 2, addressMatrizPesos, addressTotal, _
        RefEqualTo:=addressRentEsperada, EqualToValue:=stepCount, SaveRow:=3, _
        RefMatrizPesos:=addressMatrizPesos, RefRentabilidad:=addressRentEsperada, RefRiesgo:=addressRiesgo
        
        stepCount = stepCount + stepValue
        ' Reiniciar la matriz de pesos
        sht.range(addressMatrizPesos) = 1 / rowActivos.Cells.count
        
        ' Sharpe
        shtRes.range("A3").Formula2 = "=(C3-0.000049)/B3"
    Next
    
    Dim rangeSharp As range
    Set rangeSharp = shtRes.range("A2:A" & n + 1)
    Dim maxSharp As Variant, rowSharp As Variant
    Dim addressMaxSharpValue As String
    
    maxSharp = 0
    For Each rowSharp In rangeSharp
        If rowSharp > maxSharp Then
            maxSharp = rowSharp.Value
            addressMaxSharpValue = rowSharp.address(0, 0)
        End If
    Next
    
    shtRes.range(addressMaxSharpValue).EntireRow.Interior.Color = RGB(255, 255, 0)
    
    sht.UsedRange.Columns.AutoFit
End Function

Function EjecutarSolver(SetCell As String, MaxMinVal As Integer, ByChange As String, TotalCell As String, _
Optional ByVal ValueOf As Variant, Optional ByVal RefMatrizPesos As Variant, _
Optional RefRentabilidad As Variant, Optional RefRiesgo As Variant, Optional SaveRow As Variant, _
Optional EqualToValue As Variant, Optional RefEqualTo As Variant) As Variant
    'Definir variables
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("Rentabilidad Media")
    sht.Activate
    
    SolverReset
    
    ' https://docs.microsoft.com/en-us/office/vba/excel/concepts/functions/solverok-function
    ' SolverOk permite definir el objetivo a resolver
    ' SetCell: referencia de la celda objetivo
    ' MaxMinVal: opciones Max (1), Min (2), Val (3)
    ' ValueOf: Opcional, sólo aplica para la opción de Val (3), es el valor deseado
    ' ByChange: rango de celdas que deben cambiar para dar una solución a la celda objetivo
    SolverOk SetCell:=sht.range(SetCell), MaxMinVal:=MaxMinVal, ValueOf:=ValueOf, ByChange:=sht.range(ByChange), Engine:=1
    
    ' https://docs.microsoft.com/en-us/office/vba/excel/concepts/functions/solveroptions-function
    ' SolverOptions permite definir las configuraciones a usar
    ' AssumeNonNeg: se asume como valor mínimo 0
    SolverOptions AssumeNonNeg:=True
    
    ' https://docs.microsoft.com/en-us/office/vba/excel/concepts/functions/solveradd-function
    ' SolverAdd permite agregar restricciones al modelo
    ' CellRef: referencia a la celda donde se aplicará la restricción
    ' Relation: opción para el valor lógico que se debe cumplir
    ' 1: Menor que (<=)
    ' 2: Igual a (=)
    ' 3: Mayor que (>=)
    ' 4: Debe tener valores enteros
    ' 5: Debe tener entre 0 y 1
    ' 6: Debe tener valores diferentes y enteros
    SolverAdd CellRef:=sht.range(TotalCell), Relation:=2, FormulaText:=1
    If IsMissing(RefEqualTo) = False And IsMissing(EqualToValue) = False Then
        Dim val As String
        val = EqualToValue
        SolverAdd CellRef:=sht.range(RefEqualTo), Relation:=2, FormulaText:=val
    End If
    ' SolverSolve ejecuta todo lo anterior e inicia el proceso para resolver el modelo
    SolverSolve UserFinish:=True
    
    ' Guardar la solución
    ComprobarHoja "Resultados"
    Dim shtRes As Worksheet
    Set shtRes = ThisWorkbook.Worksheets("Resultados")
    Dim row As Integer
    If IsMissing(SaveRow) = True Then
        row = 2
    Else
        row = SaveRow
    End If
    
    shtRes.Cells(row, 1).EntireRow.Insert
    shtRes.Cells(row, 2) = sht.range(RefRiesgo).Value
    shtRes.Cells(row, 3) = sht.range(RefRentabilidad).Value
    shtRes.range(shtRes.Cells(row, 4).address(0, 0)).Resize(1, sht.range(RefMatrizPesos).Columns.count) = sht.range(RefMatrizPesos).Cells.Value
    
    ' SolverFinish permite manipular el resultado
    ' KeepFinal: opción (1) los valores cambiados son conservados, (2) los valores cambiados son descartados y se restauran los anteriores
    SolverFinish KeepFinal:=2
    
    EjecutarSolver = shtRes.Cells(row, 3).Value
End Function

