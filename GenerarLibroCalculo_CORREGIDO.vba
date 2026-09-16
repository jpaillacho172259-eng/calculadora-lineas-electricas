' ============================================================================
' MÓDULO: GenerarLibroCalculo - VERSIÓN CORREGIDA
' DESCRIPCIÓN: Macro completa F5 para crear calculadora de líneas eléctricas
' CORRIGE: Error 438 de compatibilidad con Merge
' ============================================================================

Option Explicit

Sub GenerarLibroCalculo()
    Dim wsLineasAereas As Worksheet
    Dim wsLineasSubterraneas As Worksheet
    Dim wsCentrosTransformacion As Worksheet
    Dim wsTablas As Worksheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Crear o limpiar hojas de trabajo
    Set wsLineasAereas = CrearHoja("LÍNEAS AÉREAS")
    Set wsLineasSubterraneas = CrearHoja("LÍNEAS SUBTERRÁNEAS")
    Set wsCentrosTransformacion = CrearHoja("CENTROS TRANSFORMACIÓN")
    Set wsTablas = CrearHoja("TABLAS REFERENCIA")
    
    ' Generar contenido de cada hoja
    GenerarHojaLineasAereas wsLineasAereas
    GenerarHojaLineasSubterraneas wsLineasSubterraneas
    GenerarHojaCentrosTransformacion wsCentrosTransformacion
    GenerarTablasReferencia wsTablas
    
    ' Crear portada
    CrearPortada
    
    ' Configuración final
    ThisWorkbook.Sheets(1).Activate
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "✓ Calculadora de líneas eléctricas generada correctamente" & vbCrLf & _
           "Presione F9 para recalcular todas las fórmulas", vbInformation, "Éxito"
    
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description & vbCrLf & "Línea: " & Erl, vbCritical, "Error"
End Sub

Function CrearHoja(nombreHoja As String) As Worksheet
    Dim ws As Worksheet
    Dim existe As Boolean
    
    existe = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    existe = (Err.Number = 0)
    On Error GoTo 0
    
    If existe Then
        ws.Cells.Delete
        Set CrearHoja = ws
    Else
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nombreHoja
        Set CrearHoja = ws
    End If
End Function

Sub CrearPortada()
    Dim ws As Worksheet
    Dim existe As Boolean
    
    existe = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PORTADA")
    existe = (Err.Number = 0)
    On Error GoTo 0
    
    If existe Then
        ws.Cells.Delete
    Else
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "PORTADA"
    End If
    
    With ws
        ' Formato de portada
        With .Range("A1:E15")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Size = 28
            .Font.Bold = True
            .Font.Color = RGB(0, 51, 102)
            .Interior.Color = RGB(200, 220, 240)
            .Value = "CALCULADORA DE LÍNEAS ELÉCTRICAS"
        End With
        
        With .Range("A17:E20")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 14
            .Font.Bold = True
            .Interior.Color = RGB(200, 220, 240)
            .Value = "Líneas Aéreas | Líneas Subterráneas | Centros de Transformación"
        End With
        
        With .Range("A25:E28")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 11
            .Interior.Color = RGB(200, 220, 240)
            .Value = "Conforme a REBT, RD 337/2014, RD 223/2008 y Normas UNE"
        End With
        
        With .Range("A35:E38")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 10
            .Font.Italic = True
            .Interior.Color = RGB(200, 220, 240)
            .Value = "Todas las celdas en AZUL requieren entrada manual"
        End With
        
        .Columns("A:E").ColumnWidth = 20
        .Rows("1:40").RowHeight = 25
    End With
End Sub

Sub GenerarHojaLineasAereas(ws As Worksheet)
    Dim fila As Long
    fila = 1
    
    With ws
        .Columns("A:F").ColumnWidth = 22
        
        ' Encabezado
        Call AgregarEncabezado(ws, fila, "CÁLCULO DE LÍNEAS AÉREAS")
        fila = fila + 3
        
        ' DATOS DE ENTRADA
        Call AgregarSeccion(ws, fila, "DATOS DE ENTRADA", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Tensión nominal (kV)", "A3", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Tipo de conductor", "A4", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Sección del conductor (mm²)", "A5", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Longitud de línea (km)", "A6", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Carga total (A)", "A7", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Factor de potencia", "A8", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Resistividad (Ω·mm²/m)", "A9", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Tipo de apoyo", "A10", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Zona de viento (Pa)", "A11", True)
        fila = fila + 1
        
        ' CÁLCULOS INTERMEDIOS
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "CÁLCULOS INTERMEDIOS", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Potencia (kW)", "=D7*D8*SQRT(3)/1000", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Resistencia total (Ω)", "=D9*D6*1000/D5", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Reactancia aproximada (Ω)", "=0.08*D6", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Impedancia total (Ω)", "=SQRT(D13^2+D14^2)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Corriente máxima teórica (A)", "=D7", False)
        fila = fila + 1
        
        ' RESULTADOS PRINCIPALES
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "RESULTADOS PRINCIPALES", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Caída de tensión (V)", "=D15*D13*SQRT(3)/1000", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Caída de tensión (%)", "=IF(D3=0,0,D17/(D3*1000)*100)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Pérdidas por efecto Joule (kW)", "=IF(D13=0,0,3*D15^2*D13/1000)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Rendimiento línea (%)", "=IF(D12=0,0,100-D19/D12*100)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Carga mecánica por viento (N/m)", "=IF(D11=0,0,D11*0.06*(SQRT(D5)/100)^1.5)", False)
        fila = fila + 1
        
        ' VERIFICACIONES NORMATIVAS
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "VERIFICACIONES NORMATIVAS (REBT/RD 337/2014)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "¿Caída tensión < 5%?", "=IF(D18<5,""✓ CUMPLE"",""✗ NO CUMPLE"")", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "¿Pérdidas < 10%?", "=IF(D20>90,""✓ CUMPLE"",""✗ NO CUMPLE"")", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Distancia mín. a suelo (m) [REBT]", "5.5", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Distancia mín. a edificios (m)", "=D3+2", False)
        
        ' Valores por defecto
        .Range("D4").Value = "Aluminio"
        .Range("D5").Value = 50
        .Range("D6").Value = 2
        .Range("D7").Value = 100
        .Range("D8").Value = 0.9
        .Range("D9").Value = 0.02826
        .Range("D10").Value = "Hormigón"
        .Range("D11").Value = 400
    End With
End Sub

Sub GenerarHojaLineasSubterraneas(ws As Worksheet)
    Dim fila As Long
    fila = 1
    
    With ws
        .Columns("A:F").ColumnWidth = 22
        
        Call AgregarEncabezado(ws, fila, "CÁLCULO DE LÍNEAS SUBTERRÁNEAS")
        fila = fila + 3
        
        ' DATOS DE ENTRADA
        Call AgregarSeccion(ws, fila, "DATOS DE ENTRADA", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Tensión nominal (kV)", "A3", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Tipo de cable (XLPE/PVC)", "A4", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Sección del cable (mm²)", "A5", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Longitud (km)", "A6", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Intensidad máxima (A)", "A7", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Factor de potencia", "A8", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Resistencia cable (Ω/km)", "A9", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Reactancia cable (Ω/km)", "A10", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Profundidad zanja (m)", "A11", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Temperatura terreno (°C)", "A12", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Factor agrupamiento cables", "A13", True)
        fila = fila + 1
        
        ' CÁLCULOS INTERMEDIOS
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "CÁLCULOS INTERMEDIOS", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Potencia (kW)", "=D7*D8*SQRT(3)/1000", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Resistencia total (Ω)", "=D9*D6", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Reactancia total (Ω)", "=D10*D6", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Impedancia (Ω)", "=SQRT(D16^2+D17^2)", False)
        fila = fila + 1
        
        ' RESULTADOS
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "RESULTADOS PRINCIPALES", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Caída de tensión (V)", "=IF(D7=0,0,D7*D18*SQRT(3)/1000)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Caída de tensión (%)", "=IF(D3=0,0,D20/(D3*1000)*100)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Pérdidas (kW)", "=IF(D7=0,0,3*D7^2*D16/1000)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Rendimiento (%)", "=IF(D15=0,0,100-IF(D22=0,0,D22/D15*100))", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Factor corrección temperatura", "=1+0.004*(D12-20)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Intensidad adm. corregida (A)", "=IF(D13=0,D7,D7*D24/D13)", False)
        fila = fila + 1
        
        ' VERIFICACIONES
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "VERIFICACIONES (REBT ITC-BT-19)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "¿Caída tensión < 3%?", "=IF(D21<3,""✓ CUMPLE"",""✗ NO CUMPLE"")", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "¿Intensidad adecuada?", "=IF(D25>=D7,""✓ CUMPLE"",""✗ NO CUMPLE"")", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Protección mecánica", "Banda + Tubo PVC", True)
        
        ' Valores por defecto
        .Range("D4").Value = "XLPE"
        .Range("D5").Value = 50
        .Range("D6").Value = 1.5
        .Range("D7").Value = 150
        .Range("D8").Value = 0.9
        .Range("D9").Value = 0.0144
        .Range("D10").Value = 0.073
        .Range("D11").Value = 1
        .Range("D12").Value = 25
        .Range("D13").Value = 0.9
    End With
End Sub

Sub GenerarHojaCentrosTransformacion(ws As Worksheet)
    Dim fila As Long
    fila = 1
    
    With ws
        .Columns("A:F").ColumnWidth = 22
        
        Call AgregarEncabezado(ws, fila, "CÁLCULO DE CENTROS DE TRANSFORMACIÓN")
        fila = fila + 3
        
        ' DATOS DE ENTRADA
        Call AgregarSeccion(ws, fila, "DATOS GENERALES DEL CT", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Identificación CT", "CT-001", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Ubicación", "A4", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Tipo de CT (Interior/Intemperie)", "A5", True)
        fila = fila + 1
        
        ' TRANSFORMADOR
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "DATOS DEL TRANSFORMADOR", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Potencia (kVA)", "A8", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Tensión primaria (kV)", "A9", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Tensión secundaria (V)", "A10", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Impedancia (Z%) ", "A11", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Pérdidas vacío (kW)", "A12", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Pérdidas carga (kW)", "A13", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Sistema puesta a tierra (TT/TN/IT)", "A14", True)
        fila = fila + 1
        
        ' CARGAS
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "DATOS DE CARGAS", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Potencia total instalada (kW)", "A17", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Factor de demanda (%)", "A18", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Factor de potencia medio", "A19", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Crecimiento futuro (%)", "A20", True)
        fila = fila + 1
        
        ' CÁLCULOS
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "CÁLCULOS Y RESULTADOS", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Demanda actual (kW)", "=IF(D18=0,0,D17*D18/100)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Demanda futura (kW)", "=D23*(1+IF(D20=0,0,D20/100))", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Factor utilización CT (%)", "=IF(D8=0,0,D24/D8*100)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Intensidad primaria (A)", "=IF(D9=0,0,D8*1000/(D9*SQRT(3)))", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Intensidad secundaria (A)", "=IF(D10=0,0,D8*1000/D10)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Caída de tensión (V)", "=IF(D10=0,0,(D11/100)*D10/SQRT(3))", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Resistencia puesta a tierra (Ω)", "A29", True)
        fila = fila + 1
        
        ' PROTECCIONES
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "PROTECCIONES Y SEGURIDAD", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Fusible primario (A)", "=D26*1.25", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Interruptor automático BT (A)", "=D27*1.25", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Sensibilidad diferencial (mA)", "A33", True)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "Tiempo desconexión máximo (s)", "=IF(D28=0,0,50/D28)", False)
        fila = fila + 1
        
        ' VERIFICACIONES
        fila = fila + 1
        Call AgregarSeccion(ws, fila, "VERIFICACIONES NORMATIVAS (RD 337/2014)", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "¿Utilización < 80%?", "=IF(D25<80,""✓ CUMPLE"",""✗ REVISAR"")", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "¿Caída tensión < 8%?", "=IF(D28/D10*100<8,""✓ CUMPLE"",""✗ NO CUMPLE"")", False)
        fila = fila + 1
        
        Call AgregarFilaDatos(ws, fila, "¿Resistencia tierra < 37 Ω?", "=IF(D30<37,""✓ CUMPLE"",""✗ NO CUMPLE"")", False)
        
        ' Valores por defecto
        .Range("D3").Value = "CT-001"
        .Range("D4").Value = "Calle Principal"
        .Range("D5").Value = "Interior"
        .Range("D8").Value = 250
        .Range("D9").Value = 20
        .Range("D10").Value = 400
        .Range("D11").Value = 5
        .Range("D12").Value = 1.2
        .Range("D13").Value = 4.5
        .Range("D14").Value = "TT"
        .Range("D17").Value = 150
        .Range("D18").Value = 80
        .Range("D19").Value = 0.9
        .Range("D20").Value = 20
        .Range("D30").Value = 20
        .Range("D33").Value = 30
    End With
End Sub

Sub GenerarTablasReferencia(ws As Worksheet)
    Dim fila As Long
    fila = 1
    
    With ws
        .Columns("A:F").ColumnWidth = 20
        
        Call AgregarEncabezado(ws, fila, "TABLAS DE REFERENCIA")
        fila = fila + 3
        
        ' TABLA 1: Conductores
        Call AgregarSeccion(ws, fila, "Tabla 1: Propiedades de Conductores", False)
        fila = fila + 1
        
        .Range("A" & fila).Value = "Tipo"
        .Range("B" & fila).Value = "Sección (mm²)"
        .Range("C" & fila).Value = "Resistencia (Ω/km)"
        .Range("D" & fila).Value = "I máx (A)"
        .Range("E" & fila).Value = "Peso (kg/km)"
        .Range("F" & fila).Value = "Precio (€/km)"
        
        With .Range("A" & fila & ":F" & fila)
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        fila = fila + 1
        .Range("A" & fila).Value = "Al 50"
        .Range("B" & fila).Value = 50
        .Range("C" & fila).Value = 0.5620
        .Range("D" & fila).Value = 120
        .Range("E" & fila).Value = 135
        .Range("F" & fila).Value = 250
        
        fila = fila + 1
        .Range("A" & fila).Value = "Al 95"
        .Range("B" & fila).Value = 95
        .Range("C" & fila).Value = 0.3
        .Range("D" & fila).Value = 180
        .Range("E" & fila).Value = 258
        .Range("F" & fila).Value = 380
        
        fila = fila + 1
        .Range("A" & fila).Value = "Al 150"
        .Range("B" & fila).Value = 150
        .Range("C" & fila).Value = 0.2
        .Range("D" & fila).Value = 240
        .Range("E" & fila).Value = 388
        .Range("F" & fila).Value = 520
        
        fila = fila + 2
        Call AgregarSeccion(ws, fila, "Tabla 2: Intensidades Máximas Admisibles (BT)", False)
        fila = fila + 1
        
        .Range("A" & fila).Value = "Sección (mm²)"
        .Range("B" & fila).Value = "Aéreo (A)"
        .Range("C" & fila).Value = "Subterráneo (A)"
        .Range("D" & fila).Value = "Tubo (A)"
        
        With .Range("A" & fila & ":D" & fila)
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(4, 33, 28, 21)
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(6, 43, 36, 27)
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(10, 60, 50, 38)
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(16, 80, 68, 51)
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(25, 106, 90, 68)
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(35, 135, 115, 87)
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(50, 160, 135, 102)
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(70, 200, 170, 128)
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(95, 245, 210, 158)
        
        fila = fila + 1
        .Range("A" & fila & ":D" & fila).Value = Array(120, 280, 245, 185)
        
        fila = fila + 3
        Call AgregarSeccion(ws, fila, "Tabla 3: Factores de Corrección por Temperatura", False)
        fila = fila + 1
        
        .Range("A" & fila).Value = "Temperatura (°C)"
        .Range("B" & fila).Value = "Factor"
        
        With .Range("A" & fila & ":B" & fila)
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        fila = fila + 1
        .Range("A" & fila & ":B" & fila).Value = Array(20, 1.00)
        
        fila = fila + 1
        .Range("A" & fila & ":B" & fila).Value = Array(25, 0.96)
        
        fila = fila + 1
        .Range("A" & fila & ":B" & fila).Value = Array(30, 0.91)
        
        fila = fila + 1
        .Range("A" & fila & ":B" & fila).Value = Array(35, 0.87)
        
        fila = fila + 1
        .Range("A" & fila & ":B" & fila).Value = Array(40, 0.82)
        
        fila = fila + 3
        Call AgregarSeccion(ws, fila, "Tabla 4: Caídas de Tensión Permitidas (REBT)", False)
        fila = fila + 1
        
        .Range("A" & fila).Value = "Tipo de instalación"
        .Range("B" & fila).Value = "Origen a cuadro"
        .Range("C" & fila).Value = "Cuadro a puntos"
        
        With .Range("A" & fila & ":C" & fila)
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        fila = fila + 1
        .Range("A" & fila).Value = "Líneas aéreas"
        .Range("B" & fila).Value = "5%"
        .Range("C" & fila).Value = "3%"
        
        fila = fila + 1
        .Range("A" & fila).Value = "Líneas subterráneas"
        .Range("B" & fila).Value = "3%"
        .Range("C" & fila).Value = "3%"
        
        fila = fila + 1
        .Range("A" & fila).Value = "Centros transformación"
        .Range("B" & fila).Value = "5%"
        .Range("C" & fila).Value = "5%"
    End With
End Sub

Sub AgregarEncabezado(ws As Worksheet, fila As Long, titulo As String)
    With ws.Range("A" & fila & ":F" & fila)
        .Value = titulo
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
End Sub

Sub AgregarSeccion(ws As Worksheet, fila As Long, titulo As String, Optional esEntrada As Boolean = False)
    With ws.Range("A" & fila & ":F" & fila)
        .Value = titulo
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        
        If esEntrada Then
            .Interior.Color = RGB(192, 0, 0)
        Else
            .Interior.Color = RGB(0, 102, 204)
        End If
        
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub AgregarFilaDatos(ws As Worksheet, fila As Long, etiqueta As String, valorFormula As String, Optional esEntrada As Boolean = False)
    Dim celdaValor As String
    celdaValor = "D" & fila
    
    ' Etiqueta en columnas A-C
    With ws.Range("A" & fila & ":C" & fila)
        .Value = etiqueta
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        
        If esEntrada Then
            .Interior.Color = RGB(173, 216, 230)
            .Font.Italic = True
            .Font.Color = RGB(0, 0, 128)
        End If
    End With
    
    ' Valor/Fórmula en columna D
    With ws.Range(celdaValor)
        .Value = valorFormula
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0.00"
        
        If esEntrada Then
            .Interior.Color = RGB(173, 216, 230)
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 128)
        Else
            .Interior.Color = RGB(240, 240, 240)
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlThin
        End If
    End With
    
    ' Columnas E-F (unidades)
    With ws.Range("E" & fila & ":F" & fila)
        .Font.Italic = True
        .Font.Size = 9
        .Interior.Color = RGB(255, 255, 255)
    End With
End Sub
