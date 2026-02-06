Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub GenerarPDFsPorCorreo()

    Dim celdaInicioBase As Range, celdaInicioLista As Range
    Dim wsBase As Worksheet, wsLista As Worksheet
    Dim colBaseCorreo As Long, colBaseVIN As Long
    Dim filaBaseInicio As Long, filaBaseFin As Long
    Dim filaListaInicio As Long, filaListaFin As Long
    Dim colListaCorreo As Long

    '=========================
    ' 1) Pedir inicio de base (correo + VIN a la derecha)
    '=========================
    On Error Resume Next
    Set celdaInicioBase = Application.InputBox( _
        Prompt:="Selecciona la celda donde INICIA la base (primer CORREO)." & vbCrLf & _
                "El VIN debe estar en la columna inmediata a la derecha.", _
        Title:="Inicio base correos/VINs", Type:=8)
    On Error GoTo 0
    If celdaInicioBase Is Nothing Then Exit Sub

    Set wsBase = celdaInicioBase.Worksheet
    colBaseCorreo = celdaInicioBase.Column
    colBaseVIN = colBaseCorreo + 1
    filaBaseInicio = celdaInicioBase.Row
    filaBaseFin = wsBase.Cells(wsBase.Rows.Count, colBaseCorreo).End(xlUp).Row
    If filaBaseFin < filaBaseInicio Then
        MsgBox "No se encontraron datos en la base.", vbExclamation
        Exit Sub
    End If

    '=========================
    ' 2) Pedir inicio de lista de correos (una columna)
    '=========================
    On Error Resume Next
    Set celdaInicioLista = Application.InputBox( _
        Prompt:="Selecciona la celda donde INICIA la lista de correos (una columna).", _
        Title:="Inicio lista de correos", Type:=8)
    On Error GoTo 0
    If celdaInicioLista Is Nothing Then Exit Sub

    Set wsLista = celdaInicioLista.Worksheet
    colListaCorreo = celdaInicioLista.Column
    filaListaInicio = celdaInicioLista.Row
    filaListaFin = wsLista.Cells(wsLista.Rows.Count, colListaCorreo).End(xlUp).Row
    If filaListaFin < filaListaInicio Then
        MsgBox "No se encontraron correos en la lista.", vbExclamation
        Exit Sub
    End If

    '=========================
    ' 3) Pedir CAMPAÑA y FECHA
    '=========================
    Dim campania As String, fechaTexto As String
    campania = Trim$(CStr(InputBox("Escribe el nombre de la CAMPAÑA", "Campaña")))
    If campania = "" Then
        MsgBox "Campaña vacía. Cancelado.", vbExclamation
        Exit Sub
    End If

    fechaTexto = Trim$(CStr(InputBox("Escribe la FECHA tal como debe aparecer (ej. 9 de diciembre de 2024):", "Fecha")))
    If fechaTexto = "" Then
        MsgBox "Fecha vacía. Cancelado.", vbExclamation
        Exit Sub
    End If

    '=========================
    ' 4) Elegir plantilla Word y carpeta destino
    '=========================
    Dim rutaPlantilla As String, carpetaSalida As String
    rutaPlantilla = ElegirArchivoWord("Selecciona tu plantilla Word")
    If rutaPlantilla = "" Then Exit Sub

    carpetaSalida = ElegirCarpeta("Selecciona la carpeta donde se guardarán los PDFs")
    If carpetaSalida = "" Then Exit Sub

    '=========================
    ' 5) Construir diccionario: correo(minúsculas) -> Collection VINs
    '=========================
    Dim dictVINsPorCorreo As Object
    Set dictVINsPorCorreo = CreateObject("Scripting.Dictionary")

    Dim r As Long, correo As String, vin As String
    Dim colVINs As Collection

    For r = filaBaseInicio To filaBaseFin
        correo = LCase$(Trim$(CStr(wsBase.Cells(r, colBaseCorreo).Value)))
        vin = Trim$(CStr(wsBase.Cells(r, colBaseVIN).Value))

        If correo <> "" And vin <> "" Then
            If Not dictVINsPorCorreo.Exists(correo) Then
                Set colVINs = New Collection
                dictVINsPorCorreo.Add correo, colVINs
            End If
            If Not ExisteEnCollection(dictVINsPorCorreo(correo), vin) Then
                dictVINsPorCorreo(correo).Add vin
            End If
        End If
    Next r

    '=========================
    ' 6) Abrir Word (una sola vez) y generar 1 PDF por correo
    '=========================
    Dim wordApp As Object, wordDoc As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    Dim correoInteres As String
    Dim vinsCorreo As Collection
    Dim totalVINs As Long
    Dim nombreLocal As String
    Dim rutaPDF As String

    For r = filaListaInicio To filaListaFin

        correoInteres = LCase$(Trim$(CStr(wsLista.Cells(r, colListaCorreo).Value)))
        If correoInteres = "" Then GoTo Siguiente

        ' Abrir doc desde plantilla
        Set wordDoc = wordApp.Documents.Add(Template:=rutaPlantilla, NewTemplate:=False)

        ' Reemplazar CAMPANIA / FECHA (ajusta los marcadores si cambian)
        ReemplazarTodo wordDoc, "<<CAMPANIA>>", campania
        ReemplazarTodo wordDoc, "<<FECHA>>", fechaTexto

        ' Obtener VINs del correo
        If dictVINsPorCorreo.Exists(correoInteres) Then
            Set vinsCorreo = dictVINsPorCorreo(correoInteres)
        Else
            Set vinsCorreo = Nothing
        End If

        ' Insertar tabla en <<VIN_TABLA>>
        InsertarTablaVINsEnMarcador wordDoc, "<<VIN_TABLA>>", vinsCorreo

        ' Nombre PDF: #VINs_localPart  (ej: 5_correoEjemplo.pdf)
        totalVINs = IIf(vinsCorreo Is Nothing, 0, vinsCorreo.Count)
        nombreLocal = ObtenerLocalPart(correoInteres) ' lo de antes del @
        nombreLocal = NombreSeguroArchivo(nombreLocal)

        rutaPDF = carpetaSalida & "\" & CStr(totalVINs) & "_" & nombreLocal & ".pdf"
        rutaPDF = RutaDisponible(rutaPDF)

        ' Exportar a PDF
        wordDoc.ExportAsFixedFormat OutputFileName:=rutaPDF, ExportFormat:=17 ' 17=PDF
        wordDoc.Close False

Siguiente:
    Next r

    wordApp.Quit
    Set wordApp = Nothing

    MsgBox "Listo. PDFs generados en:" & vbCrLf & carpetaSalida, vbInformation

End Sub

'=========================
' Helpers Word
'=========================
Private Sub ReemplazarTodo(ByVal doc As Object, ByVal buscar As String, ByVal reemplazo As String)
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = buscar
        .Replacement.Text = reemplazo
        .Wrap = 1 ' wdFindContinue
        .Execute Replace:=2 ' wdReplaceAll
    End With
End Sub

Private Sub InsertarTablaVINsEnMarcador(ByVal doc As Object, ByVal marcador As String, ByVal vins As Collection)

    Dim rng As Object
    Set rng = doc.Content

    With rng.Find
        .ClearFormatting
        .Text = marcador
        .Wrap = 0 ' wdFindStop
        If .Execute = False Then Exit Sub
    End With

    rng.Text = "" ' borrar marcador

    If vins Is Nothing Or vins.Count = 0 Then
        rng.InsertAfter "SIN VINs ENCONTRADOS"
        Exit Sub
    End If

    Dim total As Long, filas As Long
    total = vins.Count
    filas = (total + 2) \ 3 ' ceil(total/3)

    Dim tbl As Object
    Set tbl = doc.Tables.Add(rng, filas, 3)

    ' Sin bordes
    tbl.Borders.Enable = False

    ' >>> Que Word ajuste a la ventana útil
    tbl.AllowAutoFit = True
    tbl.AutoFitBehavior 2 ' wdAutoFitWindow

    ' Formato legible
    With tbl.Range
        .Font.Name = "Consolas" ' o "Courier New"
        .Font.Size = 10
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.LineSpacingRule = 0 ' single
        .ParagraphFormat.Alignment = 0 ' left
    End With

    ' Padding y altura para que no se monten
    tbl.TopPadding = 2
    tbl.BottomPadding = 2
    tbl.LeftPadding = 4
    tbl.RightPadding = 4

    tbl.Rows.AllowBreakAcrossPages = False
    tbl.Range.Cells.VerticalAlignment = 0 ' wdCellAlignVerticalTop

    Dim f As Long, c As Long, idx As Long
    idx = 1

    For c = 1 To 3
        For f = 1 To filas
            If idx <= total Then
                tbl.Cell(f, c).Range.Text = Trim$(CStr(vins(idx)))
                idx = idx + 1
            Else
                tbl.Cell(f, c).Range.Text = ""
            End If
        Next f
    Next c

    ' Quitar el párrafo extra que Word mete al final de cada celda (limpia el look)
    Dim celda As Object
    For Each celda In tbl.Range.Cells
        celda.Range.Paragraphs.Last.Range.Text = Replace(celda.Range.Paragraphs.Last.Range.Text, Chr(13) & Chr(7), Chr(7))
    Next celda

End Sub


'=========================
' Helpers Excel / Archivos
'=========================
Private Function ExisteEnCollection(ByVal col As Collection, ByVal texto As String) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If CStr(col(i)) = texto Then
            ExisteEnCollection = True
            Exit Function
        End If
    Next i
    ExisteEnCollection = False
End Function

Private Function ElegirArchivoWord(ByVal titulo As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = titulo
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Word", "*.docx;*.dotx"
        If .Show <> -1 Then
            ElegirArchivoWord = ""
        Else
            ElegirArchivoWord = .SelectedItems(1)
        End If
    End With
End Function

Private Function ElegirCarpeta(ByVal titulo As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = titulo
        If .Show <> -1 Then
            ElegirCarpeta = ""
        Else
            ElegirCarpeta = .SelectedItems(1)
        End If
    End With
End Function

Private Function ObtenerLocalPart(ByVal correo As String) As String
    Dim partes() As String
    partes = Split(correo, "@")
    If UBound(partes) >= 0 Then
        ObtenerLocalPart = partes(0)
    Else
        ObtenerLocalPart = correo
    End If
End Function

Private Function NombreSeguroArchivo(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, ":", "_")
    t = Replace(t, "*", "_")
    t = Replace(t, "?", "_")
    t = Replace(t, """", "_")
    t = Replace(t, "<", "_")
    t = Replace(t, ">", "_")
    t = Replace(t, "|", "_")
    t = Replace(t, " ", "_")
    NombreSeguroArchivo = t
End Function

Private Function RutaDisponible(ByVal ruta As String) As String
    If Dir(ruta) = "" Then
        RutaDisponible = ruta
        Exit Function
    End If

    Dim base As String, ext As String, i As Long, p As Long
    p = InStrRev(ruta, ".")
    ext = Mid$(ruta, p)
    base = Left$(ruta, p - 1)
    i = 1

    Do While Dir(base & "(" & i & ")" & ext) <> ""
        i = i + 1
    Loop

    RutaDisponible = base & "(" & i & ")" & ext
End Function


