'Muestra todos los buques en el mapa.
Private Sub BtnDebug_Click()
    For Each shp In ActiveSheet.Shapes
        If InStr(1, shp.Name, "pos_") <> 0 Then
            shp.Visible = IIf(shp.Visible = True, False, True)
        End If
    Next shp
End Sub
'IMPORTANTE:
'Fernández Mourelle, Aaron Manuel [BECARIO] (FE):
'Dejo esto por aquí también en caso de que pudiera ocurrir algún problema.
'En caso de haber borrado alguna de las formas que son 100% necesarias:
'- Los dos buques que usamos de plantilla para generar el resto.
'- Los cuadrados que guardan las posiciones absolutas para pegar los barcos nuevos.
'O en el caso de haber añadido a la tabla alguna fila (poco probable, ya que se creó en base a una teórica capacidad máxima).
'
'Dejo aquí el sistema de nombres para dichas formas.
'Para los dos buques:
'- predShapeRef_Right: buque que apunta a la derecha.
'- predShapeRef_Left: el inverso.
'Para los cuadrados que guardan posiciones relativas:
'- pos_C + fila que guarda el nombre del muelle. P.Ej. pos_C10:
'En caso de cambiar de columna, habrá que modificar el script.
'Véase primera parte del For Each:
'       positionShapeName = "pos_C" & element.Range(1).row       <- Cambiar C por la columna.
'       newShapeName = "barco_C" & element.Range(1).row         <- Cambiar C por la columna.
'
'Alt + F10 para mostrar el menú de selección y editar los nombres rápidamente.
Private Sub btnNuevoActualizar_Click()
'Variables -- Ease of use.
'--------------------------------------------------------------------------------------------------------------------------------
    'Variables de forma.
    'Formas de posición.
    Dim positionShape As Shape          'La forma que guarda la posición y rotación.
    Dim positionShapeName As String     'Almacena el nombre de la forma encargada de guardar las coordenadas y su rotación.
    '-----------------------------------------------------------------------------------------------------------------------
    'Formas base (la forma en sí y su orientación).
    Dim templateShapeName As String     'Referencia a la forma de la que se aplican copias.
    '-----------------------------------------------------------------------------------------------------------------------
    'Forma nueva.
    Dim newShape As Shape               'La forma que se creará en la iteración.
    Dim newShapeName As String          'El nombre.
    '-----------------------------------------------------------------------------------------------------------------------
    'Auxiliares.
    Dim auxShapeName As String          'Nombre temporal para hacer copy/paste y borrar la forma.
        auxShapeName = "auxShapeName"
    '-----------------------------------------------------------------------------------------------------------------------
    'Variables de datos.
    'Información del barco (tanto para booleanos como para atributos).
    Dim shipName As String              'Guarda el nombre del buque (Valor de columna).
    Dim shipDepartureDate As String     'Fecha de salida del barco (segundo criterio para mostrar la forma).
    Dim shipNameLength As Integer       'Largo del nombre para escalar el largo del barco.
    Dim predShape As String             'Referencia la forma del barco a copiar.
    Dim isMilitary As Boolean           'Referencia para el color.
    Dim isReversed As Boolean           'Referencia a la orientación (si hay texto o no).
'--------------------------------------------------------------------------------------------------------------------------------
    For Each element In ActiveSheet.ListObjects("tblDatos4").ListRows
        'Recojo el nombre de la forma de posición
        positionShapeName = "pos_C" & element.Range(1).row
        'Asigno un nombre para la futura forma.
        newShapeName = "barco_C" & element.Range(1).row
        'Recojo el nombre del barco.
        shipName = Trim(element.Range(2).Value)
        'Recojo el largo del nombre.
        shipNameLength = Len(shipName)
        'Recojo la fecha de salida.
        shipDepartureDate = Trim(element.Range(7).Value)
        'Recojo ya si es militar o no.
        isMilitary = InStr(1, UCase(Trim(element.Range(4).Value)), "MILIT", vbTextCompare) > 0
        'Recojo su orientación (si hay que invertirla o no).
        isReversed = Trim(element.Range(8).Value) <> ""
        'Recojemos ya el nombre de la forma a copiar.
        templateShapeName = IIf(isReversed, "predShapeRef_Left", "predShapeRef_Right")
    '-----------------------------------------------------------------------------------------------------------------------
        'Hacemos una distinción y copia de plantilla.
        ActiveSheet.Shapes(templateShapeName).Name = auxShapeName
        With ActiveSheet.Shapes(auxShapeName)
            .Copy
            .Name = templateShapeName
        End With
        'Pegamos la copia y seleccionamos en algún lado para evitar problemas con la forma.
        With ActiveSheet
            .Paste
            .Range("A45").Select
        End With
        'Recogemos las formas
        Set positionShape = ActiveSheet.Shapes(positionShapeName)
        '
        With ActiveSheet.Shapes(auxShapeName)
            'Aplicamos tamaño y rotación.
            .rotation = positionShape.rotation
            'Estos dos pueden no ser necesarios.
                .Width = ActiveSheet.Shapes(templateShapeName).Width
                .Height = ActiveSheet.Shapes(templateShapeName).Height
            '-----------------------------------
            'Dato dentro de la forma.
                .TextFrame.Characters.Text = shipName
                With ActiveSheet.Shapes(auxShapeName).TextFrame2
                    .AutoSize = msoAutoSizeTextToFitShape
                    .WordWrap = msoFalse
                End With
            'Controlamos si mostrarla o no.
            .Visible = Not shipName = "" And shipDepartureDate = ""
            'Color.                     Militar ?       Gris        :       Rojo
            .Fill.ForeColor.RGB = IIf(isMilitary, RGB(178, 178, 178), RGB(236, 202, 201))
            'Aquí estiramos la forma. El criterio fue más o menos a ojo.
            If (shipNameLength > 10) Then
                .Width = ActiveSheet.Shapes(templateShapeName).Width + ((shipNameLength - 10) * 3.3)
            End If
        '----------------------------------------------------------------------
            'Asignamos su posición
            .Top = positionShape.Top - positionShape.Height
            .Left = positionShape.Left - .Height * 2
        End With
        On Error Resume Next
        Set newShape = ActiveSheet.Shapes(newShapeName)
        If Not newShape Is Nothing Then
            newShape.Delete
        End If
        ActiveSheet.Shapes(auxShapeName).Name = newShapeName
        '--------------------------------------------------------------------------------------------------------------
        'Debug for value testing.
        '   Debug.Print templateShape.Name
        '   Debug.Print newShapeName & "  " & shipName & "  " & positionShapeName & "  " & shipDepartureDate & "  rev:" & isReversed
        '--------------------------------------------------------------------------------------------------------------
    Next element
End Sub
