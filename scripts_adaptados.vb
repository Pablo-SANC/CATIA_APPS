Imports INFITF
Imports MECMOD

Module scripts_adaptados
    'Language="VBSCRIPT"
    Dim CATIA = CreateObject("CATIA.Application")

    '-----------------------------------------------
    '-----------------------------------------------

    Sub Renombrar_puntos_sele()
        On Error Resume Next
        Dim Elemseleccionado As Selection
        Elemseleccionado = CATIA.ActiveDocument.Selection
        If Elemseleccionado.Count2 = 0 Then
            MsgBox("NO HAY NADA SELECCIONADO")
        End If
        Dim Identificacion_punto = "a"
        Dim I = 0
        Dim ind

        Identificacion_punto = InputBox("IDENTIFICACION DE LOS PUNTOS/ ", "Modificar Identificacion de puntos", Identificacion_punto)
        I = InputBox("NUMERO SECUENCIA ", "Índices", I)

        ind = I

        For j = 1 To Elemseleccionado.Count2

            If ind < 10 Then
                Elemseleccionado.Item(j).Value.Name = Identificacion_punto & "0" & ind
            Else
                Elemseleccionado.Item(j).Value.Name = Identificacion_punto & ind
            End If

            ind = ind + 1

        Next


    End Sub
    '-----------------------------------------------
    '-----------------------------------------------


    Sub Proyectar_Puntos_a_sketch()

        'On Error Resume Next
        Dim Seleccion
        Dim Resultado
        Dim sketch1
        Dim Elemseleccionados
        Dim ProyectarPunto
        Dim reference1
        Dim geometricElements1

        Seleccion = CATIA.ActiveDocument.Selection
        Seleccion.Clear()

        '  Seleccion del sketch
        Dim ElementoTipo(0)
        ElementoTipo(0) = "Sketch"
        Resultado = Seleccion.SelectElement2(ElementoTipo, "SELECCIONA EL SKETCH", True)
        If Resultado = "Cancel" Then
            Exit Sub
        End If
        sketch1 = Seleccion.Item(1).Value
        Dim factory2D1 As Factory2D
        factory2D1 = sketch1.OpenEdition()
        Seleccion.Clear()

        Elemseleccionados = CATIA.ActiveDocument.Selection
        ProyectarPunto = True
        Do While ProyectarPunto
            Selinter(Elemseleccionados, Resultado)
            If Resultado = "Cancel" Then
                sketch1.CloseEdition()
                ProyectarPunto = False
                Exit Sub
            End If
            reference1 = Elemseleccionados.Item(1).Value
            geometricElements1 = factory2D1.CreateProjection(reference1)
            geometricElements1.Name = Elemseleccionados.Item(1).Value.Name
            Elemseleccionados.Clear()
        Loop

    End Sub

    ' ---------------------------------------------
    ' ---------------------------------------------
    'Subrutina para Elegir INTERACTIVAMENTE PUNTOS

    Sub Selinter(Elemseleccionados, Resultado)
        Dim ElementoTipo
        ReDim ElementoTipo(0)
        ElementoTipo(0) = "Point"
        Resultado = Elemseleccionados.SelectElement2(ElementoTipo, "SELECCIONA UN PUNTO", True)

    End Sub
    '-----------------------------------------------
    '-----------------------------------------------

    Sub Links_spanish()

        Dim ThisWindow
        Dim drawingDocument1
        Dim Selection
        Dim box
        Dim a
        Dim D
        Dim Status
        Dim Vista
        Dim drawingSheet1
        Dim partDocument1
        Dim drawingViews1
        Dim drawingView1
        Dim part1
        Dim myLinks
        Dim Pregunta

        Dim InputObjectType(0)

        MsgBox("Para el funcionamiento correcto de este macro se debe tener abierto en CATIA únicamente el Part/Product y la hoja de plano a relinkar. " + Chr(10) + "Abra primero el Part/Product y después el plano" + Chr(10) + "Después de presionar OK por favor seleccione la vista a relinkar")

        a = 1

        Do While a = 1

            ThisWindow = CATIA.Windows

            If ThisWindow.Count < 2 Then
                box = MsgBox("No se abrieron los dos archivos" + Chr(10) + "El macro no se ejecuta")
                Exit Sub
            End If

            If ThisWindow.Count > 2 Then
                box = MsgBox("Hay más de dos archivos abiertos" + Chr(10) + "El macro no se ejecuta")
                Exit Sub
            End If

            ThisWindow = CATIA.Windows.Item(2)
            ThisWindow.Activate()

            drawingDocument1 = CATIA.ActiveDocument

            If TypeName(drawingDocument1) <> "DrawingDocument" Then
                box = MsgBox("Archivos cargados incorrectamente" + Chr(10) + "Por favor abra primero el Part/Product y después el plano y ejecute la macro de nuevo", vbInformation, "ERROR")
                Exit Sub
            End If

            D = drawingDocument1.Sheets

            Selection = drawingDocument1.Selection

            Selection.Clear()

            InputObjectType(0) = "AnyObject"

            Status = Selection.SelectElement2(InputObjectType, "Por favor seleccione la vista a relinkar", True)
            If (Status = "cancel") Then Exit Sub


            Vista = Selection.Item(1).Value


            drawingSheet1 = D.Item(1)

            drawingViews1 = drawingSheet1.Views

            drawingView1 = drawingViews1.Item(Vista.name)

            drawingView1.Activate()

            myLinks = drawingView1.GenerativeLinks
            myLinks.RemoveAllLinks()

            MsgBox("Todos los enlaces de la vista " & Vista.name & " han sido eliminados con éxito")

            ThisWindow = CATIA.Windows.Item(1)
            ThisWindow.Activate()

            partDocument1 = CATIA.ActiveDocument

            If TypeName(partDocument1) <> "ProductDocument" And TypeName(partDocument1) <> "PartDocument" Then
                box = MsgBox("El primer archivo no es ningún Part o Product" + Chr(10) + "Por favor abra primero el Part/Product y después el plano y ejecute la macro de nuevo", vbInformation, "ERROR")
                Exit Sub
            End If

            If TypeName(partDocument1) = "PartDocument" Then

                part1 = partDocument1.Part

                myLinks.AddLink(part1)

                ThisWindow = CATIA.Windows.Item(2)
                ThisWindow.Activate()

                MsgBox("Nuevo Link añadido. Macro finalizada con éxito!" + Chr(10) + "La vista todavía debe ser actualizada manualmente")

            End If

            If TypeName(partDocument1) = "ProductDocument" Then

                part1 = partDocument1.Product

                myLinks.AddLink(part1)

                ThisWindow = CATIA.Windows.Item(2)
                ThisWindow.Activate()

                MsgBox("Nuevo Link añadido. Macro finalizada con éxito!" + Chr(10) + "La vista todavía debe ser actualizada manualmente")

            End If

            Pregunta = MsgBox("¿Desea relinkar alguna vista más? Seleccione la vista a continuación de OK.", vbYesNo)
            If Pregunta = vbNo Then Exit Sub Else 

        Loop


    End Sub
    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Escape Then Exit Sub
    End Sub



End Module
