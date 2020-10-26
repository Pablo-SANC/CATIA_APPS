Public Class Form1


    Public Sub Main()


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Process.Start("\\Brjcasa-fs001\dep\Ingenieria\OFICINA_TECNICA\Engineering Design\MACRO CATIA\COMPROBADAS\Convert_CATDrawingToDWG.catscript")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Process.Start("\\Brjcasa-fs001\dep\Ingenieria\OFICINA_TECNICA\Engineering Design\MACRO CATIA\COMPROBADAS\Convert_CATDrawingToPDF.catscript")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Process.Start("\\Brjcasa-fs001\dep\Ingenieria\OFICINA_TECNICA\Engineering Design\MACRO CATIA\COMPROBADAS\Convert_CATDrawingToDXF.catscript")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Process.Start("\\Brjcasa-fs001\dep\Ingenieria\OFICINA_TECNICA\Engineering Design\MACRO CATIA\COMPROBADAS\MIN Material.exe")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Process.Start("\\Brjcasa-fs001\dep\Ingenieria\OFICINA_TECNICA\Engineering Design\MACRO CATIA\COMPROBADAS\Tablas_de_Empilado.exe")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Process.Start("\\Brjcasa-fs001\dep\Ingenieria\OFICINA_TECNICA\Engineering Design\MACRO CATIA\COMPROBADAS\FreeRenaming_V5R16.exe")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Process.Start("\\Brjcasa-fs001\dep\Ingenieria\OFICINA_TECNICA\Engineering Design\MACRO CATIA\COMPROBADAS\RENOMBRAR_ANTENAS.catscript")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Process.Start("\\Brjcasa-fs001\dep\Ingenieria\OFICINA_TECNICA\Engineering Design\MACRO CATIA\AIRBUS\RENOMBRAR_INSTANCIAS.catscript")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        MessageBox.Show("•En caso de fallos o errores con Pablo Sánchez" & vbCrLf & " e-mail: pablo.sanchez_delosbueis@airbus.com", "Contacto", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, 0, "mailto:pablo.sanchez_delosbueis@airbus.com") '0 is default otherwise use MessageBoxOptions Enum
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Call Proyectar_Puntos_a_sketch()
        'Process.Start("PROYECTAR_PUNTOS_A_SKETCH.CATScript")
    End Sub


    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Call Renombrar_puntos_sele()
        'Process.Start("RENOMBRAR_PUNTOS_SELECCIONADOS.CATScript")
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Call Links_spanish()
    End Sub


    Private Sub Button1_Mouseover(sender As Object, e As EventArgs) Handles Button1.MouseHover
        ToolTip1.Show("Convierte lotes de planos CATDrawing a DXF" & vbCrLf & vbCrLf & "•Elige una carpeta donde están los CATDrawing." & vbCrLf & "•Espera a que CATIA termine" & vbCrLf & "•Los planos en DXF se generan en la misma ubicación.", Button1)
    End Sub
    Private Sub Button2_Mouseover(sender As Object, e As EventArgs) Handles Button2.MouseHover
        ToolTip2.Show("MIN Material extrae en un excel el minimo material necesario para piezas de mecanizado." & vbCrLf & vbCrLf & "•Abre un CATPart o un CATProduct en modo diseño." & vbCrLf & "•Pulsa Analizar y espera a que salga un mensaje." & vbCrLf & vbCrLf & "•Solo analiza CATParts que tengan el PartBody como Define In Work Object." & vbCrLf & "•Omite la mayoría de Standards.", Button2)
    End Sub
    Private Sub Button3_Mouseover(sender As Object, e As EventArgs) Handles Button3.MouseHover
        ToolTip3.Show("Convierte lotes de planos CATDrawing a DWG" & vbCrLf & vbCrLf & "•Elige una carpeta donde están los CATDrawing." & vbCrLf & "•Espera a que CATIA termine" & vbCrLf & "•Los planos en DWG se generan en la misma ubicación.", Button3)
    End Sub
    Private Sub Button4_Mouseover(sender As Object, e As EventArgs) Handles Button4.MouseHover
        ToolTip4.Show("Convierte lotes de planos CATDrawing a PDF" & vbCrLf & vbCrLf & "•Elige una carpeta donde están los CATDrawing." & vbCrLf & "•Espera a que CATIA termine" & vbCrLf & "•Los planos en PDF se generan en la misma ubicación.", Button4)
    End Sub
    Private Sub Button5_Mouseover(sender As Object, e As EventArgs) Handles Button5.MouseHover
        ToolTip5.Show("Genera una tabla de empilado en un CATDrawing" & vbCrLf & vbCrLf & "•Te permite seleccionar el tipo de empilado." & vbCrLf & "•Sigue las preguntas que te formula." & vbCrLf & "•Espera a que se genere.", Button5)
    End Sub
    Private Sub Button6_Mouseover(sender As Object, e As EventArgs) Handles Button6.MouseHover
        ToolTip6.Show("Cambia el nombre de la instancia de los Part Numbers" & vbCrLf & vbCrLf & "•Selecciona uno o varios elementos y se renombrarán secuenciales." & vbCrLf & "•Sigue las preguntas que te formula." & vbCrLf & "•Espera a que se genere el cambio de instancia.", Button6)
    End Sub
    Private Sub Button7_Mouseover(sender As Object, e As EventArgs) Handles Button7.MouseHover
        ToolTip7.Show("Renombra los PartNumbers en sustitución de los caracteres que se quiera." & vbCrLf & vbCrLf & "•Te preguntará que nombre se quiere sustituir." & vbCrLf & "•saltará el número de cambios que se van a realizar." & vbCrLf & "•Espera a se cambien los nombres.", Button7)
    End Sub
    Private Sub Button8_Mouseover(sender As Object, e As EventArgs) Handles Button8.MouseHover
        ToolTip8.Show("Renombra las instancias al PartNumber seguido de ' . ' y la secuencia que le toque." & vbCrLf & vbCrLf & "•El modelo que esté abierto tiene que estar en modo diseño." & vbCrLf & "•Espera a se cambien los nombres de las instancias.", Button8)
    End Sub

    Private Sub Button10_Mouseover(sender As Object, e As EventArgs) Handles Button10.MouseHover
        ToolTip10.Show("Proyecta puntos del 3D a un sketch renombrando la proyección." & vbCrLf & vbCrLf & "•Primero selecciona el sketch donde se quieren proyectar los puntos." & vbCrLf & "•Selecciona los puntos uno a uno en el 3D." & vbCrLf & "•Para dejar de utilizar el comando, pulsar la tecla escape (ESC).", Button10)
    End Sub
    Private Sub Button11_Mouseover(sender As Object, e As EventArgs) Handles Button11.MouseHover
        ToolTip11.Show("Renombra los elementos seleccionados dentro de un CATPart o un CATDrawing. (Lineas, Puntos, Sketches, Geometrical Sets, Bodies, Vistas...)" & vbCrLf & vbCrLf & "•Primero selecciona los elementos que quiere renombrar." & vbCrLf & "•Introduce el nombre para los puntos." & vbCrLf & "•Continua asignando el primer valor de conteo.", Button11)
    End Sub
    Private Sub Button12_Mouseover(sender As Object, e As EventArgs) Handles Button12.MouseHover
        ToolTip12.Show("Actualiza los Links de los CATDrawings." & vbCrLf & vbCrLf & "•Primero abre el modelo 3D(CATPart o CATProduct)." & vbCrLf & "•Abre el plano (CATDrawing)." & vbCrLf & "•Revista que en el CATDrawing no hay dos vistas con nombres iguales de las que quieras volver a enlazar." & vbCrLf & "   En el caso de que exista una o más de una, renombralas para que no se llamen igual." & vbCrLf & "   (Puedes utilizar la macro de 'Renombrar Elementos' si quieres secuenciarlos)." & vbCrLf & "•La macro te preguntará si has acabado o quieres repetir la operación para más vistas.", Button12)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Color.FromArgb(19, 42, 96)
        Me.AutoSize = True
        TabPage1.BackColor = Color.FromArgb(19, 42, 96)
        TabPage1.Text = "CATIA 3D"
        TabPage2.Text = "CATIADrawing"
        'Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)
    End Sub


End Class
