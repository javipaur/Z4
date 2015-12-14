Option Strict Off
Imports System.IO
Imports System.Array
Imports System
Imports System.Text
Imports Microsoft.VisualBasic

Public Class Form1
    Dim directorio As String = My.Application.Info.DirectoryPath
    Dim FileNumber As Integer = FreeFile()
    'Dim codificacion = System.Text.Encoding.UTF8

    'Inicializamos las variables globales
    Dim Ldirectorio As New ArrayList()
    'Guardamos el path del directorio raiz y le añadimos $transformationene y modificamos la fila especificada
    Dim Ldirectoriotrafo As New ArrayList()
    'Dim directorioInicial As String = "L:\Practicas Mercedes Benz Vitoria\Proyecto Visual\PruebasZ4"
    'Dim directorioDestino As String = "L:\Practicas Mercedes Benz Vitoria\Proyecto Visual\PruebasZ4Nuevo"

    Dim directorioInicial As String = "C:\DatosW7\Z4"
    Dim directorioDestino As String = "C:\DatosW7\Z4nuevo"

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        borrarLog()

        ' Inicializa los formatos . , y fechas
        ' Inicializa los formatos . , y fechas
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("es-CO")
        System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = "yyyy/MM/dd"
        System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator = "."
        System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyGroupSeparator = ","
        System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator = "."
        System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator = ","

        'Realizar copia del directorio Z4 en Z4Nuevo
        CopiarZ4(directorioInicial, directorioDestino)

        'Modificar archivo *.csv y trafos
        ModificarCsvTrafos()

        'Modificar el Pie de todos los trafos del directorio
        ModificarPieTrafoDirectorioCmompeto()

        'Generar Log
        generarLog()
    End Sub

    'Copiar Directorio Z4
    Private Sub CopiarZ4(ByVal directorioInicial As String, ByVal directorioDestino As String)

        Dim fso, f As Object
        fso = CreateObject("Scripting.FileSystemObject")
        Dim s As String
        Try
            filesListBox.Items.Add("** Copiar Directorios **")

            'Informacion directorio inicial
            f = fso.Getfolder(directorioInicial)
            s = (String.Format("{0:N2}", CInt(f.size) / 1024) & "KB")
            filesListBox.Items.Add("  Origen:   " & directorioInicial & " ==> " & s)

            'Realizamos la copia
            My.Computer.FileSystem.CopyDirectory(directorioInicial, directorioDestino, True)

            'Informacion resultados de la copia
            f = fso.Getfolder(directorioDestino)
            s = (String.Format("{0:N2}", CInt(f.size) / 1024) & "KB")
            filesListBox.Items.Add("  Destino: " & directorioDestino & " ==> " & s)

            filesListBox.Items.Add("** Copia Realizada **")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    'Buscar Csv Y Trafos
    Private Sub BuscarArchivosCsv(ByVal directorioDestino As String)
        Dim directorioAModificarCsv As String
        Dim ultimaBarra As Integer
        Try
            Dim dir As New System.IO.DirectoryInfo(directorioDestino)
            'Buscamos Archivos en todo el directorio que tengan extension .csv
            Dim fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories)
            Dim fileQuery = From file In fileList _
                             Where file.Extension =
                             ".csv" _
                             Order By file.Name _
                             Select file

            For Each file In fileQuery
                directorioAModificarCsv = file.FullName
                'Comprobamos si existen los trafos1,2 y 3
                ultimaBarra = directorioAModificarCsv.LastIndexOf("\")
                'MsgBox("Ultima barra " & ultimaBarra)
                Dim dirTrafo01 As String = directorioAModificarCsv.Substring(0, ultimaBarra) + "\$Transformationen\TraFo1.txt"
                Dim dirTrafo02 As String = directorioAModificarCsv.Substring(0, ultimaBarra) + "\$Transformationen\TraFo2.txt"
                Dim dirTrafo03 As String = directorioAModificarCsv.Substring(0, ultimaBarra) + "\$Transformationen\TraFo3.txt"
                'MsgBox(dirTrafo01 & " - " & My.Computer.FileSystem.FileExists(dirTrafo01) & vbCrLf & dirTrafo02 & " - " & My.Computer.FileSystem.FileExists(dirTrafo02) & vbCrLf & dirTrafo03 & " - " & My.Computer.FileSystem.FileExists(dirTrafo03))
                If My.Computer.FileSystem.FileExists(dirTrafo01) And My.Computer.FileSystem.FileExists(dirTrafo02) And My.Computer.FileSystem.FileExists(dirTrafo03) Then
                    'En Ldirectorio global guardamos el path completo del Csv 
                    Ldirectorio.Add(directorioAModificarCsv)
                End If
            Next
        Catch ex As Exception
            MsgBox("Descripcion del error:" & _
                   ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub
    'ModificarCsvTrafos
    Private Sub ModificarCsvTrafos()
        Dim directorioTrafo As String
        'Buscar dentro del arblo de Z4Nuevo subdirectorios que contengan la carpeta $transformationene con los tres archivos trafo.txt
        BuscarArchivosCsv(directorioDestino)
        'Modificar csv
        Dim directorioCSV As String
        'Por cada directorio del csv
        For Each pathCsv In Ldirectorio
            filesListBox.Items.Add(vbCrLf)
            filesListBox.Items.Add("** Modificar Csv **")
            directorioCSV = CStr(pathCsv)
            filesListBox.Items.Add("  " & directorioCSV)

            'Modificamos el CSV
            ModificarCsv(directorioCSV)
            filesListBox.Items.Add("** Modificacion Csv Realizada **")
            'Modificamos los Trafos
            Dim ultimaBarra As Integer = directorioCSV.LastIndexOf("\")
            directorioTrafo = directorioCSV.Substring(0, ultimaBarra)
            ModificarTrafo(directorioTrafo)
        Next

    End Sub
    'ModificarTrafo
    Private Sub ModificarTrafo(ByVal directorioTrafo As String)
        'variables Globales
        Dim trafo As String
        Dim lineas() As String
        Dim lineasInicio() As String
        Dim buscar As String
        Dim primerCaracter As Integer

        'Modificar trafo01
        trafo = directorioTrafo + "\$Transformationen\TraFo1.txt"
        'Lectura de fichero
        lineas = System.IO.File.ReadAllLines(trafo, Encoding.Default)
        lineasInicio = System.IO.File.ReadAllLines(trafo, Encoding.Default)
        'Cambio de variables
        buscar = " -245.000;"
        primerCaracter = lineas(9).IndexOf(buscar)
        If primerCaracter <> -1 Then
            filesListBox.Items.Add(vbCrLf)
            filesListBox.Items.Add("** Modificar TraFo1 **")
            filesListBox.Items.Add("  " & trafo)
            filesListBox.Items.Add("    " & lineas(9))
            lineas(9) = "Translation = 0.000;0.000;0.000"
            filesListBox.Items.Add("    " & lineas(9))
            filesListBox.Items.Add("** Modificacion realizada TraFo1 **")
        End If
        System.IO.File.WriteAllLines(trafo, lineas)

        'Modificar trafo02
        trafo = directorioTrafo + "\$Transformationen\TraFo2.txt"
        'Lectura de fichero
        lineas = System.IO.File.ReadAllLines(trafo, Encoding.Default)
        lineasInicio = System.IO.File.ReadAllLines(trafo, Encoding.Default)
        'Cambio de variables
        buscar = " 0.000;"
        primerCaracter = lineas(9).IndexOf(buscar)
        If primerCaracter <> -1 Then
            filesListBox.Items.Add(vbCrLf)
            filesListBox.Items.Add("** Modificar TraFo2 **")
            filesListBox.Items.Add("  " & trafo)
            filesListBox.Items.Add("    " & lineas(9))
            lineas(9) = "Translation = +245.000;0.000;0.000"
            filesListBox.Items.Add("    " & lineas(9))
            filesListBox.Items.Add("** Modificacion realizada TraFo2 **")
        End If
        System.IO.File.WriteAllLines(trafo, lineas)

        'Modificar trafo03
        trafo = directorioTrafo + "\$Transformationen\TraFo3.txt"
        'Lectura de fichero
        lineas = System.IO.File.ReadAllLines(trafo, Encoding.Default)
        lineasInicio = System.IO.File.ReadAllLines(trafo, Encoding.Default)
        'Cambio de variables
        buscar = " 230.000;"
        primerCaracter = lineas(9).IndexOf(buscar)
        If primerCaracter <> -1 Then
            filesListBox.Items.Add(vbCrLf)
            filesListBox.Items.Add("** Modificar TraFo3 **")
            filesListBox.Items.Add("  " & trafo)
            filesListBox.Items.Add("    " & lineas(9))
            lineas(9) = "Translation = 475.000;0.000;0.000"
            filesListBox.Items.Add("    " & lineas(9))
            filesListBox.Items.Add("** Modificacion realizada TraFo3 **")
        End If
        System.IO.File.WriteAllLines(trafo, lineas)

    End Sub
    'Modificar Csv
    Private Sub ModificarCsv(ByVal directorioCSV As String)
        ' Leer Todos los registros del fichero
        Dim registros As String() = File.ReadAllLines(directorioCSV, Encoding.Default)
        Dim listaFilas As List(Of String) = New List(Of String)()
        For Each fila As String In registros
            ' Por cada fila separa los atributos por comas
            Dim valorAtributo As String() = fila.Split(","c)
            'Solo se tratan las filas cuyo primer atributo corresponda con alguno de estos
            If valorAtributo(0) = "PT" Or
               valorAtributo(0) = "SLT" Or
               valorAtributo(0) = "CIR" Or
               valorAtributo(0) = "BPT" Or
               valorAtributo(0) = "PTC" Or
               valorAtributo(0) = "OPR" Then

                If valorAtributo.Length > 2 Then
                    Dim coorX As Double
                    Try
                        coorX = Convert.ToDouble(valorAtributo(2))
                    Catch e As FormatException
                    Catch e As OverflowException
                    End Try
                    filesListBox.Items.Add("    ")
                    filesListBox.Items.Add("    " & fila)
                    coorX = coorX - 245
                    valorAtributo(2) = String.Format("{0:0.00}", coorX)
                    'valorAtributo(3) = String.Format("{0}", coorX)
                    filesListBox.Items.Add("    " & String.Join(",", valorAtributo))

                End If
            End If
            ' Añadimos a la lista la fila 
            listaFilas.Add(String.Join(",", valorAtributo))
        Next
        ' Salvamos la nueva lista en el mismo fichero CSV
        File.WriteAllLines(directorioCSV, listaFilas.ToArray(), Encoding.Default)
    End Sub
    'Modificar el Pie de todos los Trafos
    Private Sub ModificarPieTrafoDirectorioCmompeto()
        'filesListBox.Items.Add(vbCrLf)
        'filesListBox.Items.Add("** Buscando archivos Trafo0*.txt en el todo el directorio **")
        'Buscamos todos los Trafos del directorio
        BuscarArchivosTrafoX(directorioDestino)
        filesListBox.Items.Add(vbCrLf)
        filesListBox.Items.Add("** Modificar pie de archivos TraFo*.txt en todo el directorio **")
        'Una vez Encontrados todos los trafos del directorio
        ModificarPie()

    End Sub

    Private Sub BuscarArchivosTrafoX(ByVal directorioDestino As String)
        Dim directorioAModificarTrafoX As String
        Dim ultimaBarra As Integer
        Try
            Dim dir As New System.IO.DirectoryInfo(directorioDestino)
            'Buscamos Archivos en todo el directorio los archivos que tengan extension trafo0*.txt
            Dim fileList = dir.GetFiles("TraFo*.*", System.IO.SearchOption.AllDirectories)
            Dim fileQuery = From file In fileList _
                             Where file.Extension =
                             ".txt" _
                             Order By file.Name _
                             Select file

            For Each file In fileQuery
                directorioAModificarTrafoX = file.FullName
                'Comprobamos si existen los trafos1,2 y 3
                ultimaBarra = directorioAModificarTrafoX.LastIndexOf("\")
                'En Ldirectoriotrafo guardamos el path completo de los trafos de todo el directorio
                Ldirectoriotrafo.Add(directorioAModificarTrafoX)
            Next
        Catch ex As Exception
            MsgBox("Descripcion del error:" & _
                   ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub ModificarPie()
        Dim archivoTrafo As String
        Dim fichero As String
        Dim lineas() As String
        Dim lineasInicio() As String
        Dim primerCaracter As Integer
        Dim datoscambio As String
        Dim Inicio As String

        'Cambiar Pie
        For Each archivoTrafo In Ldirectoriotrafo
            fichero = archivoTrafo
            'filesListBox.Items.Add(vbCrLf)
            filesListBox.Items.Add("  " & fichero)
            'Lectura de fichero

            lineas = System.IO.File.ReadAllLines(fichero, Encoding.Default)

            lineasInicio = System.IO.File.ReadAllLines(fichero, Encoding.Default)
            If lineasInicio.Length > 17 Then

                'Cambio linea 12 --> 18 
                primerCaracter = lineasInicio(12).IndexOf("=")
                datoscambio = lineasInicio(12).Substring(primerCaracter - 1)

                primerCaracter = lineasInicio(18).IndexOf("=")
                Inicio = lineasInicio(18).Substring(0, primerCaracter - 1)
                lineas(18) = Inicio + datoscambio

                filesListBox.Items.Add("  12) " & lineasInicio(12) & " --> 18) " & lineas(18))


                'Cambio linea 13 --> 12
                primerCaracter = lineasInicio(13).IndexOf("=")
                datoscambio = lineasInicio(13).Substring(primerCaracter - 1)

                primerCaracter = lineasInicio(12).IndexOf("=")
                Inicio = lineasInicio(12).Substring(0, primerCaracter - 1)
                lineas(12) = Inicio + datoscambio

                filesListBox.Items.Add("  13) " & lineasInicio(13) & "    --> 12) " & lineas(12))

                'Cambiar linea 17
                primerCaracter = lineasInicio(17).IndexOf("=")
                datoscambio = lineasInicio(17).Substring(primerCaracter - 1)

                primerCaracter = lineasInicio(17).IndexOf("=")
                Inicio = "Fzg-Konf.Ländervariante"
                lineas(17) = Inicio + datoscambio

                filesListBox.Items.Add("  17) " & lineasInicio(17) & "    --> 17) " & lineas(17))

                'Cambio linea 18 --> 13
                primerCaracter = lineasInicio(18).IndexOf("=")
                datoscambio = lineasInicio(18).Substring(primerCaracter - 1)

                primerCaracter = lineasInicio(13).IndexOf("=")
                Inicio = lineasInicio(13).Substring(0, primerCaracter - 1)
                lineas(13) = Inicio + datoscambio

                filesListBox.Items.Add("  18) " & lineasInicio(18) & "    --> 13) " & lineas(13))
            End If
            System.IO.File.WriteAllLines(fichero, lineas, Encoding.Default)
        Next
        filesListBox.Items.Add("** Modificados los pies de archivos TraFo*.txt en todo el directorio **")

    End Sub

    Private Sub generarLog()

        FileOpen(FileNumber, directorio & "\log.txt", OpenMode.Output)
        For Each Item As Object In filesListBox.Items
            PrintLine(FileNumber, Item.ToString)
        Next
        FileClose(FileNumber)
    End Sub

    Private Sub borrarLog()

        Using file As New IO.StreamWriter(directorio & "\log.txt")
            file.Flush()
        End Using
    End Sub

    Private Sub PictureBox1_Click(sender As System.Object, e As System.EventArgs) Handles PictureBox1.Click

    End Sub
End Class