Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Data.SqlClient

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://localhost/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class WSConsultaExcel
    Inherits System.Web.Services.WebService

    'Conexion al archivo Excel
    'Dim path_archivo_xls As String = "C:\Users\CAMH\Desktop\Practica_Excel.xlsx"
    Dim path_archivo_xls As String = "C:\inetpub\wwwroot\WSCExcel\Practica_Excel.xlsx"
    Dim cadena_xls As String = "Provider=Microsoft.ACE.OLEDB.12.0;" &
                                    "Extended Properties='Excel 12.0 Xml;HDR=Yes;';" &
                                    "Data Source=" & path_archivo_xls
    'Conexion a la base de datos
    Dim conexion_db As SqlConnection = New SqlConnection("Data Source=DESKTOP-NLCUO3R;Initial Catalog=PRUEBA_PROPIETARIO;User ID=sa;Password=123456")

    Dim conexion_xls As OleDbConnection = New OleDbConnection(cadena_xls)
    Dim comando As SqlCommand
    Dim comando_xls As OleDbCommand
    Dim conjunto_datos As New DataSet
    Dim filas_consultadas As Integer

    <WebMethod(Description:="Consulta los datos en un archivo excel")>
    Public Function ConsultaDatosExcel() As String
        conexion_xls.Open()
        Dim adaptador As New OleDbDataAdapter
        Dim conjunto_datos As OleDbDataReader

        Dim str As String
        comando_xls = New OleDbCommand("Select * from [datos$]", conexion_xls)
        conjunto_datos = comando_xls.ExecuteReader
        str = "Resumen de Datos" & StrDup(2, vbCr)
        Do While conjunto_datos.Read
            str = str & "La persona con ID = " & conjunto_datos.GetString(0) & " es " & conjunto_datos.GetString(1) & " " & conjunto_datos.GetString(2) & " tiene " & conjunto_datos.GetString(3) & " años y vive en " & conjunto_datos.GetString(4) & StrDup(1, vbCr)
        Loop
        conexion_xls.Close()
        Return str
    End Function

    <WebMethod(Description:="Insersion de registro en la tabla datos del archivo excel")>
    Public Function InsertarDatosExcel(id As String, nombre As String, apellido As String, edad As String, direccion As String) As String

        conexion_xls.Open()

        Try
            comando_xls = New OleDbCommand("Insert into [datos$] values (@id,@n,@a,@e,@d)", conexion_xls)
            comando_xls.Parameters.AddWithValue("@id", id)
            comando_xls.Parameters.AddWithValue("@n", nombre)
            comando_xls.Parameters.AddWithValue("@a", apellido)
            comando_xls.Parameters.AddWithValue("@e", edad)
            comando_xls.Parameters.AddWithValue("@d", direccion)
            comando_xls.ExecuteNonQuery()
            conexion_xls.Close()
        Catch ex As Exception
            Return "Error"

        End Try

        Return "OK...."
    End Function

    <WebMethod(Description:="Insersion de registros en la tabla marcas del archivo excel")>
    Public Function InsertarMarcasExcel(id As String, nombre As String) As String
        Dim comando As OleDbCommand
        conexion_xls.Open()

        Try
            comando = New OleDbCommand("Insert into [marcas$] values (@id,@n)", conexion_xls)
            comando.Parameters.AddWithValue("@id", id)
            comando.Parameters.AddWithValue("@n", nombre)
            comando.ExecuteNonQuery()
            conexion_xls.Close()

        Catch ex As Exception
            Return "Error"
        End Try

        Return "OK...."
    End Function

    <WebMethod(Description:="Carga la tabla Producto de la base de datos en un archivo excel")>
    Public Function CargarProductos() As String
        '-----
        Dim mensaje = "OK...."
        Call crear_hoja_productos(mensaje)
        Call consultar_productos()

        Dim fila As DataRow
        Dim indice As Integer

        conexion_xls.Open()
        For indice = 0 To filas_consultadas - 1
            fila = conjunto_datos.Tables("productos").Rows(indice)
            comando_xls = New OleDbCommand("Insert into [productos$] values (@id,@n,@c,@s,@p)", conexion_xls)
            comando_xls.Parameters.AddWithValue("@id", fila(0))
            comando_xls.Parameters.AddWithValue("@n", fila(1))
            comando_xls.Parameters.AddWithValue("@c", fila(2))
            comando_xls.Parameters.AddWithValue("@s", fila(3))
            comando_xls.Parameters.AddWithValue("@p", fila(4))
            comando_xls.ExecuteNonQuery()
        Next indice
        conexion_xls.Close()
        Return mensaje
    End Function

    'Funcion para crear la hoja productos en el archivo excel
    Private Sub crear_hoja_productos(mensaje As String)
        conexion_xls.Open()
        Try
            'comando.Connection = conexion_xls
            'comando_xls = conexion_xls.CreateCommand()
            comando_xls = New OleDbCommand("Create table [productos]
                                           (Producto_ID int,
                                            Producto_Nombre char(30),
                                            Categoria char(30),
                                            Producto_Stock int,
                                            Producto_Precio float)", conexion_xls)
            comando_xls.ExecuteNonQuery()
        Catch ex As Exception
            mensaje = "Error"
        End Try
        mensaje = "OK...."
        conexion_xls.Close()
    End Sub

    'Funcion que sirve para consultar la tabla Producto de la base de datos
    Private Sub consultar_productos()
        conexion_db.Open()
        comando = New SqlCommand("Select p.Producto_ID, p.Producto_Nombre, c.Categoria_Nombre, p.Producto_Stock, p.Producto_Precio 
                                  From Categoria c,Producto p 
                                  Where c.Categoria_ID = p.Categoria_ID", conexion_db)
        Dim adaptador As New SqlDataAdapter
        adaptador.SelectCommand = comando
        adaptador.Fill(conjunto_datos, "productos")
        filas_consultadas = conjunto_datos.Tables("productos").Rows.Count
        conexion_db.Close()
    End Sub

    <WebMethod(Description:="Carga la tabla Categoria de la base datos a un archi excel")>
    Public Function CargarCategorias() As String
        Dim mensaje As String = "OK..."
        Call crear_hoja_categorias(mensaje)
        Call consultar_categorias()

        conexion_xls.Open()

        Dim fila As DataRow
        Dim indice As Integer

        For indice = 0 To filas_consultadas - 1
            fila = conjunto_datos.Tables("categorias").Rows(indice)
            comando_xls = New OleDbCommand("Insert into [categorias$] values(@id,@n,@d)", conexion_xls)
            comando_xls.Parameters.AddWithValue("@id", fila(0))
            comando_xls.Parameters.AddWithValue("@n", fila(1))
            comando_xls.Parameters.AddWithValue("@d", fila(2))
            comando_xls.ExecuteNonQuery()
        Next indice

        conexion_xls.Close()
        Return mensaje
    End Function

    'Funcion para crear la hoja categorias en el archivo excel
    Private Sub crear_hoja_categorias(mensaje As String)
        conexion_xls.Open()
        Try
            'comando.Connection = conexion_xls
            'comando_xls = conexion_xls.CreateCommand()
            comando_xls = New OleDbCommand("Create table [categorias]
                                           (Categoria_ID int,
                                            Categoria_Nombre char(30),
                                            Categoria_Descripcion char(30))", conexion_xls)
            comando_xls.ExecuteNonQuery()
        Catch ex As Exception
            mensaje = "Error"
        End Try
        mensaje = "OK...."
        conexion_xls.Close()
    End Sub

    'Funcion para consultar los tabla Categorias de la base de datos
    Private Sub consultar_categorias()
        conexion_db.Open()
        comando = New SqlCommand("Select * From Categoria", conexion_db)
        Dim adaptador As New SqlDataAdapter
        adaptador.SelectCommand = comando
        adaptador.Fill(conjunto_datos, "categorias")
        filas_consultadas = conjunto_datos.Tables("categorias").Rows.Count
        conexion_db.Close()
    End Sub

    <WebMethod(Description:="Carga la tabla Marca de la base de datos en un archivo excel")>
    Public Function CargarMarcas() As String
        Call Crear_hoja_marcas()
        Call Consultar_marcas()
        Dim fila As DataRow
        Dim indice As Integer

        conexion_xls.Open()

        For indice = 0 To filas_consultadas - 1
            fila = conjunto_datos.Tables("marcas").Rows(indice)
            comando_xls = New OleDbCommand("Insert into [marcas$] values (@id,@n)", conexion_xls)
            comando_xls.Parameters.AddWithValue("@id", fila(0))
            comando_xls.Parameters.AddWithValue("@n", fila(1))
            comando_xls.ExecuteNonQuery()
        Next indice

        conexion_xls.Close()
        Return "OK..."
    End Function

    'Funcion que sirve para crear la hoja marcas en el archivo excel
    Public Sub Crear_hoja_marcas()
        Dim s As String
        conexion_xls.Open()
        Try
            'comando.Connection = conexion_xls
            comando_xls = conexion_xls.CreateCommand()
            comando_xls.CommandText = "Create table [marcas] " &
                                       "(Marca_ID int, " &
                                       "Nombre char(30)) "
            comando_xls.ExecuteNonQuery()
        Catch ex As Exception
            s = ex.Message
        End Try
        conexion_xls.Close()
    End Sub

    'Funcion que sirve para consultar la tabla marcas de la base de datos
    Public Sub Consultar_marcas()
        conexion_db.Open()
        comando = New SqlCommand("Select * From Marca", conexion_db)
        Dim adaptador As SqlDataAdapter = New SqlDataAdapter
        adaptador.SelectCommand = comando
        adaptador.Fill(conjunto_datos, "marcas")
        filas_consultadas = conjunto_datos.Tables("marcas").Rows.Count
        conexion_db.Close()
    End Sub

    <WebMethod(Description:="Carga la tabla datos de excel a la base de datos")>
    Public Function CargarTablaDatos() As String
        Dim mensaje As String = "OK..."
        Call Crear_tabla_datos(mensaje)
        Call Consultar_hoja_datos()
        Dim fila As DataRow
        Dim indice As Integer

        conexion_db.Open()

        For indice = 0 To filas_consultadas - 1
            fila = conjunto_datos.Tables("datos").Rows(indice)
            comando = New SqlCommand("Insert into Datos values (@id,@n,@a,@e,@d)", conexion_db)
            comando.Parameters.AddWithValue("@id", fila(0))
            comando.Parameters.AddWithValue("@n", fila(1))
            comando.Parameters.AddWithValue("@a", fila(2))
            comando.Parameters.AddWithValue("@e", fila(3))
            comando.Parameters.AddWithValue("@d", fila(4))
            comando.ExecuteNonQuery()
        Next indice
        conexion_db.Close()
        Return mensaje
    End Function

    'Funcion que sirve para crear la tabla Datos en la Base de datos
    Private Sub Crear_tabla_datos(mensaje As String)
        conexion_db.Open()
        Try
            'comando.Connection = conexion_xls
            comando = conexion_db.CreateCommand()
            comando.CommandText = "CREATE TABLE Datos
                                       (Codigo char(50),
                                        Nombre char(50),
                                        Apellido char(50),
                                        Edad char(10),
                                        Direccion char(40))"
            comando.ExecuteNonQuery()
            mensaje = "Tabla Creada Sastisfatioriamente"
        Catch ex As Exception
            mensaje = ex.Message
        End Try
        conexion_db.Close()
    End Sub

    'Funcion para consultar los datos de la hoja de excel
    Private Sub Consultar_hoja_datos()
        conexion_xls.Open()
        comando_xls = New OleDbCommand("Select * From [datos$]", conexion_xls)
        Dim adaptador As OleDbDataAdapter = New OleDbDataAdapter
        adaptador.SelectCommand = comando_xls
        adaptador.Fill(conjunto_datos, "datos")
        filas_consultadas = conjunto_datos.Tables("datos").Rows.Count
        conexion_xls.Close()
    End Sub

End Class