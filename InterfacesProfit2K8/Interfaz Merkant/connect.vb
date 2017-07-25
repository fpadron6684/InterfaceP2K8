Imports System.Data.SqlClient

Public Class connect
    Private bdName As String
    Private inst As String
    Public CONN1 As New SqlConnection
    Public Property BD() As String
        Get
            ' Gets the property value.
            Return bdName
        End Get

        Set(ByVal Value As String)
            ' Sets the property value.
            bdName = Value
        End Set
    End Property

    Public Property sql() As String
        Get
            ' Gets the property value.
            Return inst
        End Get

        Set(ByVal Value As String)
            ' Sets the property value.
            inst = Value
        End Set
    End Property

    Public Property connName() As SqlConnection
        Get
            ' Gets the property value.
            Return CONN1
        End Get

        Set(ByVal Value As SqlConnection)
            ' Sets the property value.
            CONN1 = Value
        End Set
    End Property


    Public Sub New(ByVal db As String, ByVal insSql As String)
        ' Set the property value.
        Me.bdName = db
        Me.inst = insSql
    End Sub


    Public Sub conectar()
        Try
            CONN1.ConnectionString = "data source =" & inst & "; INitial catalog = " & bdName & "; user id = profit; password = profit" ' Base de Datos de Prueba
            'CONN1.ConnectionString = "data source = 192.168.1.225; INitial catalog = " & bdName & "; user id = profit; password = profit" ' Base de Datos de Prueba
            CONN1.Open()
        Catch ex As Exception
            EscribirLog("Error al Conectar a la Base de Datos <<" & bdName & ">> " & ex.Message, EventLogEntryType.INformation)
        End Try

    End Sub
    Public Sub conectar2()
        Try
            'CONN1.Connectionstring = "data source = 192.168.1.151; INitial catalog = auxiliares; user id = profit; password = profit" ' Base de Datos de Prueba
            CONN1.ConnectionString = "data source = " & inst & "; INitial catalog = auxiliares; user id = profit; password = profit" ' Base de Datos de Prueba
            CONN1.Open()
        Catch ex As Exception
            EscribirLog("Error al Conectar a la Base de Datos <<AUXILIARES>> " & ex.Message, EventLogEntryType.INformation)
        End Try

    End Sub

    Private Sub EscribirLog(ByVal Texto_Evento As String, ByVal tipo_entrada As EventLogEntryType)
        Dim MaquINa As String = "."
        Dim Origen As String = "Interface Merkant"
        'Escribimos en los Registros de Aplicación
        Dim Elog As EventLog
        Elog = New EventLog("Application", MaquINa, Origen)
        Elog.WriteEntry(Texto_Evento, tipo_entrada, 100, CType(50, Short))
        Elog.Close()
        Elog.Dispose()
    End Sub

    Public Sub cerrar()
        CONN1.Close()
        CONN1.Dispose()
    End Sub
End Class
