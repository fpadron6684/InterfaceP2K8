Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports System.Timers.Timer
Imports System.Timers

Public Class InterfazMerkant

    Private Declare Auto Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal procHandle As IntPtr, ByVal min As Int32, ByVal max As Int32) As Boolean

    'Funcion de liberacion de memoria
    Public Sub ClearMemory()

        Try
            Dim Mem As Process
            Mem = Process.GetCurrentProcess()
            SetProcessWorkingSetSize(Mem.Handle, -1, -1)
        Catch ex As Exception
            'Control de errores
        End Try

    End Sub


    Private tiempo As New System.Timers.Timer
    Private tiempoII As New System.Timers.Timer


    Private varConfGral(16) As String
    Private suc(4) As String
    Private ruta(4) As String
    Private bd(4) As String
    Private sSets(34) As Boolean
    Private Fecha As String
    Private Unique As Guid
    Private artGral As String
    Private cantGral As Double
    Private FechaCob As String
    Private flag1 As Boolean = False
    Private flag2 As Boolean = False
    Private flag3 As Boolean = False

    Private cmdSQL As New SqlCommand
    Private tabla As New DataTable
    Private tablaTrans As New DataTable
    Private tablaTrans2 As New DataTable
    Private tablaAct As New DataTable
    Private tablaComp As New DataTable
    Private adp As New SqlDataAdapter
    Private cmdBld As New SqlCommandBuilder

    Private tbl_gral As New DataTable

    Dim SaldoD As Double
    Dim SaldoC As Double

    Private cont As Integer = 0
    Private NC As Boolean = False
    Private fileDir(14) As String






    Protected Overrides Sub OnStart(ByVal args() As String)
        '/////////////////////CARGA DE ARCHIVOS////////////////////////

        'Lee el archivo de configuracion del servicio
        Dim confGral As String = "C:\ConfiguracionGeneral.txt"
        Dim confSuc As String = "C:\ConfiguracionSucursal.txt"

        If Dir$(confGral) <> "" Then
            Dim iniFileConf As New StreamReader(confGral)

            For i As Integer = 0 To varConfGral.GetUpperBound(0)
                varConfGral(i) = iniFileConf.ReadLine()
            Next

            iniFileConf.Close()
            'tiempo de iteracion del temporizador
            Dim time As Integer = CInt(varConfGral(0))
            Dim timeII As Integer = CInt(varConfGral(16))

            AddHandler tiempo.Elapsed, AddressOf tiempo_Tick
            AddHandler tiempoII.Elapsed, AddressOf tiempoII_Tick

            tiempo.Enabled = True
            tiempo.Interval = time

            tiempoII.Enabled = True
            tiempoII.Interval = timeII


            'determina cuantas sucursales habran activas (5 maximo) posiciones impares
            'determina rutas de las carpetas de las sucursales (5maximo) posiciones pares
            suc(0) = varConfGral(1)
            ruta(0) = varConfGral(2)
            bd(0) = varConfGral(3)

            suc(1) = varConfGral(4)
            ruta(1) = varConfGral(5)
            bd(1) = varConfGral(6)

            suc(2) = varConfGral(7)
            ruta(2) = varConfGral(8)
            bd(2) = varConfGral(9)

            suc(3) = varConfGral(10)
            ruta(3) = varConfGral(11)
            bd(3) = varConfGral(12)

            suc(4) = varConfGral(13)
            ruta(4) = varConfGral(14)
            bd(4) = varConfGral(15)


            If Dir$(confSuc) <> "" Then
                Dim iniFileSuc As New StreamReader(confSuc)
                For i As Integer = 0 To 34
                    sSets(i) = CBool(iniFileSuc.ReadLine())
                Next
                iniFileSuc.Close()
            Else
                MsgBox("No Hay Archivo de Configuración SubSets Disponible, debe generarlo en la aplicacion de Configuración", MsgBoxStyle.Critical)
            End If

            If suc(0) = 1 Then
                vigilante1.EnableRaisingEvents = True
                vigilante1.Path = ruta(0) + "\Importacion\"
            End If
            If suc(1) = 1 Then
                vigilante2.EnableRaisingEvents = True
                vigilante2.Path = ruta(1) + "\Importacion\"
            End If
            If suc(2) = 1 Then
                vigilante3.EnableRaisingEvents = True
                vigilante3.Path = ruta(2) + "\Importacion\"
            End If
            If suc(3) = 1 Then
                vigilante4.EnableRaisingEvents = True
                vigilante4.Path = ruta(3) + "\Importacion\"
            End If
            If suc(4) = 1 Then
                vigilante5.EnableRaisingEvents = True
                vigilante5.Path = ruta(4) + "\Importacion\"
            End If




        Else
            MsgBox("No Hay Archivo de Configuración General Disponible, debe generarlo en la aplicacion de Configuración", MsgBoxStyle.Critical)
            'End
        End If

        '/////////////////////EXTRACCION DE ARCHIVOS////////////////////////


    End Sub

#Region "FileSystemWatcher"
    'RUTA(0)
    'BD(0)
    'SUC(0)
    'SSETS (0 A 6)
    Private Sub sspedidos(fileDir As String, fileDir1 As String, BD As String)
        Dim conn As New connect(BD) 'aclarar instancia SQL
        conn.conectar()

        Dim Trans As SqlTransaction
        Trans = conn.connName.BeginTransaction

        Try
            cmdSQL.Connection = conn.connName
            cmdSQL.Transaction = Trans
            cmdSQL.CommandType = CommandType.StoredProcedure

            pedidos(fileDir, cmdSQL)
            pedidoDetalle(fileDir1, cmdSQL, conn, Trans)

            Trans.Commit()
            EscribirLog("Sub-Set de archivos <<PEDIDOS>> fue cargado con Exito", EventLogEntryType.Information)

        Catch ex As Exception
            Trans.Rollback()
            EscribirLog("La carga del Sub-Set de archivos <<PEDIDOS>> ha presentado el siguiente error " & ex.Message, EventLogEntryType.Error)

        Finally
            conn.cerrar()
        End Try
    End Sub
    Private Sub ssdevolucion(fileDir As String, fileDir1 As String, BD As String)
        Dim conn As New connect(BD) 'aclarar instancia SQL
        conn.conectar()

        Dim Trans As SqlTransaction
        Trans = conn.connName.BeginTransaction

        Try
            cmdSQL.Connection = conn.connName
            cmdSQL.Transaction = Trans
            cmdSQL.CommandType = CommandType.StoredProcedure

            devolucion(fileDir, cmdSQL)
            devoluciondetalle(fileDir1, cmdSQL, conn)

            Trans.Commit()
            EscribirLog("Sub-Set de archivos <<DEVOLUCIONES>> fue cargado con Exito", EventLogEntryType.Information)

        Catch ex As Exception
            Trans.Rollback()
            EscribirLog("La carga del Sub-Set de archivos <<DEVOLUCIONES>> ha presentado el siguiente error " & ex.Message, EventLogEntryType.Error)

        Finally
            conn.cerrar()
        End Try
    End Sub
    Private Sub sscobranza(fileDir As String, fileDir1 As String, fileDir2 As String, bd As String)
        Dim conn As New connect(bd) 'aclarar instancia SQL
        conn.conectar()

        Dim Trans As SqlTransaction
        Trans = conn.connName.BeginTransaction

        Try
            cmdSQL.Connection = conn.connName
            cmdSQL.Transaction = Trans
            cmdSQL.CommandType = CommandType.StoredProcedure
       
            cobranza(fileDir, cmdSQL)
            cobranzadetalle(fileDir1, cmdSQL, conn)
            cobranzapago(fileDir2, cmdSQL, conn)


            Trans.Commit()
            EscribirLog("Sub-Set de archivos <<COBRANZAS>> fue cargado con Exito", EventLogEntryType.Information)

        Catch ex As Exception
            Trans.Rollback()
            EscribirLog("La carga del Sub-Set de archivos <<COBRANZAS>> ha presentado el siguiente error " & ex.Message, EventLogEntryType.Error)

        Finally
            conn.cerrar()

        End Try
    End Sub
    Private Sub ssautovta(fileDir As String, fileDir1 As String, bd As String)
        Dim conn As New connect(bd) 'aclarar instancia SQL
        conn.conectar()

        Dim Trans As SqlTransaction
        Trans = conn.connName.BeginTransaction

        Try
            cmdSQL.Connection = conn.connName
            cmdSQL.Transaction = Trans
            cmdSQL.CommandType = CommandType.StoredProcedure
 
            factura(fileDir, cmdSQL)
            facturaDetalle(fileDir1, cmdSQL, conn)

            Trans.Commit()
            EscribirLog("Sub-Set de archivos <<FACTURAS AUTOVENTA>> fue cargado con Exito", EventLogEntryType.Information)

        Catch ex As Exception
            Trans.Rollback()
            EscribirLog("La carga del Sub-Set de archivos <<FACTURAS AUTOVENTA>> ha presentado el siguiente error " & ex.Message, EventLogEntryType.Error)

        Finally
            conn.cerrar()

        End Try

    End Sub
    Private Sub ssNCautovta(fileDir As String, bd As String)
        Dim conn As New connect(bd) 'aclarar instancia SQL
        conn.conectar()

        Dim Trans As SqlTransaction
        Trans = conn.connName.BeginTransaction

        Try
            cmdSQL.Connection = conn.connName
            cmdSQL.Transaction = Trans
            cmdSQL.CommandType = CommandType.StoredProcedure

            notaCredito(fileDir, cmdSQL, conn)

            Trans.Commit()
            EscribirLog("Sub-Set de archivos <<N/CR AUTOVENTA>> fue cargado con Exito", EventLogEntryType.Information)

        Catch ex As Exception
            Trans.Rollback()
            EscribirLog("La carga del Sub-Set de archivos <<N/CR AUTOVENTA>> ha presentado el siguiente error " & ex.Message, EventLogEntryType.Error)

        Finally
            conn.cerrar()
        End Try
    End Sub
    Private Sub ssdepositos(fileDir As String, fileDir1 As String, bd As String)
        Dim conn As New connect(bd) 'aclarar instancia SQL
        conn.conectar()

        Dim Trans As SqlTransaction
        Trans = conn.connName.BeginTransaction

        Try
            cmdSQL.Connection = conn.connName
            cmdSQL.Transaction = Trans
            cmdSQL.CommandType = CommandType.StoredProcedure

            deposito(fileDir, cmdSQL)
            depositoDetalle(fileDir1, cmdSQL, conn)
                
            Trans.Commit()
            EscribirLog("Sub-Set de archivos <<DEPOSITOS>> fue cargado con Exito", EventLogEntryType.Information)

        Catch ex As Exception
            Trans.Rollback()
            EscribirLog("La carga del Sub-Set de archivos <<DEPOSITOS>> ha presentado el siguiente error " & ex.Message, EventLogEntryType.Error)


        Finally
            conn.cerrar()

        End Try
    End Sub



    Private Sub vigilante1_Created(sender As Object, e As IO.FileSystemEventArgs) Handles vigilante1.Created
        Dim ename As String


        ename = Reemplaza2(e.Name)

        If suc(0) = 1 Then

            'SubSet I Pedidos
            If sSets(0) = True Then

                If ename Like "Pedido2*" Or ename Like "PedidoDetalle*" Then
                    If ename Like "Pedido2*" Then
                        fileDir(0) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(1) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        sspedidos(fileDir(0), fileDir(1), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(1))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                End If
            End If


            'SubSet II Devoluciones
            If sSets(1) = True Then

                If ename Like "Devolucion2*" Or ename Like "DevolucionDetalle*" Then
                    If ename Like "Devolucion2*" Then
                        fileDir(2) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(3) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdevolucion(fileDir(2), fileDir(3), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(2))
                        My.Computer.FileSystem.DeleteFile(fileDir(3))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If


                End If
            End If

            'SubSet III Cobranzas
            If sSets(2) = True Then
                If ename Like "Cobranza2*" Or ename Like "CobranzaDetalle*" Or ename Like "CobranzaPago*" Then
                    If ename Like "Cobranza2*" Then
                        fileDir(4) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaDetalle*" Then
                        fileDir(5) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaPago*" Then
                        fileDir(6) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 3 Then
                        sscobranza(fileDir(4), fileDir(5), fileDir(6), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(4))
                        My.Computer.FileSystem.DeleteFile(fileDir(5))
                        My.Computer.FileSystem.DeleteFile(fileDir(6))
                        cont = 0
                    End If
                End If
            End If

            'SubSet IV Facturas AutoVenta
            If sSets(2) = True Then
                If ename Like "Factura2*" Or ename Like "FacturaDetalle*" Or ename Like "NotaCredito*" Then
                    If ename Like "Factura2*" Then
                        fileDir(7) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "FacturaDetalle*" Then
                        fileDir(8) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "NotaCredito*" Then
                        fileDir(9) = ruta(0) + "\Importacion\" + ename
                        NC = True
                    End If

                    If cont = 2 Then
                        ssautovta(fileDir(7), fileDir(8), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        NC = False
                    End If
                End If
            End If

            'SubSet V Depositos
            If sSets(2) = True Then
                If ename Like "Deposito2*" Or ename Like "DepositoDetalle*" Then
                    If ename Like "Deposito2*" Then
                        fileDir(10) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "DepositoDetalle*" Then
                        fileDir(11) = ruta(0) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdepositos(fileDir(9), fileDir(10), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        My.Computer.FileSystem.DeleteFile(fileDir(10))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If

                End If
            End If
        End If
    End Sub

    'RUTA(1)
    'BD(1)
    'SUC(1)
    'SSETS (7 A 13)
    Private Sub vigilante2_Created(sender As Object, e As IO.FileSystemEventArgs) Handles vigilante2.Created
            Dim ename As String


        ename = Reemplaza2(e.Name)

        If suc(1) = 1 Then

            'SubSet I Pedidos
            If sSets(0) = True Then

                If ename Like "Pedido2*" Or ename Like "PedidoDetalle*" Then

                    If ename Like "Pedido2*" Then
                        fileDir(0) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(1) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        sspedidos(fileDir(0), fileDir(1), bd(1))
                        My.Computer.FileSystem.DeleteFile(fileDir(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(1))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                End If
            End If


            'SubSet II Devoluciones
            If sSets(1) = True Then

                If ename Like "Devolucion2*" Or ename Like "DevolucionDetalle*" Then
                    If ename Like "Devolucion2*" Then
                        fileDir(2) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(3) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdevolucion(fileDir(2), fileDir(3), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(2))
                        My.Computer.FileSystem.DeleteFile(fileDir(3))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If


                End If
            End If

            'SubSet III Cobranzas
            If sSets(2) = True Then
                If ename Like "Cobranza2*" Or ename Like "CobranzaDetalle*" Or ename Like "CobranzaPago*" Then
                    If ename Like "Cobranza2*" Then
                        fileDir(4) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaDetalle*" Then
                        fileDir(5) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaPago*" Then
                        fileDir(6) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 3 Then
                        sscobranza(fileDir(4), fileDir(5), fileDir(6), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(4))
                        My.Computer.FileSystem.DeleteFile(fileDir(5))
                        My.Computer.FileSystem.DeleteFile(fileDir(6))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                End If
            End If

            'SubSet IV Facturas AutoVenta
            If sSets(2) = True Then
                If ename Like "Factura2*" Or ename Like "FacturaDetalle*" Or ename Like "NotaCredito*" Then
                    If ename Like "Factura2*" Then
                        fileDir(7) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "FacturaDetalle*" Then
                        fileDir(8) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "NotaCredito*" Then
                        fileDir(9) = ruta(1) + "\Importacion\" + ename
                        NC = True
                    End If

                    If cont = 2 Then
                        ssautovta(fileDir(7), fileDir(8), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        NC = False
                    End If
                End If
            End If

            'SubSet V Depositos
            If sSets(2) = True Then
                If ename Like "Deposito2*" Or ename Like "DepositoDetalle*" Then
                    If ename Like "Deposito2*" Then
                        fileDir(10) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "DepositoDetalle*" Then
                        fileDir(11) = ruta(1) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdepositos(fileDir(9), fileDir(10), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        My.Computer.FileSystem.DeleteFile(fileDir(10))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If

                End If
            End If
        End If
    End Sub

    'RUTA(2)
    'BD(2)
    'SUC(2)
    'SSETS (14 A 20)
    Private Sub vigilante3_Created(sender As Object, e As IO.FileSystemEventArgs) Handles vigilante3.Created
                   Dim ename As String


        ename = Reemplaza2(e.Name)

        If suc(2) = 1 Then

            'SubSet I Pedidos
            If sSets(0) = True Then

                If ename Like "Pedido2*" Or ename Like "PedidoDetalle*" Then

                    If ename Like "Pedido2*" Then
                        fileDir(0) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(1) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        sspedidos(fileDir(0), fileDir(1), bd(2))
                        My.Computer.FileSystem.DeleteFile(fileDir(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(1))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                End If
            End If


            'SubSet II Devoluciones
            If sSets(1) = True Then

                If ename Like "Devolucion2*" Or ename Like "DevolucionDetalle*" Then
                    If ename Like "Devolucion2*" Then
                        fileDir(2) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(3) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdevolucion(fileDir(2), fileDir(3), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(2))
                        My.Computer.FileSystem.DeleteFile(fileDir(3))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If


                End If
            End If

            'SubSet III Cobranzas
            If sSets(2) = True Then
                If ename Like "Cobranza2*" Or ename Like "CobranzaDetalle*" Or ename Like "CobranzaPago*" Then
                    If ename Like "Cobranza2*" Then
                        fileDir(4) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaDetalle*" Then
                        fileDir(5) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaPago*" Then
                        fileDir(6) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 3 Then
                        sscobranza(fileDir(4), fileDir(5), fileDir(6), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(4))
                        My.Computer.FileSystem.DeleteFile(fileDir(5))
                        My.Computer.FileSystem.DeleteFile(fileDir(6))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                End If
            End If

            'SubSet IV Facturas AutoVenta
            If sSets(2) = True Then
                If ename Like "Factura2*" Or ename Like "FacturaDetalle*" Or ename Like "NotaCredito*" Then
                    If ename Like "Factura2*" Then
                        fileDir(7) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "FacturaDetalle*" Then
                        fileDir(8) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "NotaCredito*" Then
                        fileDir(9) = ruta(2) + "\Importacion\" + ename
                        NC = True
                    End If

                    If cont = 2 Then
                        ssautovta(fileDir(7), fileDir(8), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        NC = False
                    End If
                End If
            End If

            'SubSet V Depositos
            If sSets(2) = True Then
                If ename Like "Deposito2*" Or ename Like "DepositoDetalle*" Then
                    If ename Like "Deposito2*" Then
                        fileDir(10) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "DepositoDetalle*" Then
                        fileDir(11) = ruta(2) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdepositos(fileDir(9), fileDir(10), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        My.Computer.FileSystem.DeleteFile(fileDir(10))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If

                End If
            End If
        End If
    End Sub

    'RUTA(3)
    'BD(3)
    'SUC(3)
    'SSETS (21 A 27)
    Private Sub vigilante4_Created(sender As Object, e As IO.FileSystemEventArgs) Handles vigilante4.Created
             Dim ename As String


        ename = Reemplaza2(e.Name)

        If suc(3) = 1 Then

            'SubSet I Pedidos
            If sSets(0) = True Then

                If ename Like "Pedido2*" Or ename Like "PedidoDetalle*" Then

                    If ename Like "Pedido2*" Then
                        fileDir(0) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(1) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        sspedidos(fileDir(0), fileDir(1), bd(3))
                        My.Computer.FileSystem.DeleteFile(fileDir(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(1))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                End If
            End If


            'SubSet II Devoluciones
            If sSets(1) = True Then

                If ename Like "Devolucion2*" Or ename Like "DevolucionDetalle*" Then
                    If ename Like "Devolucion2*" Then
                        fileDir(2) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(3) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdevolucion(fileDir(2), fileDir(3), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(2))
                        My.Computer.FileSystem.DeleteFile(fileDir(3))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If


                End If
            End If

            'SubSet III Cobranzas
            If sSets(2) = True Then
                If ename Like "Cobranza2*" Or ename Like "CobranzaDetalle*" Or ename Like "CobranzaPago*" Then
                    If ename Like "Cobranza2*" Then
                        fileDir(4) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaDetalle*" Then
                        fileDir(5) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaPago*" Then
                        fileDir(6) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 3 Then
                        sscobranza(fileDir(4), fileDir(5), fileDir(6), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(4))
                        My.Computer.FileSystem.DeleteFile(fileDir(5))
                        My.Computer.FileSystem.DeleteFile(fileDir(6))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                End If
            End If

            'SubSet IV Facturas AutoVenta
            If sSets(2) = True Then
                If ename Like "Factura2*" Or ename Like "FacturaDetalle*" Or ename Like "NotaCredito*" Then
                    If ename Like "Factura2*" Then
                        fileDir(7) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "FacturaDetalle*" Then
                        fileDir(8) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "NotaCredito*" Then
                        fileDir(9) = ruta(3) + "\Importacion\" + ename
                        NC = True
                    End If

                    If cont = 2 Then
                        ssautovta(fileDir(7), fileDir(8), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        NC = False
                    End If
                End If
            End If

            'SubSet V Depositos
            If sSets(2) = True Then
                If ename Like "Deposito2*" Or ename Like "DepositoDetalle*" Then
                    If ename Like "Deposito2*" Then
                        fileDir(10) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "DepositoDetalle*" Then
                        fileDir(11) = ruta(3) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdepositos(fileDir(9), fileDir(10), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        My.Computer.FileSystem.DeleteFile(fileDir(10))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If

                End If
            End If
        End If

    End Sub

    'RUTA(4)
    'BD(4)
    'SUC(4)
    'SSETS (28 A 34)
    Private Sub vigilante5_Created(sender As Object, e As IO.FileSystemEventArgs) Handles vigilante5.Created
            Dim ename As String


        ename = Reemplaza2(e.Name)

        If suc(4) = 1 Then

            'SubSet I Pedidos
            If sSets(0) = True Then

                If ename Like "Pedido2*" Or ename Like "PedidoDetalle*" Then

                    If ename Like "Pedido2*" Then
                        fileDir(0) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(1) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        sspedidos(fileDir(0), fileDir(1), bd(4))
                        My.Computer.FileSystem.DeleteFile(fileDir(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(1))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                End If
            End If


            'SubSet II Devoluciones
            If sSets(1) = True Then

                If ename Like "Devolucion2*" Or ename Like "DevolucionDetalle*" Then
                    If ename Like "Devolucion2*" Then
                        fileDir(2) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    Else
                        fileDir(3) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdevolucion(fileDir(2), fileDir(3), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(2))
                        My.Computer.FileSystem.DeleteFile(fileDir(3))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If


                End If
            End If

            'SubSet III Cobranzas
            If sSets(2) = True Then
                If ename Like "Cobranza2*" Or ename Like "CobranzaDetalle*" Or ename Like "CobranzaPago*" Then
                    If ename Like "Cobranza2*" Then
                        fileDir(4) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaDetalle*" Then
                        fileDir(5) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "CobranzaPago*" Then
                        fileDir(6) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 3 Then
                        sscobranza(fileDir(4), fileDir(5), fileDir(6), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(4))
                        My.Computer.FileSystem.DeleteFile(fileDir(5))
                        My.Computer.FileSystem.DeleteFile(fileDir(6))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                End If
            End If

            'SubSet IV Facturas AutoVenta
            If sSets(2) = True Then
                If ename Like "Factura2*" Or ename Like "FacturaDetalle*" Or ename Like "NotaCredito*" Then
                    If ename Like "Factura2*" Then
                        fileDir(7) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "FacturaDetalle*" Then
                        fileDir(8) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "NotaCredito*" Then
                        fileDir(9) = ruta(4) + "\Importacion\" + ename
                        NC = True
                    End If

                    If cont = 2 Then
                        ssautovta(fileDir(7), fileDir(8), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        NC = False
                    End If
                End If
            End If

            'SubSet V Depositos
            If sSets(2) = True Then
                If ename Like "Deposito2*" Or ename Like "DepositoDetalle*" Then
                    If ename Like "Deposito2*" Then
                        fileDir(10) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    ElseIf ename Like "DepositoDetalle*" Then
                        fileDir(11) = ruta(4) + "\Importacion\" + ename
                        cont = cont + 1
                    End If

                    If cont = 2 Then
                        ssdepositos(fileDir(9), fileDir(10), bd(0))
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        My.Computer.FileSystem.DeleteFile(fileDir(10))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If

                End If
            End If
        End If
    End Sub


#End Region

#Region "Timers"

    Private Sub tiempo_Tick(sender As Object, e As ElapsedEventArgs)
        EscribirLog("Exportación Archivos Prioridad I iniciada a las: " & DateTime.Now, EventLogEntryType.Information)
        tiempo.Enabled = False

        If suc(0) = 1 Then

            Dim conn As New connect(bd(0)) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(0) + "\Exportacion\", ruta(0) + "\Exportacion\", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Lp(conn, ruta(0) + "\Exportacion\", "01", "30", "01", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(0) + "\Exportacion\", "01", "30")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try


            ClearMemory()
            'tiempo.Enabled = True

        End If

        If suc(1) = 1 Then
            Dim conn As New connect(bd(1)) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(1) + "\Exportacion\", ruta(1) + "\Exportacion\", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Lp(conn, ruta(1) + "\Exportacion\", "01", "30", "01", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(1) + "\Exportacion\", "01", "30")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try


            ClearMemory()
            'tiempo.Enabled = True
        End If

        If suc(2) = 1 Then
            Dim conn As New connect(bd(2)) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(2) + "\Exportacion\", ruta(2) + "\Exportacion\", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Lp(conn, ruta(2) + "\Exportacion\", "01", "30", "01", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(2) + "\Exportacion\", "01", "30")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try


            ClearMemory()
            'tiempo.Enabled = True
        End If

        If suc(3) = 1 Then
            Dim conn As New connect(bd(3)) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(3) + "\Exportacion\", ruta(3) + "\Exportacion\", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Lp(conn, ruta(3) + "\Exportacion\", "01", "30", "01", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(3) + "\Exportacion\", "01", "30")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try


            ClearMemory()
            'tiempo.Enabled = True
        End If

        If suc(4) = 1 Then
            Dim conn As New connect(bd(4)) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(4) + "\Exportacion\", ruta(4) + "\Exportacion\", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Lp(conn, ruta(4) + "\Exportacion\", "01", "30", "01", "01")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(4) + "\Exportacion\", "01", "30")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try


            ClearMemory()
            'tiempo.Enabled = True
        End If
        EscribirLog("Exportación Archivos Prioridad I Culminada a las: " & DateTime.Now, EventLogEntryType.Information)
        tiempo.Enabled = True

    End Sub

    Private Sub tiempoII_Tick(sender As Object, e As ElapsedEventArgs)
        tiempoII.Enabled = False
        EscribirLog("Exportación Archivos Prioridad II iniciada a las: " & DateTime.Now, EventLogEntryType.Information)
        If suc(0) = 1 Then

            Dim conn As New connect(bd(0)) 'aclarar instancia SQL

            Try
                conn.conectar2()
                MotNoVis(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotNoVta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotDev(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DE DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                incidencias(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO inCIDENCIAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls1(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls2(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls3(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Bco(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO BANCO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Hist(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO HISTORIA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DetHis(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DETALLE HISTORIA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            ClearMemory()

        End If


        If suc(1) = 1 Then
            Dim conn As New connect(bd(1)) 'aclarar instancia SQL

            Try
                conn.conectar2()
                MotNoVis(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotNoVta(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotDev(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                incidencias(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls1(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls2(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls3(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Bco(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO BANCO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Hist(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO HISTORICA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DetHis(conn, ruta(1) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DETALLE HISTORIA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

        End If

        If suc(2) = 1 Then
            Dim conn As New connect(bd(2)) 'aclarar instancia SQL

            Try
                conn.conectar2()
                MotNoVis(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotNoVta(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotDev(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                incidencias(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO inCIDENCIAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls1(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls2(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls3(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Bco(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO BANCO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Hist(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO HISTORICA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DetHis(conn, ruta(2) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DETALLE HISTORIA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

        End If

        If suc(3) = 1 Then
            Dim conn As New connect(bd(3)) 'aclarar instancia SQL

            Try
                conn.conectar2()
                MotNoVis(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotNoVta(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotDev(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                incidencias(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO inCIDENCIAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls1(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls2(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls3(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Bco(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO BANCO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Hist(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO HISTORICA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DetHis(conn, ruta(3) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DETALLE HISTORIA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try
        End If

        If suc(4) = 1 Then
            Dim conn As New connect(bd(4)) 'aclarar instancia SQL

            Try
                conn.conectar2()
                MotNoVis(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotNoVta(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                MotDev(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                incidencias(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO inCIDENCIAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls1(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls2(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar2()
                Cls3(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Bco(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO BANCO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Hist(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO HISTORICA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                DetHis(conn, ruta(4) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DETALLE HISTORIA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try
        End If

        ClearMemory()
        EscribirLog("Exportación de Archivos Prioridad II Culminada a las: " & DateTime.Now, EventLogEntryType.Information)
        tiempoII.Enabled = True
    End Sub


#End Region

#Region "MétodosCarga"

    Private Sub pedidos(fileDir As String, cmdsql As SqlCommand)
        tabla = txtRead(fileDir)
        cmdsql.CommandText = "pinsertarPedidoVenta"
        cmdsql.Parameters.Clear()

        cmdsql.Parameters.Add("@sDoc_Num", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sdFec_Emis", SqlDbType.SmallDateTime)
        cmdsql.Parameters.Add("@sCo_Cond", SqlDbType.VarChar)
        cmdsql.Parameters("@sCo_Cond").Value = "000002" 'configuracion necesaria en archivo de configuracion
        cmdsql.Parameters.Add("@sStatus", SqlDbType.VarChar)
        cmdsql.Parameters("@sStatus").Value = "0" 'estatus sin procesar
        cmdsql.Parameters.Add("@scampo2", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@scampo3", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@deTotal_Neto", SqlDbType.Decimal)
        cmdsql.Parameters.Add("@deMonto_Imp", SqlDbType.Decimal)
        cmdsql.Parameters.Add("@deMonto_Desc_Glob", SqlDbType.Decimal)
        cmdsql.Parameters.Add("@sCo_Ven", SqlDbType.Char)
        cmdsql.Parameters.Add("@sCo_Cli", SqlDbType.Char)
        cmdsql.Parameters.Add("@sPorc_Desc_Glob", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sDescrip", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@scampo4", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sCo_Tran", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sCo_Mone", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sdFec_Venc", SqlDbType.SmallDateTime)
        cmdsql.Parameters.Add("@sdFec_Reg", SqlDbType.SmallDateTime)
        cmdsql.Parameters.Add("@bAnulado", SqlDbType.Bit)
        cmdsql.Parameters("@bAnulado").Value = False 'no anulado
        cmdsql.Parameters.Add("@bVen_Ter", SqlDbType.Bit)
        cmdsql.Parameters("@bVen_Ter").Value = False
        cmdsql.Parameters.Add("@deTasa", SqlDbType.Decimal)
        cmdsql.Parameters("@deTasa").Value = 1.0
        cmdsql.Parameters.Add("@deMonto_Reca", SqlDbType.Decimal)
        cmdsql.Parameters("@deMonto_Reca").Value = "0.00"
        cmdsql.Parameters.Add("@deTotal_Bruto", SqlDbType.Decimal)
        cmdsql.Parameters.Add("@deMonto_Imp2", SqlDbType.Decimal)
        cmdsql.Parameters("@deMonto_Imp2").Value = "0.00"
        cmdsql.Parameters.Add("deMonto_Imp3", SqlDbType.Decimal)
        cmdsql.Parameters("deMonto_Imp3").Value = "0.00"
        cmdsql.Parameters.Add("@deOtros1", SqlDbType.Decimal)
        cmdsql.Parameters("@deOtros1").Value = "0.00"
        cmdsql.Parameters.Add("@deOtros2", SqlDbType.Decimal)
        cmdsql.Parameters("@deOtros2").Value = "0.00"
        cmdsql.Parameters.Add("@deOtros3", SqlDbType.Decimal)
        cmdsql.Parameters("@deOtros3").Value = "0.00"
        cmdsql.Parameters.Add("@deSaldo", SqlDbType.Decimal)
        cmdsql.Parameters.Add("@bContrib", SqlDbType.Bit)
        cmdsql.Parameters("@bContrib").Value = True
        cmdsql.Parameters.Add("@bImpresa", SqlDbType.Bit)
        cmdsql.Parameters("@bImpresa").Value = False
        cmdsql.Parameters.Add("@sCo_Us_in", SqlDbType.VarChar)
        cmdsql.Parameters("@sCo_Us_in").Value = "ISExportacionH" 'para efectos de prueba
        cmdsql.Parameters.Add("@sN_Control", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sComentario", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sDir_Ent", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sSalestax", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sImpfis", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sImpfisfac", SqlDbType.VarChar)
        cmdsql.Parameters.Add("@sCo_Sucu_in", SqlDbType.VarChar)
        'cmdsql.Parameters.Add("@fe_us_in", SqlDbType.DateTime)
        'cmdsql.Parameters.Add("@co_us_mo", SqlDbType.VarChar)
        'cmdsql.Parameters("@co_us_mo").Value = "ISExportacionH" 'para efectos de prueba
        'cmdsql.Parameters.Add("@fe_us_mo", SqlDbType.DateTime)

        For r = 0 To tabla.Rows.Count - 2
            Try

                'rellenamos el valor de los parámetros
                cmdsql.Parameters("@sDoc_Num").Value = tabla.Rows(r)("column1") 'codPedido
                Fecha = tabla.Rows(r)("column2").ToString 'fecha
                cmdsql.Parameters("@sdFec_Emis").Value = FechaSDT(Fecha)
                cmdsql.Parameters("@sCampo2").Value = tabla.Rows(r)("column3") 'fechaDespacho NO SOLICITADO POR EL PROFIT
                cmdsql.Parameters("@sCampo3").Value = tabla.Rows(r)("column4") 'cantidadTotal NO SOLICITADO POR EL PROFIT
                Dim tot_neto As String = tabla.Rows(r)("column5") 'monto
                'Dim total_netoReng As Decimal = totalRengPed(tabla.Rows(r)("column1").ToString)
                cmdsql.Parameters("@deTotal_Neto").Value = Dec(tot_neto)
                Dim monto_imp As String = tabla.Rows(r)("column6") 'impuesto
                cmdsql.Parameters("@deMonto_Imp").Value = Dec(monto_imp)
                cmdsql.Parameters("@deMonto_Desc_Glob").Value = Dec(tabla.Rows(r)("column7")) 'descuento
                cmdsql.Parameters("@sCo_Ven").Value = tabla.Rows(r)("column8") 'codVendedor
                cmdsql.Parameters("@sCo_Cli").Value = tabla.Rows(r)("column9") 'codCliente
                cmdsql.Parameters("@sPorc_Desc_Glob").Value = tabla.Rows(r)("column10")
                cmdsql.Parameters("@sDescrip").Value = tabla.Rows(r)("column11") 'comentario
                cmdsql.Parameters("@sCampo4").Value = tabla.Rows(r)("column11") 'OrdenCompra NO SOLICITADO POR EL PROFIT
                cmdsql.Parameters("@sCo_Tran").Value = "01" 'inFORMACION GENERICA
                cmdsql.Parameters("@sCo_Mone").Value = "BS" 'inFORMACION GENERICA
                cmdsql.Parameters("@sdFec_Venc").Value = FechaSDT(Fecha)
                cmdsql.Parameters("@sdFec_Reg").Value = FechaSDT(Fecha)
                cmdsql.Parameters("@deTotal_Bruto").Value = Dec(tot_neto) - Dec(monto_imp)
                cmdsql.Parameters("@deSaldo").Value = Dec(tabla.Rows(r)("column5")) '=tot_neto
                'Paremtros adicionales solicitados por el SP
                cmdsql.Parameters("@sN_Control").Value = DBNull.Value
                cmdsql.Parameters("@sComentario").Value = DBNull.Value
                cmdsql.Parameters("@sDir_Ent").Value = DBNull.Value
                cmdsql.Parameters("@sSalestax").Value = DBNull.Value
                cmdsql.Parameters("@sImpfis").Value = DBNull.Value
                cmdsql.Parameters("@sImpfisfac").Value = DBNull.Value
                cmdsql.Parameters("@sCo_Sucu_in").Value = DBNull.Value

                'cmdsql.Parameters("@fe_us_in").Value = FechaDatTime(Fecha)
                'cmdsql.Parameters("@fe_us_mo").Value = FechaDatTime(Fecha)
                'realizamos el alta

                cmdsql.ExecuteNonQuery()

                'logArr(r) = logMsg

            Catch ex As Exception
                EscribirLog("El pedido de venta: " & tabla.Rows(r)("column1") & " no pudo ser cargado por el siguiente error: " & ex.Message, EventLogEntryType.Warning)
            End Try

        Next
        tabla.Clear()

    End Sub

    Private Sub pedidoDetalle(fileDir As String, cmdSQL As SqlCommand, conn As connect, tran As SqlTransaction)
        tabla = txtRead(fileDir)
        cmdSQL.CommandText = "pinsertarRenglonesPedidoVenta_IS"
        cmdSQL.Parameters.Clear()

        cmdSQL.Parameters.Add("@RowGuideVald", SqlDbType.UniqueIdentifier)
        cmdSQL.Parameters.Add("@iReng_Num", SqlDbType.Int)
        cmdSQL.Parameters.Add("@sDoc_Num", SqlDbType.Char)
        cmdSQL.Parameters.Add("@sCo_Art", SqlDbType.Char)
        cmdSQL.Parameters.Add("@sCo_Alma", SqlDbType.Char)  'Pendiente para modificaciones con VentaDirecta
        cmdSQL.Parameters.Add("@deTotal_Art", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deSTotal_Art", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@sCo_Uni", SqlDbType.Char)  'Debe Manejarse Una Sola unidad
        cmdSQL.Parameters.Add("@sCo_Precio", SqlDbType.Char)
        cmdSQL.Parameters.Add("@dePrec_Vta", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_Desc", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@sPorc_Desc", SqlDbType.VarChar)
        cmdSQL.Parameters.Add("@sTipo_Imp", SqlDbType.Char)
        cmdSQL.Parameters.Add("@dePorc_Imp", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@dePorc_Imp2", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@dePorc_Imp3", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_Imp", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_Imp2", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_Imp3", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deReng_Neto", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@dePendiente", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@dePendiente2", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_Desc_Glob", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_reca_Glob", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deOtros1_glob", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deOtros2_glob", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deOtros3_glob", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_imp_afec_glob", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_imp2_afec_glob", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_imp3_afec_glob", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deTotal_Dev", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deMonto_Dev", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@deOtros", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@sCo_Us_in", SqlDbType.Char)
        'Paremetros adicionales solicitados por el SP
        cmdSQL.Parameters.Add("@sTipo_Doc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@gRowguid_Doc", SqlDbType.UniqueIdentifier)
        cmdSQL.Parameters.Add("@sNum_Doc", SqlDbType.VarChar)
        cmdSQL.Parameters.Add("@sComentario", SqlDbType.Char)
        cmdSQL.Parameters.Add("sCo_Sucu_in", SqlDbType.Char)
        cmdSQL.Parameters.Add("@sREVISADO", SqlDbType.Char)
        cmdSQL.Parameters.Add("@sTRASNFE", SqlDbType.Char)

        Dim ctrReng As Boolean = False
        Dim nroDoc As String


        For r = 0 To tabla.Rows.Count - 2
            Try

                Dim Total_Act As Double = 0
                Dim Total_Comp As Double = 0

                Dim tablaAct As DataTable = Stk(tabla.Rows(r)("column2").ToString.Trim, conn, tran, False)
                Dim tablaComp As DataTable = Stk(tabla.Rows(r)("column2").ToString.Trim, conn, tran, True)

                If tablaAct.Rows.Count > 0 Then
                    Total_Act = Dec(tablaAct.Rows(0)("stock"))
                Else
                    Total_Act = 0
                End If

                If tablaComp.Rows.Count > 0 Then
                    Total_Comp = Dec(tablaComp.Rows(0)("stock"))
                Else
                    Total_Comp = 0
                End If

                tablaAct.Clear()
                tablaComp.Clear()


                Dim Total_art As String = Dec(tabla.Rows(r)("column3"))



                'If tablaAct.Rows.Count > 0 Then
                'Total_Act = Dec(tablaAct.Rows(0)("stock"))
                ' Else
                'Total_Act = 0
                ' End If

                ' If tablaComp.Rows.Count > 0 Then
                'Total_Comp = Dec(tablaComp.Rows(0)("stock"))
                ' Else
                ' Total_Comp = 0
                'End If

                'tablaAct.Clear()
                'tablaComp.Clear()

                Dim disp As Double = Total_Act - Total_Comp
                'Dim nvoStk As Double = Total_Comp + Total_art

                If disp > 0 And Total_art <= disp Then

                    'rellenamos el valor de los parámetros
                    cmdSQL.Parameters("@sDoc_Num").Value = tabla.Rows(r)("column1")

                    'Colocar Numeros de Renglon de Acuerdo al Documento al que pertenezcan
                    Dim ArtPrgRow As Integer


                    If (tabla.Rows(r)("column1") = nroDoc) Then
                        'Count = 1
                        'Brake = tabla.Rows(r - 1)("column1")
                        'If tabla.Rows(r)("column1") = nroDoc Then
                        'ArtPrgRow = Count + ArtPrgRow
                        'Else
                        'ArtPrgRow = 1
                        'Count = 0
                        'ctrReng = False
                        'End If
                        ArtPrgRow = ArtPrgRow + 1
                    Else

                        nroDoc = tabla.Rows(r)("column1").ToString
                        ArtPrgRow = 1
                        'ctrReng = True

                    End If


                    cmdSQL.Parameters("@iReng_Num").Value = ArtPrgRow
                    cmdSQL.Parameters("@sCo_Art").Value = tabla.Rows(r)("column2")
                    'parametro para validacion
                    artGral = tabla.Rows(r)("column2")
                    cmdSQL.Parameters("@sCo_Alma").Value = "01"  'Pendiente para modificaciones con VentaDirecta




                    cmdSQL.Parameters("@deTotal_Art").Value = Dec(Total_art)
                    'parametro para validacion
                    cantGral = Dec(Total_art)

                    cmdSQL.Parameters("@deSTotal_Art").Value = 0
                    cmdSQL.Parameters("@sCo_Uni").Value = "FR"  'Debe Manejarse Una Sola unidad LA PRIMARIA PRinCIPAL
                    cmdSQL.Parameters("@sCo_Precio").Value = "01" 'tipo de precio deberia configurarse por archivo, para evitar malas cargas

                    'Consulta de precios para evitar carga de articulos sin precio
                    Dim CMD2 As New SqlCommand
                    CMD2.Connection = conn.connName
                    CMD2.CommandType = CommandType.StoredProcedure
                    CMD2.Transaction = tran
                    CMD2.CommandText = "RepArticuloConPrecio747"
                    CMD2.Parameters.Clear()

                    CMD2.Parameters.Add("@sCo_art_d", SqlDbType.VarChar)
                    CMD2.Parameters.Add("@sCo_art_h", SqlDbType.VarChar)
                    CMD2.Parameters.Add("@sCo_Almacen", SqlDbType.Char)
                    CMD2.Parameters.Add("@sCo_Precio01", SqlDbType.Char)

                    CMD2.Parameters("@sCo_art_d").Value = tabla.Rows(r)("column2")
                    CMD2.Parameters("@sCo_art_h").Value = tabla.Rows(r)("column2")
                    CMD2.Parameters("@sCo_Almacen").Value = "01" 'almacen debe configurarse
                    CMD2.Parameters("@sCo_Precio01").Value = "01" 'tipo de precio deberia configurarse por archivo, para evitar malas cargas

                    tablaTrans.Clear()

                    adp = New SqlDataAdapter(CMD2)
                    adp.Fill(tablaTrans)

                    Dim prec_vta As String = tablaTrans.Rows(0)("Precio01")

                    tablaTrans.Clear()
                    CMD2.Parameters.Clear()

                    Dim CMD3 As New SqlCommand
                    CMD3.Connection = conn.connName
                    CMD3.CommandType = CommandType.StoredProcedure
                    CMD3.Transaction = tran
                    CMD3.CommandText = "ObtenerTasa"
                    CMD3.Parameters.Clear()

                    CMD3.Parameters.Add("@coArt", SqlDbType.VarChar)


                    CMD3.Parameters("@coArt").Value = tabla.Rows(r)("column2")


                    tablaTrans2.Clear()

                    adp = New SqlDataAdapter(CMD3)
                    adp.Fill(tablaTrans2)

                    Dim tsa As Integer = tablaTrans2.Rows(0)("tipo_imp")
                    Dim porcImp As Double

                    tablaTrans2.Clear()
                    CMD3.Parameters.Clear()

                    If tsa <> 1 Then
                        porcImp = 0
                    Else
                        porcImp = 12
                    End If

                    Dim reng_neto As Double = Dec(Total_art) * Dec(prec_vta)

                    cmdSQL.Parameters("@dePrec_Vta").Value = Dec(prec_vta)
                    cmdSQL.Parameters("@deMonto_Desc").Value = Dec(tabla.Rows(r)("column5"))
                    cmdSQL.Parameters("@sPorc_Desc").Value = tabla.Rows(r)("column6")
                    cmdSQL.Parameters("@sTipo_Imp").Value = tsa
                    cmdSQL.Parameters("@dePorc_Imp").Value = porcImp
                    cmdSQL.Parameters("@dePorc_Imp2").Value = 0
                    cmdSQL.Parameters("@dePorc_Imp3").Value = 0
                    cmdSQL.Parameters("@deMonto_Imp").Value = Dec((porcImp / 100) * reng_neto)
                    cmdSQL.Parameters("@deMonto_Imp2").Value = 0
                    cmdSQL.Parameters("@deMonto_Imp3").Value = 0
                    cmdSQL.Parameters("@deReng_Neto").Value = reng_neto
                    cmdSQL.Parameters("@dePendiente").Value = Dec(tabla.Rows(r)("column3"))
                    cmdSQL.Parameters("@dePendiente2").Value = 0
                    'cmdsql.Parameters("@lote_asignado").value= 0)
                    cmdSQL.Parameters("@deMonto_Desc_Glob").Value = Dec(tabla.Rows(r)("column5"))
                    cmdSQL.Parameters("@deMonto_reca_Glob").Value = 0
                    cmdSQL.Parameters("@deOtros1_glob").Value = 0
                    cmdSQL.Parameters("@deOtros2_glob").Value = 0
                    cmdSQL.Parameters("@deOtros3_glob").Value = 0
                    cmdSQL.Parameters("@deMonto_imp_afec_glob").Value = 0
                    cmdSQL.Parameters("@deMonto_imp2_afec_glob").Value = 0
                    cmdSQL.Parameters("@deMonto_imp3_afec_glob").Value = 0
                    cmdSQL.Parameters("@deTotal_Dev").Value = 0
                    cmdSQL.Parameters("@deMonto_Dev").Value = 0
                    cmdSQL.Parameters("@deOtros").Value = 0
                    cmdSQL.Parameters("@sCo_Us_in").Value = "ISExportacionH"
                    'Paremetros adicionales solicitados por el SP
                    cmdSQL.Parameters("@sTipo_Doc").Value = DBNull.Value
                    cmdSQL.Parameters("@gRowguid_Doc").Value = DBNull.Value
                    cmdSQL.Parameters("@sNum_Doc").Value = DBNull.Value
                    cmdSQL.Parameters("@sComentario").Value = DBNull.Value
                    cmdSQL.Parameters("sCo_Sucu_in").Value = DBNull.Value
                    cmdSQL.Parameters("@sREVISADO").Value = DBNull.Value
                    cmdSQL.Parameters("@sTRASNFE").Value = DBNull.Value
                    Unique = Guid.NewGuid
                    cmdSQL.Parameters("@RowGuideVald").Value = Unique

                    'cmdsql.Parameters("@co_us_mo").value= "ISExportacionH")
                    'cmdsql.Parameters("@fe_us_mo").value= FechaDatTime(Fecha))


                    'realizamos el alta

                    cmdSQL.ExecuteNonQuery()

                    'Validacion
                    Modificar_Stock(artGral, Dec(cantGral), conn, tran)

                Else

                    EscribirLog("El articulo: " & tabla.Rows(r)("column2").ToString & " del pedido de venta: " & nroDoc & " no pudo ser cargado por no haber suficiente Stock Disponible", EventLogEntryType.Warning)

                End If

            Catch ex As Exception
                EscribirLog("El articulo: " & tabla.Rows(r)("column2").ToString & " del pedido de venta: " & nroDoc & " no pudo ser cargado por presentarse el siguiente error: " & ex.Message, EventLogEntryType.Warning)

            End Try
        Next

            tabla.Clear()

    End Sub

    Private Sub devolucion(fileDir As String, cmdSQL As SqlCommand)
        Throw New NotImplementedException
    End Sub

    Private Sub devoluciondetalle(fileDir As String, cmdSQL As SqlCommand, conn As connect)
        Throw New NotImplementedException
    End Sub

    Private Sub cobranza(fileDir As String, cmdSQL As SqlCommand)

        tabla = txtRead(fileDir)
        cmdSQL.CommandText = "insertarCobroMovil"
        cmdSQL.Parameters.Clear()


        cmdSQL.Parameters.Add("@codCob", SqlDbType.Char)
        cmdSQL.Parameters.Add("@montoTotal", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@fecha", SqlDbType.SmallDateTime)
        cmdSQL.Parameters.Add("@codCte", SqlDbType.Char)
        cmdSQL.Parameters.Add("@codVen", SqlDbType.Char)
        cmdSQL.Parameters.Add("@aplicada", SqlDbType.Bit) 'campo añadido para validar aplicacion de la cobranza


        For r = 0 To tbl_gral.Rows.Count - 2
            Try

                cmdSQL.Parameters("@codCob").Value = tbl_gral.Rows(r)("column1")
                cmdSQL.Parameters("@montoTotal").Value = Dec(tbl_gral.Rows(r)("column2"))
                cmdSQL.Parameters("@fecha").Value = FechaDatTime(tbl_gral.Rows(r)("column3"))
                cmdSQL.Parameters("@codCte").Value = tbl_gral.Rows(r)("column4")
                cmdSQL.Parameters("@codVen").Value = tbl_gral.Rows(r)("column5")
                cmdSQL.Parameters("@aplicada").Value = False

                cmdSQL.ExecuteNonQuery()

            Catch ex As Exception
                EscribirLog("La cobranza N° : " & tbl_gral.Rows(r)("column1") & " no pudo ser cargado por el siguiente error: " & ex.Message, EventLogEntryType.Warning)
            End Try
        Next
        tabla.Clear()

    End Sub

    Private Sub cobranzadetalle(fileDir As String, cmdSQL As SqlCommand, conn As connect)
        tabla = txtRead(fileDir)
        cmdSQL.CommandText = "insertarCobroMovil_Detalle"
        cmdSQL.Parameters.Clear()



        cmdSQL.Parameters.Add("@codCob", SqlDbType.Char)
        cmdSQL.Parameters.Add("@codDoc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@fecha", SqlDbType.SmallDateTime)
        cmdSQL.Parameters.Add("@mtoAbo", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@sdo", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@dctoPP", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@porDctoPP", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@dctoComerc", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@porDctoComerc", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@tipoDoc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@reng", SqlDbType.Int) ' columna adicional solicitada por el profit


        For r = 0 To tabla.Rows.Count - 2
            'rellenamos el valor de los parámetros

            cmdSQL.Parameters("@codCob").Value = tabla.Rows(r)("column1")
            cmdSQL.Parameters("@codDoc").Value = tabla.Rows(r)("column2")
            cmdSQL.Parameters("@fecha").Value = FechaDatTime(tabla.Rows(r)("column3"))
            cmdSQL.Parameters("@mtoAbo").Value = Dec(tabla.Rows(r)("column4"))
            cmdSQL.Parameters("@sdo").Value = Dec(tabla.Rows(r)("column5"))
            cmdSQL.Parameters("@dctoPP").Value = Dec(tabla.Rows(r)("column6"))
            cmdSQL.Parameters("@porDctoPP").Value = Dec(tabla.Rows(r)("column7"))
            cmdSQL.Parameters("@dctoComerc").Value = Dec(tabla.Rows(r)("column8"))
            cmdSQL.Parameters("@porDctoComerc").Value = Dec(tabla.Rows(r)("column9"))


            Dim TipDoc As String = tabla.Rows(r)("column10")

            Select Case TipDoc
                Case "ND"
                    TipDoc = "N/DB"
                Case "NC"
                    TipDoc = "N/CR"
                Case "FA"
                    TipDoc = "FACT"
                Case "AD"
                    TipDoc = "ADEL"
                Case "AJP"
                    TipDoc = "AJPM"
                Case "AJN"
                    TipDoc = "AJNM"
                Case "CH"
                    TipDoc = "CHEQ"
            End Select

            cmdSQL.Parameters("@tipoDoc").Value = TipDoc

            Dim ArtPrgRow As Integer
            Dim nroDoc As String

            If (tabla.Rows(r)("column1") = nroDoc) Then
                'Count = 1
                'Brake = tabla.Rows(r - 1)("column1")
                'If tabla.Rows(r)("column1") = nroDoc Then
                'ArtPrgRow = Count + ArtPrgRow
                'Else
                'ArtPrgRow = 1
                'Count = 0
                'ctrReng = False
                'End If
                ArtPrgRow = ArtPrgRow + 1
            Else

                nroDoc = tabla.Rows(r)("column1").ToString
                ArtPrgRow = 1
                'ctrReng = True

            End If

            cmdSQL.Parameters("@reng").Value = ArtPrgRow


            cmdSQL.ExecuteNonQuery()

        Next


        tabla.Clear()
    End Sub

    Private Sub cobranzapago(fileDir As String, cmdSQL As SqlCommand, conn As connect)
        tabla = txtRead(fileDir)
        cmdSQL.CommandText = "insertarCobroMovil_Pago"
        cmdSQL.Parameters.Clear()

        cmdSQL.Parameters.Add("@codCobranza", SqlDbType.Char)
        cmdSQL.Parameters.Add("@codDoc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@formPag", SqlDbType.Char)
        cmdSQL.Parameters.Add("@monto", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@codBco", SqlDbType.Char)
        cmdSQL.Parameters.Add("@reng", SqlDbType.Int)

        cmdSQL.Parameters.Add("@montoAfecDoc", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@fechaEmis", SqlDbType.SmallDateTime)
        cmdSQL.Parameters.Add("@codTipDoc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@codDocCanc", SqlDbType.Char)


        For r = 0 To tabla.Rows.Count - 2
            cmdSQL.Parameters("@codCobranza").Value = tabla.Rows(r)("column1")
            cmdSQL.Parameters("@codDoc").Value = tabla.Rows(r)("column2")
            Dim forma_pag As String = tabla.Rows(r)("column3")
            Select Case forma_pag
                Case "Efectivo"
                    forma_pag = "EF"
                Case "Cheque"
                    forma_pag = "CH"
                Case "Depósito"
                    forma_pag = "DP"
                Case "Retención"
                    forma_pag = "RT"
            End Select
            cmdSQL.Parameters("@formPag").Value = forma_pag
            cmdSQL.Parameters("@monto").Value = Dec(tabla.Rows(r)("column4"))
            cmdSQL.Parameters("@montoAfecDoc").Value = Dec(tabla.Rows(r)("column5"))
            cmdSQL.Parameters("@codBco").Value = tabla.Rows(r)("column6")
            cmdSQL.Parameters("@fechaEmis").Value = FechaSDT(tabla.Rows(r)("column7"))
            cmdSQL.Parameters("@codTipDoc").Value = tabla.Rows(r)("column8")
            cmdSQL.Parameters("@codDocCanc").Value = tabla.Rows(r)("column9")

            Dim ArtPrgRow As Integer
            Dim nroDoc As String

            If (tabla.Rows(r)("column1") = nroDoc) Then
                'Count = 1
                'Brake = tabla.Rows(r - 1)("column1")
                'If tabla.Rows(r)("column1") = nroDoc Then
                'ArtPrgRow = Count + ArtPrgRow
                'Else
                'ArtPrgRow = 1
                'Count = 0
                'ctrReng = False
                'End If
                ArtPrgRow = ArtPrgRow + 1
            Else

                nroDoc = tabla.Rows(r)("column1").ToString
                ArtPrgRow = 1
                'ctrReng = True

            End If

            cmdSQL.Parameters("@reng").Value = ArtPrgRow



            cmdSQL.ExecuteNonQuery()

        Next

        tabla.Clear()

    End Sub
    'Determinar Su Real Necesidad
    Private Sub factura(fileDir As String, cmdSQL As SqlCommand)
        Throw New NotImplementedException
    End Sub

    Private Sub facturaDetalle(fileDir As String, cmdSQL As SqlCommand, conn As connect)
        Throw New NotImplementedException
    End Sub

    Private Sub notaCredito(fileDir As String, cmdSQL As SqlCommand, conn As connect)
        Throw New NotImplementedException
    End Sub

    Private Sub deposito(fileDir As String, cmdSQL As SqlCommand)
        tabla = txtRead(fileDir)
        cmdSQL.CommandText = "dbo.insertarDepositoMovil"
        cmdSQL.Parameters.Clear()

        cmdSQL.Parameters.Add("@codDeposito", SqlDbType.Char)
        cmdSQL.Parameters.Add("@codVen", SqlDbType.Char)
        cmdSQL.Parameters.Add("@numDep", SqlDbType.Char)
        cmdSQL.Parameters.Add("@bco", SqlDbType.Char)
        cmdSQL.Parameters.Add("@fecha", SqlDbType.Date)
        cmdSQL.Parameters.Add("@total", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@numCta", SqlDbType.Char)


        For r = 0 To tabla.Rows.Count - 2
            cmdSQL.Parameters("@codDeposito").Value = tabla.Rows(r)("column1")
            cmdSQL.Parameters("@codVen").Value = tabla.Rows(r)("column2")
            cmdSQL.Parameters("@numDep").Value = tabla.Rows(r)("column3")
            cmdSQL.Parameters("@bco").Value = tabla.Rows(r)("column4")
            cmdSQL.Parameters("@fecha").Value = tabla.Rows(r)("column5")
            cmdSQL.Parameters("@total").Value = tabla.Rows(r)("column6")
            cmdSQL.Parameters("@numCta").Value = tabla.Rows(r)("column7")

            cmdSQL.ExecuteNonQuery()

        Next

        tabla.Clear()

    End Sub

    Private Sub depositoDetalle(fileDir As String, cmdSQL As SqlCommand, conn As connect)
        tabla = txtRead(fileDir)
        cmdSQL.CommandText = "dbo.insertarDepositoMovil_detalle"
        cmdSQL.Parameters.Clear()

        cmdSQL.Parameters.Add("@codDeposito", SqlDbType.Char)
        cmdSQL.Parameters.Add("@numDep", SqlDbType.Char)
        cmdSQL.Parameters.Add("@tipo", SqlDbType.Char)
        cmdSQL.Parameters.Add("@numDoc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@monto ", SqlDbType.Decimal)


        For r = 0 To tabla.Rows.Count - 2
            cmdSQL.Parameters("@codDeposito").Value = tabla.Rows(r)("column1")
            cmdSQL.Parameters("@numDep").Value = tabla.Rows(r)("column2")
            cmdSQL.Parameters("@tipo").Value = tabla.Rows(r)("column3")
            cmdSQL.Parameters("@numDoc").Value = tabla.Rows(r)("column4")
            cmdSQL.Parameters("@monto").Value = tabla.Rows(r)("column5")

            cmdSQL.ExecuteNonQuery()
        Next

        tabla.Clear()

    End Sub



#End Region

#Region "MétodosExtracción"
    Public Sub Vendedores(ByVal Conex As connect, Ruta As String)
        '////////////////////////////////////////////////
        '//////////////////Vendedor.txt//////////////////
        '////////////////////////////////////////////////

        adp = New SqlDataAdapter("select co_ven,ven_des,campo1,campo2,inactivo from saVendedor where co_ven not like '%sup%'", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)


        Dim ArrVen(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Ven1 As String = tbl_gral.Rows(r)("co_ven").ToString.Trim
            Ven1 = Delimitador(6, Ven1)
            Dim Ven2 As String = tbl_gral.Rows(r)("ven_des").ToString.Trim
            Ven2 = Delimitador(60, Ven2)

            Dim Ven3Bool As Boolean = tbl_gral.Rows(r)("inactivo")
            Dim Ven3 As String
            Select Case Ven3Bool
                Case True
                    Ven3 = "0"
                Case False
                    Ven3 = "1"
            End Select
            Ven3 = Delimitador(1, Ven3)

            Dim Ven4 As String

            If Not DBNull.Value.Equals(tbl_gral.Rows(r)("campo2")) Then
                Ven4 = tbl_gral.Rows(r)("campo2").ToString.Trim
            Else
                Ven4 = ""
            End If

            Dim Ven5 As String
            If Not DBNull.Value.Equals(tbl_gral.Rows(r)("campo1")) Then
                Ven5 = tbl_gral.Rows(r)("campo1").ToString.Trim
            Else
                Ven5 = ""
            End If



            Dim FilaVen As String
            FilaVen = Ven1 + vbTab + Ven2 + vbTab + Ven3 + vbTab + Ven4 + vbTab + Ven5

            ArrVen(r) = FilaVen
        Next

        'genera Archivo Vendedores

        txt(Ruta & "Vendedor.txt", tbl_gral.Rows.Count, ArrVen)
        tbl_gral.Clear()

    End Sub

    Public Sub TipoNegocio(ByVal Conex As connect, Ruta As String)
        '////////////////////////////////////////////////
        '//////////////////TipoNegocio.txt////////////
        '////////////////////////////////////////////////

        adp = New SqlDataAdapter("select * from saSegmento", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrTp(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Tp1 As String = ((tbl_gral.Rows(r)("co_seg")).ToString).Trim
            Tp1 = Delimitador(6, Tp1)
            Dim Tp2 As String = ((tbl_gral(r)("seg_des").ToString).Trim)
            Tp2 = Delimitador(60, Tp2)
            Dim Tp3 As String = "1"
            Tp3 = Delimitador(1, Tp3)
            Dim FilaTp As String = Tp1 + vbTab + Tp2 + vbTab + Tp3
            ArrTp(r) = FilaTp
        Next


        'Genera archivo TipoNegocio
        txt(Ruta & "TipoNegocio.txt", tbl_gral.Rows.Count, ArrTp)
        tbl_gral.Clear()
    End Sub

    Public Sub Clientes(ByVal Conex As connect, Ruta As String, Ruta2 As String, Tcl As String)
        adp = New SqlDataAdapter("select cli.co_cli,cli.cli_des,cli.telefonos,cli.direc1,cli.rif,cli.fe_us_mo,cli.inactivo,cli.mont_cre,cli.co_ven,cp.dias_cred,cli.respons,cli.campo3,cli.ciudad,cli.co_seg,cli.sincredito,cli.email,cli.tip_cli,cli.campo4,cli.lunes,cli.martes,cli.miercoles,cli.jueves,cli.viernes,cli.sabado,cli.domingo from sacliente Cli inner join saCondicionPago CP on Cli.cond_pag = CP.co_cond", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)
        adp = New SqlDataAdapter("select * from v_saDocumentoVenta where (co_tipo_doc = 'FACT' or co_tipo_doc='N/CR')and(saldo>0)", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tablaTrans)

        Dim Arreglo(tbl_gral.Rows.Count) As String
        Dim ArrCp(tbl_gral.Rows.Count) As String

        Dim r As Integer
        Dim TBL_Alt As DataTable
        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Cte1 As String = ((tbl_gral.Rows(r)("co_cli").ToString).Trim)
            Cte1 = Delimitador(9, Cte1)
            'Parametros para la funcion
            Dim Cte As String = CStr(Cte1)
            Dim Filtro As String = Convert.ToString(Cte)
            TBL_Alt = SelectDataTable(tablaTrans, Filtro)
            'Continuamos con cada campo independiente
            Dim Cte2 As String = ((tbl_gral.Rows(r)("cli_des").ToString).Trim)
            Cte2 = Delimitador(60, Cte2)
            Dim Cte3 As String = ((tbl_gral.Rows(r)("telefonos").ToString).Trim)
            Cte3 = Delimitador(30, Cte3)
            Dim Cte4 As String = ((tbl_gral.Rows(r)("direc1").ToString).Trim)
            Cte4 = Delimitador(255, Cte4)
            Dim Cte5 As String = ((tbl_gral.Rows(r)("rif").ToString).Trim)
            Cte5 = Delimitador(20, Cte5)

            Dim Cte6_DT As DateTime = tbl_gral.Rows(r)("fe_us_mo")
            Dim Cte6 As String
            Cte6_DT = DateTime.Parse(Cte6_DT)
            Cte6 = Cte6_DT.ToString("yyyyMMdd hhmmss")

            Dim Cte7_Bool As Boolean = ((tbl_gral.Rows(r)("inactivo").ToString).Trim)
            Dim Cte7 As String
            If Cte7_Bool = True Then
                Cte7 = "0"
            Else
                Cte7 = "1"
            End If
            Cte7 = Delimitador(1, Cte7)
            Dim Cte8 As String = ((tbl_gral.Rows(r)("mont_cre").ToString).Trim)
            Cte8 = Delimitador(19, Cte8)
            Dim Bandera As Boolean = False
            SaldoD = 0
            SaldoC = 0
            If TBL_Alt.Rows.Count <> 0 Then
                For r1 As Integer = 0 To TBL_Alt.Rows.Count - 1
                    Dim CustCodCli As String = (tbl_gral.Rows(r)("co_cli").ToString).Trim
                    Dim CustCodAB As String = (TBL_Alt.Rows(r1)("co_cli").ToString).Trim
                    Dim saldo As Double = (TBL_Alt.Rows(r1)("saldo"))
                    If CustCodAB = CustCodCli Then
                        Dim TipDoc As String = (TBL_Alt.Rows(r1)("co_tipo_doc").ToString).Trim
                        Select Case TipDoc
                            Case "FACT"
                                SaldoD = SaldoD + saldo
                            Case "N/CR"
                                SaldoC = SaldoC + saldo
                        End Select
                        Bandera = True
                    End If
                Next
            Else
                Bandera = False
            End If
            TBL_Alt.Clear()

            Dim Cte9 As String
            If Bandera = True Then
                Cte9 = (SaldoD - SaldoC).ToString
            Else
                Cte9 = "0.00"
            End If
            Cte9 = Delimitador(19, Cte9)
            Dim Cte10 As String = ((tbl_gral.Rows(r)("co_ven").ToString).Trim)
            Cte10 = Delimitador(6, Cte10)
            Dim Cte11 As String = ((tbl_gral.Rows(r)("dias_cred").ToString).Trim)
            Cte11 = Delimitador(3, Cte11)
            Dim Cte12 As String = ((tbl_gral.Rows(r)("respons").ToString).Trim)
            Cte12 = Delimitador(50, Cte12)
            Dim Cte13 As String = ((tbl_gral.Rows(r)("ciudad").ToString).Trim)
            Cte13 = Delimitador(30, Cte13)
            Dim Cte14 As String = ((tbl_gral.Rows(r)("ciudad").ToString).Trim)
            Cte14 = Delimitador(30, Cte14)
            Dim Cte15 As String = ((tbl_gral.Rows(r)("co_seg").ToString).Trim)
            Cte15 = Delimitador(100, Cte15)
            Dim Cte16_Bool As Boolean = tbl_gral.Rows(r)("sincredito")
            Dim Cte16 As String
            If Cte16_Bool <> True Then
                Cte16 = "2"
            Else
                Cte16 = "0"
            End If
            Cte16 = Delimitador(1, Cte16)
            Dim Cte17 As String = ((tbl_gral.Rows(r)("email").ToString).Trim)
            Cte17 = Delimitador(50, Cte17)
            Dim Cte18 As String = Tcl  '((TBL_GRAL.Rows(r)("tip_cli").Tostring).Trim)
            Cte18 = Delimitador(10, Cte18)
            Dim Fila As String
            Fila = Cte1 + vbTab + Cte2 + vbTab + Cte3 + vbTab + Cte4 + vbTab + Cte5 + vbTab + Cte6 + vbTab + Cte7 + vbTab + Cte8 + vbTab + Cte9 + _
                vbTab + Cte10 + vbTab + Cte11 + vbTab + Cte12 + vbTab + Cte13 + vbTab + Cte14 + vbTab + Cte15 + vbTab + Cte16 + vbTab + Cte2 + vbTab + Cte17 + vbTab + Cte18
            Fila = Fila.Replace(",", ".")
            Arreglo(r) = Fila

            Dim Cp1 As String
            Dim Dia(6) As Boolean
            Dim DiaV As Integer
            Dia(0) = tbl_gral.Rows(r)("lunes")
            Dia(1) = tbl_gral.Rows(r)("martes")
            Dia(2) = tbl_gral.Rows(r)("miercoles")
            Dia(3) = tbl_gral.Rows(r)("jueves")
            Dia(4) = tbl_gral.Rows(r)("viernes")
            Dia(5) = tbl_gral.Rows(r)("sabado")
            Dia(6) = tbl_gral.Rows(r)("domingo")
            For i As Integer = 0 To 6
                If Dia(i) = True Then
                    DiaV = i
                End If
            Next
            Select Case DiaV
                Case 0
                    Cp1 = "Lunes"
                Case 1
                    Cp1 = "Martes"
                Case 2
                    Cp1 = "Miercoles"
                Case 3
                    Cp1 = "Jueves"
                Case 4
                    Cp1 = "Viernes"
                Case 5
                    Cp1 = "Sabado"
                Case 6
                    Cp1 = "Domingo"
            End Select
            Cp1 = Delimitador(20, Cp1)
            'Orden de la Visita
            Dim Cp2 As String = ((tbl_gral.Rows(r)("campo4").ToString).Trim)
            Cp2 = Delimitador(2, Cp2)
            Dim FilaCp As String
            FilaCp = Cte1 + vbTab + Cp1 + vbTab + Cp2
            ArrCp(r) = FilaCp
        Next

        'Genera archivo Cliente
        txt(Ruta & "Cliente.txt", tbl_gral.Rows.Count, Arreglo)
        'Genera archivo PlanCliente
        txt(Ruta2 & "ClientePLanificacion.txt", tbl_gral.Rows.Count, ArrCp)

        tbl_gral.Clear()
        tablaTrans.Clear()
    End Sub

    Private Sub MotNoVis(ByVal Conex As connect, Ruta As String)
        '////////////////////////////////////////////////
        '//////////////////MotivoNoVisita.txt////////////
        '////////////////////////////////////////////////
        adp = New SqlDataAdapter("select * from ismotivonovisita", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrMnV(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Mnv1 As String = tbl_gral.Rows(r)("CodMotNoVis").ToString
            Mnv1 = Delimitador(20, Mnv1)
            Dim Mnv2 As String = tbl_gral.Rows(r)("descripcion").ToString
            Mnv2 = Delimitador(50, Mnv2)
            Dim Mnv3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim Mnv3 As String
            If Mnv3_Bool = True Then
                Mnv3 = "0"
            Else
                Mnv3 = "1"
            End If
            Mnv3 = Delimitador(1, Mnv3)

            Dim FilaMnv As String = Mnv1 + vbTab + Mnv2 + vbTab + Mnv3
            ArrMnV(r) = FilaMnv
        Next
        'Genera Archivo MotivosNoVisita
        txt(Ruta & "MotivoNoVisita.txt", tbl_gral.Rows.Count, ArrMnV)

        tbl_gral.Clear()
    End Sub

    Private Sub MotNoVta(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from ismotivonoventa", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrMnVta(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Mnv1 As String = tbl_gral.Rows(r)("CodMotNoVta").ToString
            Mnv1 = Delimitador(20, Mnv1)
            Dim Mnv2 As String = tbl_gral.Rows(r)("descripcion").ToString
            Mnv2 = Delimitador(50, Mnv2)
            Dim Mnv3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim Mnv3 As String
            If Mnv3_Bool = True Then
                Mnv3 = "0"
            Else
                Mnv3 = "1"
            End If
            Mnv3 = Delimitador(1, Mnv3)

            Dim FilaMnVta As String = Mnv1 + vbTab + Mnv2 + vbTab + Mnv3
            ArrMnVta(r) = FilaMnVta
        Next
        'Genera Archivo MotivosNoVenta
        txt(Ruta & "MotivoNoVenta.txt", tbl_gral.Rows.Count, ArrMnVta)

        tbl_gral.Clear()
    End Sub

    Private Sub MotDev(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from motivoDev", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrDev(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim mdev1 As String = tbl_gral.Rows(r)("CodMotDev").ToString
            mdev1 = Delimitador(20, mdev1)
            Dim mdev2 As String = tbl_gral.Rows(r)("descripcion").ToString
            mdev2 = Delimitador(50, mdev2)
            Dim mdev3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim mdev3 As String
            If mdev3_Bool = True Then
                mdev3 = "0"
            Else
                mdev3 = "1"
            End If
            mdev3 = Delimitador(1, mdev3)

            Dim Filamdevta As String = mdev1 + vbTab + mdev2 + vbTab + mdev3
            ArrDev(r) = Filamdevta
        Next
        'Genera Archivo MotivosDevolucion
        txt(Ruta & "MotivoDevolucion.txt", tbl_gral.Rows.Count, ArrDev)

        tbl_gral.Clear()
    End Sub

    Private Sub incidencias(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from incidencias", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim Arrinc(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim inc1 As String = tbl_gral.Rows(r)("Codinc").ToString
            inc1 = Delimitador(20, inc1)
            Dim inc2 As String = tbl_gral.Rows(r)("descripcion").ToString
            inc2 = Delimitador(50, inc2)
            Dim inc3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim inc3 As String
            If inc3_Bool = True Then
                inc3 = "0"
            Else
                inc3 = "1"
            End If
            inc3 = Delimitador(1, inc3)

            Dim Filaincta As String = inc1 + vbTab + inc2 + vbTab + inc3
            Arrinc(r) = Filaincta
        Next
        'Genera Archivo MotivosDevolucion
        txt(Ruta & "VisitaMotivo.txt", tbl_gral.Rows.Count, Arrinc)

        tbl_gral.Clear()
    End Sub

    Private Sub Cls1(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from ISClasificacion1_Proveedores", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrCl1(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Cl1_1 As String = tbl_gral.Rows(r)("CodClasificacion").ToString
            Cl1_1 = Delimitador(18, Cl1_1)
            Dim Cl1_2 As String = tbl_gral.Rows(r)("descripcion").ToString
            Cl1_2 = Delimitador(50, Cl1_2)
            Dim Cl1_3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim Cl1_3 As String
            If Cl1_3_Bool = True Then
                Cl1_3 = "0"
            Else
                Cl1_3 = "1"
            End If
            Cl1_3 = Delimitador(1, Cl1_3)

            Dim FilaCl1 As String = Cl1_1 + vbTab + Cl1_2 + vbTab + Cl1_3
            ArrCl1(r) = FilaCl1
        Next

        'Genera Archivo Clasificacion1
        txt(Ruta & "Clasificacion1.txt", tbl_gral.Rows.Count, ArrCl1)

        tbl_gral.Clear()
    End Sub

    Private Sub Cls2(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from ISClasificacion2_Linea", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrCl2(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Cl2_1 As String = tbl_gral.Rows(r)("CodClasificacion2").ToString
            Cl2_1 = Delimitador(18, Cl2_1)
            Dim Cl2_2 As String = tbl_gral.Rows(r)("descripcion").ToString
            Cl2_2 = Delimitador(50, Cl2_2)
            Dim Cl2_3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim Cl2_3 As String
            If Cl2_3_Bool = True Then
                Cl2_3 = "0"
            Else
                Cl2_3 = "1"
            End If
            Cl2_3 = Delimitador(1, Cl2_3)
            Dim Cl2_4 As String = tbl_gral.Rows(r)("CodClasificacion").ToString
            Cl2_4 = Delimitador(18, Cl2_4)



            Dim FilaCl2 As String = Cl2_1 + vbTab + Cl2_4 + vbTab + Cl2_2 + vbTab + Cl2_3
            ArrCl2(r) = FilaCl2
        Next

        'Genera Archivo Clasificacion2
        txt(Ruta & "Clasificacion2.txt", tbl_gral.Rows.Count, ArrCl2)

        tbl_gral.Clear()
    End Sub

    Private Sub Cls3(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from ISClasificacion3_SubLinea", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrCl3(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Cl3_1 As String = tbl_gral.Rows(r)("CodClasificacion3").ToString
            Cl3_1 = Delimitador(18, Cl3_1)
            Dim Cl3_2 As String = tbl_gral.Rows(r)("descripcion").ToString
            Cl3_2 = Delimitador(50, Cl3_2)
            Dim Cl3_3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim Cl3_3 As String
            If Cl3_3_Bool = True Then
                Cl3_3 = "0"
            Else
                Cl3_3 = "1"
            End If
            Cl3_3 = Delimitador(1, Cl3_3)
            Dim Cl3_4 As String = tbl_gral.Rows(r)("CodClasificacion2").ToString
            Cl3_4 = Delimitador(18, Cl3_4)



            Dim FilaCl3 As String = Cl3_1 + vbTab + Cl3_4 + vbTab + Cl3_2 + vbTab + Cl3_3
            ArrCl3(r) = FilaCl3
        Next

        'Genera Archivo Clasificacion 3
        txt(Ruta & "Clasificacion3.txt", tbl_gral.Rows.Count, ArrCl3)
        tbl_gral.Clear()
    End Sub

    Private Sub Lp(ByVal Conex As connect, Ruta As String, linead As String, lineah As String, almcn As String, prec As String)

        adp = New SqlDataAdapter("RepArticuloConCostoYPrecio", Conex.connName)
        adp.SelectCommand.CommandType = CommandType.StoredProcedure
        adp.SelectCommand.Parameters.Add("@sCo_Almacen", SqlDbType.Char)
        adp.SelectCommand.Parameters.Add("@sCo_Precio01", SqlDbType.Char)
        adp.SelectCommand.Parameters.Add("@sCo_Linea_d", SqlDbType.Char)
        adp.SelectCommand.Parameters.Add("@sCo_Linea_h", SqlDbType.Char)
        adp.SelectCommand.Parameters("@sCo_Linea_d").Value = linead
        adp.SelectCommand.Parameters("@sCo_Linea_h").Value = lineah
        adp.SelectCommand.Parameters("@sCo_Almacen").Value = almcn
        adp.SelectCommand.Parameters("@sCo_Precio01").Value = prec
        adp.Fill(tbl_gral)


        Dim ArrLp(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Lp1 As String = tbl_gral.Rows(r)("pre01").ToString.Trim
            Lp1 = Delimitador(10, Lp1)
            Dim Lp2 As String = tbl_gral.Rows(r)("co_art").ToString.Trim
            Lp2 = Delimitador(30, Lp2)
            Dim Lp3_doub As Double = tbl_gral.Rows(r)("ultimo")
            Dim Lp3 As String = Delimitador(21, Format(Lp3_doub, "###,##0.00"))
            Dim Lp4_doub As Double = tbl_gral.Rows(r)("precio1")
            Dim Lp4 As String = Delimitador(21, Format(Lp4_doub, "###,##0.00"))
            Dim Lp5 As String = "FR"
            Lp5 = Delimitador(20, Lp5)

            Dim FilaLp As String = Lp1 + vbTab + Lp2 + vbTab + Lp3 + vbTab + Lp4 + vbTab + Lp5
            FilaLp = Replace(FilaLp, ",", ".")
            ArrLp(r) = FilaLp
        Next
        'genera Archivo ListaPrecio
        txt(Ruta & "ListaPrecio.txt", tbl_gral.Rows.Count, ArrLp)

        tbl_gral.Clear()
    End Sub

    Private Sub UndMed(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from saUnidad", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrUm(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Um1 As String = (tbl_gral.Rows(r)("co_uni").ToString).Trim
            Um1 = Delimitador(20, Um1)
            Dim Um2 As String = (tbl_gral.Rows(r)("co_uni").ToString).Trim
            Um2 = Delimitador(50, Um2)
            Dim Um3 As String = "1"
            Um3 = Delimitador(1, Um3)
            Dim FilaUm As String = Um1 + vbTab + Um2 + vbTab + Um3
            ArrUm(r) = FilaUm

        Next

        'genera Archivo UnidadMedida
        txt(Ruta & "UnidadMedida.txt", tbl_gral.Rows.Count, ArrUm)

        tbl_gral.Clear()
    End Sub

    Private Sub Sku(ByVal Conex As connect, Ruta As String, linead As String, lineah As String)
        Dim Consulta As String = "select a.co_art,dc.desc_abrev,a.anulado,a.tipo_imp,au.equivalencia,a.peso,a.campo4,a.campo5,a.campo6,au.co_uni from saarticulo A inner join saArtUnidad AU on a.co_art =au.co_art inner join DescripcionCorta dc on a.co_art=dc.co_art where au.uni_principal =0 and au.co_uni='cr' and (a.co_lin>='" + linead + "' and a.co_lin<='" + lineah + "')"
        adp = New SqlDataAdapter(Consulta, Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrPr(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Pr1 As String = tbl_gral.Rows(r)("co_art").ToString.Trim
            Pr1 = Delimitador(30, Pr1)
            Dim Pr2 As String = tbl_gral.Rows(r)("desc_abrev").ToString.Trim
            Pr2 = Delimitador(35, Pr2)
            Dim Pr3_bool As Boolean = tbl_gral.Rows(r)("anulado")
            Dim Pr3 As String
            If Pr3_bool = False Then
                Pr3 = "1"
            Else
                Pr3 = "0"
            End If
            Pr3 = Delimitador(1, Pr3)
            Dim Pr4_int As Integer = tbl_gral.Rows(r)("tipo_imp")
            Dim Pr4 As String
            If Pr4_int = 5 Or Pr4_int = 7 Then
                Pr4 = "0"
            Else
                Pr4 = "1"
            End If
            Pr4 = Delimitador(1, Pr4)
            Dim Pr5 As String = "0"
            Pr5 = Delimitador(1, Pr5)
            Dim Pr6_int As Integer = tbl_gral.Rows(r)("equivalencia").ToString.Trim
            Dim Pr6 As String
            Pr6 = Delimitador(4, Pr6_int)
            Dim Pr7_Dou As Double = tbl_gral.Rows(r)("peso")
            Dim Pr7_Calc As Double = Pr6_int * Pr7_Dou
            Dim Pr7 As String
            Pr7 = Delimitador(21, Format(Pr7_Calc, "##,##0.00"))
            Dim Pr8 As String = tbl_gral.Rows(r)("campo4").ToString.Trim
            Pr8 = Delimitador(50, Pr8)
            Dim Pr9 As String = tbl_gral.Rows(r)("campo5").ToString.Trim
            Pr9 = Delimitador(50, Pr9)
            Dim Pr10 As String = tbl_gral.Rows(r)("campo6").ToString.Trim
            Pr10 = Delimitador(50, Pr10)
            Dim Pr11 As String = "FR"
            Pr11 = Delimitador(20, Pr11)
            Dim Pr12 As String = "0"
            Dim Pr13 As String = "                                                  "


            Dim FilaPr As String
            FilaPr = Pr1 + vbTab + Pr2 + vbTab + Pr3 + vbTab + Pr4 + vbTab + Pr5 + vbTab + Pr6 + vbTab + Pr7 + vbTab + Pr8 + vbTab + Pr9 + vbTab + Pr10 + _
                 vbTab + Pr13 + vbTab + Pr11 + vbTab + Pr12
            FilaPr = Replace(FilaPr, ",", ".")
            ArrPr(r) = FilaPr
        Next

        'genera Archivo Producto
        txt(Ruta & "Producto.txt", tbl_gral.Rows.Count, ArrPr)

        tbl_gral.Clear()
    End Sub

    Private Sub Bco(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from saBanco", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrBco(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Bco1 As String = tbl_gral.Rows(r)("co_ban").ToString.Trim
            Bco1 = Delimitador(6, Bco1)
            Dim Bco2 As String = tbl_gral.Rows(r)("des_ban").ToString.Trim
            Bco2 = Delimitador(30, Bco2)
            Dim Bco3 As String = "1"

            Dim FilaBco As String = Bco1 + vbTab + Bco2 + vbTab + Bco3

            ArrBco(r) = FilaBco
        Next

        'genera Archivo ListaPrecio
        txt(Ruta & "Banco.txt", tbl_gral.Rows.Count, ArrBco)

        tbl_gral.Clear()
    End Sub

    Private Sub Hist(ByVal Conex As connect, Ruta As String)
        'Equalizer(conex.connname)
        adp = New SqlDataAdapter("select * from Historia", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrHis(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim His1 As String = tbl_gral.Rows(r)("doc_num").ToString.Trim
            His1 = Delimitador(30, His1)
            Dim His2_int As Integer = tbl_gral.Rows(r)("campo1")
            Dim His2 As String
            Select Case His2_int
                Case 1
                    His2 = "PE"
                Case 2
                    His2 = "FA"
                Case 3
                    His2 = "NE"
            End Select
            His2 = Delimitador(2, His2)
            Dim His3 As String = tbl_gral.Rows(r)("co_cli").ToString.Trim
            His3 = Delimitador(9, His3)
            Dim His4_DT As DateTime = tbl_gral.Rows(r)("fec_emis")
            Dim His4 As String
            His4_DT = DateTime.Parse(His4_DT)
            His4 = His4_DT.ToString("yyyyMMdd hhmmss")
            Dim His5_Doub As Double = tbl_gral.Rows(r)("total_neto")
            Dim His5 As String
            His5 = Delimitador(21, Format(His5_Doub, "##,##0.00"))
            Dim His6_Doub As Double = tbl_gral.Rows(r)("monto_desc_glob")
            Dim His6 As String
            His6 = Delimitador(21, Format(His6_Doub, "##,##0.00"))
            Dim His7_Doub As Double = tbl_gral.Rows(r)("monto_imp")
            Dim His7 As String
            His7 = Delimitador(21, Format(His7_Doub, "##,##0.00"))

            Dim FilaHis As String
            FilaHis = His1 + vbTab + His2 + vbTab + His3 + vbTab + His4 + vbTab + His5 + vbTab + His6 + vbTab + His7
            FilaHis = Replace(FilaHis, ".", "")
            FilaHis = Replace(FilaHis, ",", ".")
            ArrHis(r) = FilaHis
        Next

        'genera Archivo Historia
        txt(Ruta & "Historia.txt", tbl_gral.Rows.Count, ArrHis)

        tbl_gral.Clear()
    End Sub

    Private Sub DetHis(ByVal Conex As connect, Ruta As String)
        'Equalizer(conex.connname)
        adp = New SqlDataAdapter("Select * from HistoriaDetalle ", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrDH(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Dh1 As String = tbl_gral.Rows(r)("doc_num").ToString.Trim
            Dh1 = Delimitador(30, Dh1)
            Dim Dh2_int As Integer = tbl_gral.Rows(r)("campo1")
            Dim Dh2 As String
            Select Case Dh2_int
                Case 1
                    Dh2 = "PE"
                Case 2
                    Dh2 = "FA"
                Case 3
                    Dh2 = "NE"
            End Select
            Dh2 = Delimitador(2, Dh2)
            Dim Dh3 As String = tbl_gral.Rows(r)("co_art").ToString.Trim
            Dh3 = Delimitador(30, Dh3)
            Dim Dh4 As String = tbl_gral.Rows(r)("co_uni").ToString.Trim
            Dh4 = Delimitador(20, Dh4)
            Dim Dh5_Doub As Double = tbl_gral.Rows(r)("total_art")
            Dim Dh5 As String = Delimitador(21, Format(Dh5_Doub, "##,##0.00"))
            Dim Dh6_Doub As Double = tbl_gral.Rows(r)("prec_vta")
            Dim Dh6 As String = Delimitador(21, Format(Dh6_Doub, "##,##0.00"))
            Dim Dh7_Doub As Double = tbl_gral.Rows(r)("reng_neto")
            Dim Dh7 As String = Delimitador(21, Format(Dh7_Doub, "##,##0.00"))
            Dim Dh8_Doub As Double = tbl_gral.Rows(r)("monto_desc")
            Dim Dh8 As String = Delimitador(21, Format(Dh8_Doub, "##,##0.00"))
            Dim Dh9_Doub As Double = tbl_gral.Rows(r)("monto_imp")
            Dim Dh9 As String = Delimitador(21, Format(Dh9_Doub, "##,##0.00"))
            Dim Dh10 As String
            Dim Dh11 As String
            Dim Dh11_str As String
            If Not DBNull.Value.Equals(tbl_gral.Rows(r)("tipo_doc")) Then
                Dh10 = tbl_gral.Rows(r)("num_doc").ToString.Trim
                Dh11_str = tbl_gral.Rows(r)("tipo_doc").ToString.Trim
                Select Case Dh11_str
                    Case "PCLI"
                        Dh11 = "PE"
                    Case "FACT"
                        Dh11 = "FA"
                    Case "NENT"
                        Dh11 = "NE"
                End Select
            Else
                Dh10 = "N/A"
                Dh11 = "NA"
            End If
            Dh10 = Delimitador(30, Dh10)
            Dh11 = Delimitador(2, Dh11)

            Dim FilaDh As String = Dh1 + vbTab + Dh2 + vbTab + Dh3 + vbTab + Dh4 + vbTab + Dh5 + vbTab + Dh6 + vbTab + _
                Dh7 + vbTab + Dh8 + vbTab + Dh9 + vbTab + Dh10 + vbTab + Dh11

            FilaDh = Replace(FilaDh, ".", "")
            FilaDh = Replace(FilaDh, ",", ".")
            ArrDH(r) = FilaDh
        Next

        'genera Archivo HistoriaDetalle
        txt(Ruta & "HistoriaDetalle.txt", tbl_gral.Rows.Count, ArrDH)

        tbl_gral.Clear()
    End Sub

    Private Sub Almcn(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("RepStockArticulosxAlmacen", Conex.connName)
        adp.SelectCommand.CommandType = CommandType.StoredProcedure
        adp.SelectCommand.Parameters.Add("@sCo_Almacen_d", SqlDbType.Char)
        adp.SelectCommand.Parameters.Add("@sCo_Almacen_h", SqlDbType.Char)
        adp.SelectCommand.Parameters.Add("@sTipoStock", SqlDbType.Char)
        adp.SelectCommand.Parameters.Add("@sCo_NivelStock", SqlDbType.Char)
        adp.SelectCommand.Parameters("@sCo_Almacen_d").Value = "01"
        adp.SelectCommand.Parameters("@sCo_Almacen_h").Value = "01"
        adp.SelectCommand.Parameters("@sTipoStock").Value = "DIS"
        adp.SelectCommand.Parameters("@sCo_NivelStock").Value = "MAY"
        adp.Fill(tbl_gral)

        Dim ArrAlm(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Alm1 As String = tbl_gral.Rows(r)("co_alma").ToString.Trim
            Alm1 = Delimitador(20, Alm1)
            Dim Alm2 As String = tbl_gral.Rows(r)("co_art").ToString.Trim
            Alm2 = Delimitador(30, Alm2)
            Dim Alm3 As String = tbl_gral.Rows(r)("co_uni").ToString.Trim
            Alm3 = Delimitador(20, Alm3)
            Dim Alm4_Doub As Double = tbl_gral.Rows(r)("StockActual")
            Dim Alm4 As String = Delimitador(21, Format(Alm4_Doub, "##,##0.00"))

            Dim FilaAlm As String
            FilaAlm = Alm1 + vbTab + Alm2 + vbTab + Alm3 + vbTab + Alm4
            FilaAlm = Replace(FilaAlm, ".", "")
            FilaAlm = Replace(FilaAlm, ",", ".")

            ArrAlm(r) = FilaAlm

        Next
        'genera Archivo Almacen
        txt(Ruta & "Almacen.txt", tbl_gral.Rows.Count, ArrAlm)

        tbl_gral.Clear()
    End Sub

    Private Sub Docs(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from saDocumentoVenta ", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)
        Dim ArrDoc(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Doc1 As String = tbl_gral.Rows(r)("nro_doc").ToString.Trim
            Doc1 = Delimitador(14, Doc1)
            Dim Doc2 As String = tbl_gral.Rows(r)("co_cli").ToString.Trim
            Doc2 = Delimitador(9, Doc2)
            Dim TP As String = tbl_gral.Rows(r)("co_tipo_doc").ToString.Trim
            Dim Doc3 As String
            Select Case TP
                Case "AJNM", "IVAN", "AJNA"
                    Doc3 = "AJN"
                Case "AJPM", "IVAP", "AJPA"
                    Doc3 = "AJP"
                Case "N/DB"
                    Doc3 = "ND"
                Case "N/CR"
                    Doc3 = "NC"
                Case "FACT"
                    Doc3 = "FA"
                Case "ADEL"
                    Doc3 = "AD"
                Case "CHEQ"
                    Doc3 = "CH"
            End Select
            Doc3 = Delimitador(3, Doc3)
            Dim Sdo As Double
            Dim Doc7, Doc4 As String
            Sdo = tbl_gral.Rows(r)("saldo")
            If Sdo = 0 Then
                Doc4 = "1"
            Else
                Doc4 = "0"
            End If
            Doc7 = Delimitador(21, Format(Sdo, "##,##0.00"))
            Dim Doc5_DT As DateTime = tbl_gral.Rows(r)("fec_emis")
            Dim Doc5 As String
            Doc5_DT = DateTime.Parse(Doc5_DT)
            Doc5 = Doc5_DT.ToString("yyyyMMdd hhmmss")
            Dim Doc6_DT As DateTime = tbl_gral.Rows(r)("fec_venc")
            Dim Doc6 As String
            Doc6_DT = DateTime.Parse(Doc6_DT)
            Doc6 = Doc6_DT.ToString("yyyyMMdd hhmmss")
            Dim Doc8_Doub As Double = tbl_gral.Rows(r)("total_neto")
            Dim Doc8 As String = Delimitador(21, Format(Doc8_Doub, "##,##0.00"))
            Dim Doc9_Bool As Boolean = tbl_gral.Rows(r)("anulado")
            Dim Doc9 As String
            If Doc9_Bool = True Then
                Doc9 = "1"
            Else
                Doc9 = "0"
            End If
            Doc9 = Delimitador(1, Doc9)

            Dim FilaDoc As String = Doc1 + vbTab + Doc2 + vbTab + Doc3 + vbTab + Doc4 + vbTab + Doc5 + vbTab + Doc6 + vbTab + Doc7 + _
                vbTab + Doc8 + vbTab + Doc9
            FilaDoc = Replace(FilaDoc, ".", "")
            FilaDoc = Replace(FilaDoc, ",", ".")

            ArrDoc(r) = FilaDoc
        Next

        'genera Archivo Documentos
        txt(Ruta & "Documentos.txt", tbl_gral.Rows.Count, ArrDoc)

        tbl_gral.Clear()
    End Sub

    Private Sub DespDir(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from saCliente ", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrDD(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim DD1 As String = tbl_gral.Rows(r)("co_cli").ToString.Trim
            DD1 = Delimitador(9, DD1)
            Dim DD2 As String = tbl_gral.Rows(r)("direc1").ToString.Trim
            DD2 = Delimitador(155, DD2)
            Dim DD3 As String = "1"
            DD3 = Delimitador(10, DD3)
            Dim DD4 As String = Delimitador(30, DD1)
            Dim DD5 As String = "1"

            Dim FilaDD As String
            FilaDD = DD1 + vbTab + DD2 + vbTab + DD3 + vbTab + DD4 + vbTab + DD5

            ArrDD(r) = FilaDD
        Next

        'genera Archivo DireccionesDespacho
        txt(Ruta & "DireccionesDespacho.txt", tbl_gral.Rows.Count, ArrDD)

        tbl_gral.Clear()
    End Sub

    Private Sub Supervisor(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from saVendedor where co_ven like 'SUP%'", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrSup(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Sup1 As String = tbl_gral.Rows(r)("co_ven").ToString.Trim
            Sup1 = Delimitador(20, Sup1)
            Dim Sup2 As String = tbl_gral.Rows(r)("ven_des").ToString.Trim
            Sup2 = Delimitador(50, Sup1)
            Dim Sup3Bool As Boolean = tbl_gral.Rows(r)("inactivo")
            Dim Sup3 As String
            Select Case Sup3Bool
                Case True
                    Sup3 = "1"
                Case False
                    Sup3 = "0"
            End Select
            Sup3 = Delimitador(1, Sup3)

            Dim FilaSup As String
            FilaSup = Sup1 + vbTab + Sup2 + vbTab + Sup3

            ArrSup(r) = FilaSup
        Next

        'genera Archivo DireccionesDespacho
        txt(Ruta & "Supervisor.txt", tbl_gral.Rows.Count, ArrSup)

        tbl_gral.Clear()
    End Sub

    Private Sub Zona(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("select * from saZona where campo1 = '1'", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrZon(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Zon1 As String = tbl_gral.Rows(r)("co_zon").ToString.Trim
            Zon1 = Delimitador(20, Zon1)
            Dim Zon2 As String = tbl_gral.Rows(r)("zon_des").ToString.Trim
            Zon2 = Delimitador(50, Zon2)
            Dim Zon3 As String = "1"
            Zon3 = Delimitador(1, Zon3)

            Dim Filazon As String
            Filazon = Zon1 + vbTab + Zon2 + vbTab + Zon3

            ArrZon(r) = Filazon
        Next


        'genera Archivo Zona
        txt(Ruta & "Zona.txt", tbl_gral.Rows.Count, ArrZon)

        tbl_gral.Clear()
    End Sub
#End Region

#Region "FuncionesCarga"

    Private Sub Modificar_Stock(ByVal art As String, cant As Double, ByVal conn As connect, ByVal tran As SqlTransaction)
        Dim CMD2 As New SqlCommand
        CMD2.Connection = conn.connName
        CMD2.CommandType = CommandType.StoredProcedure
        CMD2.Transaction = tran
        CMD2.CommandText = "pStockActualizar"
        CMD2.Parameters.Clear()


        'Parametros
        CMD2.Parameters.Add("@sCo_Alma", SqlDbType.Char)
        CMD2.Parameters.Add("@sCo_Art", SqlDbType.Char)
        CMD2.Parameters.Add("@sCo_Uni", SqlDbType.Char)
        CMD2.Parameters.Add("@deCantidad", SqlDbType.Decimal)
        CMD2.Parameters.Add("@sTipoStock", SqlDbType.Char)
        CMD2.Parameters.Add("@bSumarStock", SqlDbType.Bit)
        CMD2.Parameters.Add("@bPermiteStockNegativo", SqlDbType.Bit)

        'Valores
        CMD2.Parameters("@sCo_Alma").Value = "01" 'almacen que se esta modificando
        CMD2.Parameters("@sCo_Art").Value = art
        CMD2.Parameters("@sCo_Uni").Value = "FR" 'unidad primaria
        CMD2.Parameters("@deCantidad").Value = cant
        CMD2.Parameters("@sTipoStock").Value = "COM"
        CMD2.Parameters("@bSumarStock").Value = 1
        CMD2.Parameters("@bPermiteStockNegativo").Value = 0

        CMD2.ExecuteNonQuery()



    End Sub

    Function txtRead(ByVal ruta As String) As DataTable
        Dim sr As New StreamReader(ruta)

        Dim fullFileStr As String = sr.ReadToEnd()
        sr.Close()
        sr.Dispose()

        Dim lines As String() = fullFileStr.Split(ControlChars.Lf)
        Dim recs As New DataTable
        Dim sArr As String() = lines(0).Split(vbTab)
        For Each s As String In sArr
            recs.Columns.Add(New DataColumn())
        Next

        Dim row As DataRow
        Dim finalLine As String = ""
        For Each line As String In lines
            row = recs.NewRow()
            finalLine = line.Replace(Convert.ToString(ControlChars.Cr), "")
            row.ItemArray = finalLine.Split(vbTab)
            recs.Rows.Add(row)
        Next
        Return recs
    End Function

    Function FechaSDT(ByVal fecha As String) As DateTime
        Dim fech As DateTime
        fech = DateTime.ParseExact(fecha, "yyyyMMdd HHmmss", CultureInfo.InvariantCulture)
        Return fech
    End Function

    Function Reemplaza(ByVal expr As String) As String
        Dim Res As String

        Res = Replace(expr, ".", ",")
        Return Res
    End Function

    Function Reemplaza2(ByVal expr As String) As String
        Dim Res As String

        Res = Replace(expr, ".tmp", "")
        Return Res
    End Function

    Function Dec(ByVal valor As String) As Double
        Dim val As String
        val = Reemplaza(valor)
        Dim valDec As Double = CDbl(val)

        Return valDec
    End Function

    Function Stk(ByVal filtro As String, ByVal conn As connect, ByVal tr As SqlTransaction, ByVal comp As Boolean) As DataTable
        Dim tbl As New DataTable
        Dim CMD2 As New SqlCommand
        CMD2.Connection = conn.connName
        CMD2.CommandType = CommandType.StoredProcedure
        CMD2.Transaction = tr
        CMD2.CommandText = "StockCompvsStockAct"
        CMD2.Parameters.Clear()

        CMD2.Parameters.Add("@coArt", SqlDbType.VarChar)
        CMD2.Parameters.Add("@COMP", SqlDbType.Bit)
        CMD2.Parameters.Add("@almc", SqlDbType.VarChar)


        CMD2.Parameters("@coArt").Value = filtro
        CMD2.Parameters("@COMP").Value = comp
        CMD2.Parameters("@almc").Value = "01"

        adp = New SqlDataAdapter(CMD2)
        adp.Fill(tbl)
        Return tbl
    End Function

    Function FechaDatTime(ByVal fecha As String) As DateTime
        Dim fech As DateTime
        fech = DateTime.ParseExact(fecha, "yyyyMMdd HHmmssss", CultureInfo.InvariantCulture)
        Return fech
    End Function

#End Region

#Region "FuncionesExtracción"
    Public Sub txt(ByVal Ruta As String, ByVal Cta As Integer, ByVal Arr() As String)
        Using File As New System.IO.StreamWriter(Ruta)
            For i = 0 To Cta
                File.WriteLine(Arr(i))
            Next
            File.Close()
        End Using
    End Sub
    Function SelectDataTable(ByVal dt As DataTable, ByVal filter As String) As DataTable
        Dim row As DataRow()
        Dim dtNew As DataTable
        ' copy table structure
        dtNew = dt.Clone()
        ' sort and filter data
        row = dt.Select("co_cli" & "=" & "'" & filter & "'")
        ' fill dtNew with selected rows
        For Each dr As DataRow In row
            dtNew.ImportRow(dr)
        Next
        ' return filtered dt
        Return dtNew
    End Function
    Function Delimitador(ByVal lim As Integer, ByVal Expr As String) As String
        Dim Expresion As String
        If Expr.Length >= lim Then
            Expresion = Expr.Substring(0, lim)
        Else
            Expresion = Expr.PadRight(lim, " ")
        End If
        Expresion = Expresion.Replace(Chr(10), "")
        Expresion = Expresion.Replace(Chr(13), "")
        Expresion = Expresion.Replace(Chr(9), "")
        Return Expresion
    End Function


#End Region

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.

        tiempo.Enabled = False
        ClearMemory()
    End Sub


    Private Sub EscribirLog(ByVal Texto_Evento As String, ByVal tipo_entrada As EventLogEntryType)
        Dim Maquina As String = "."
        Dim Origen As String = "interface Merkant"
        'Escribimos en los Registros de Aplicación
        Dim Elog As EventLog
        Elog = New EventLog("Application", Maquina, Origen)
        Elog.WriteEntry(Texto_Evento, tipo_entrada, 100, CType(50, Short))
        Elog.Close()
        Elog.Dispose()
    End Sub

    Private Function totalRengPed(p1 As String) As Decimal
        Throw New NotImplementedException
    End Function












End Class
