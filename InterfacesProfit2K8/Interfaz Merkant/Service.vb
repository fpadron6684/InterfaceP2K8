Imports System.Data.SqlClient
Imports System.IO
Imports System.Globalization
Imports System.Timers.Timer
Imports System.Timers
Imports Interfaz_Merkant
Imports System.Threading

''' <summary>
''' Servicio que tomara su informacion de configuracion desde un archivo txt alojado en la carpeta raiz de la maquina que lo tenga instalado
''' cargara y extraera los datos para el uso de el sfera service, para el funcionamiento de Merkant.
''' Realizado por Sergio Mendoza Rivero
''' </summary>
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

    Private timeII As Integer

    Private th As Thread


    Private varConfGral(17) As String
    Private suc(4) As String
    Private ruta(4) As String
    Private bd(4) As String
    Private sSets(34) As Boolean
    Private sql As String

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
    Private tabla1 As New DataTable
    Private tabla2 As New DataTable
    Private tablaTrans As New DataTable
    Private tablaTrans1 As New DataTable
    Private tablaTrans2 As New DataTable
    Private tablaAct As New DataTable
    Private tablaComp As New DataTable
    Private adp As New SqlDataAdapter
    Private cmdBld As New SqlCommandBuilder

    Private tbl_gral As New DataTable

    Dim SaldoD As Double
    Dim SaldoC As Double

    Private contPed As Integer = 0
    Private contCob As Integer = 0
    Private contDep As Integer = 0
    Private cont As Integer=0
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
            timeII = CInt(varConfGral(16))

            'AddHandler tiempo.Elapsed, AddressOf tiempo_Tick
            AddHandler tiempoII.Elapsed, AddressOf tiempoII_Tick

            tiempo.Enabled = True
            tiempo.Interval = time


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

            'Selecciona Instancia SQL (solo la primera)
            sql = varConfGral(17)


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

                exportacionNivel1()
                'vigilante1.InternalBufferSize = 32768
                'vigilante1.Filter = "*.txt"
                vigilante1.IncludeSubdirectories = False
                'AddHandler vigilante1.Error, AddressOf OnError
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




    End Sub

#Region "FileSystemWatcher"
    'RUTA(0)
    'BD(0)
    'SUC(0)
    'SSETS (0 A 6)
    Private Sub sspedidos(fileDir As String, fileDir1 As String, BD As String, inst As String, ruta As String)
        Dim resultado = pedidos(fileDir, fileDir1, BD, inst)
        'pedidoDetalle(fileDir1, cmdSQL, conn, Trans)

        If resultado = True Then
            My.Computer.FileSystem.DeleteFile(fileDir)
            My.Computer.FileSystem.DeleteFile(fileDir1)
        Else
            Dim now As DateTime
            Dim nowStr As String
            now = DateTime.Parse(DateTime.Now)
            nowStr = now.ToString("yyyyMMdd-hhmmss")
            My.Computer.FileSystem.MoveFile(fileDir, ruta & "\Importacion\error\Pedido" & nowStr & ".txt")
            My.Computer.FileSystem.MoveFile(fileDir1, ruta + "\Importacion\error\PedidoDetalle" + nowStr + ".txt")
        End If
    End Sub
    Private Sub ssdevolucion(fileDir As String, fileDir1 As String, BD As String, inst As String)
        Dim conn As New connect(BD, inst) 'aclarar instancia SQL
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

    Private Sub sscobranza(fileDir As String, fileDir1 As String, fileDir2 As String, bd As String, inst As String, ruta As String)

        Dim resultado As Boolean = cobranza(fileDir, fileDir1, fileDir2, bd, inst)

        If resultado = True Then
            My.Computer.FileSystem.DeleteFile(fileDir)
            My.Computer.FileSystem.DeleteFile(fileDir1)
            My.Computer.FileSystem.DeleteFile(fileDir2)
        Else
            Dim now As DateTime
            Dim nowStr As String
            now = DateTime.Parse(DateTime.Now)
            nowStr = now.ToString("yyyyMMdd-hhmmss")
            My.Computer.FileSystem.MoveFile(fileDir, ruta + "\Importacion\error\Cobranza" + nowStr + ".txt")
            My.Computer.FileSystem.MoveFile(fileDir1, ruta + "\Importacion\error\CobranzaDetalle" + nowStr + ".txt")
            My.Computer.FileSystem.MoveFile(fileDir2, ruta + "\Importacion\error\CobranzaPago" + nowStr + ".txt")
        End If


    End Sub

    Private Sub ssautovta(fileDir As String, fileDir1 As String, bd As String, inst As String)
        Dim conn As New connect(bd, inst) 'aclarar instancia SQL
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
    Private Sub ssNCautovta(fileDir As String, bd As String, inst As String)
        Dim conn As New connect(bd, inst) 'aclarar instancia SQL
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
    Private Sub ssdepositos(fileDir As String, fileDir1 As String, bd As String, inst As String)
        Dim conn As New connect(bd, inst) 'aclarar instancia SQL
        conn.conectar()

        Dim Trans As SqlTransaction
        Trans = conn.connName.BeginTransaction

        Try
            cmdSQL.Connection = conn.connName
            cmdSQL.Transaction = Trans
            cmdSQL.CommandType = CommandType.StoredProcedure

            deposito(fileDir, fileDir1, cmdSQL)
            ' depositoDetalle(fileDir1, cmdSQL, conn)

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
                        contPed = contPed + 1
                    Else
                        fileDir(1) = ruta(0) + "\Importacion\" + ename
                        contPed = contPed + 1
                    End If

                    If contPed = 2 Then
                        vigilante1.EnableRaisingEvents = False
                        sspedidos(fileDir(0), fileDir(1), bd(0), sql, ruta(0))

                        contPed = 0
                        'Llama automaticamente la exportacion

                        exportacionNivel1()


                        fileDir(0) = ""
                        fileDir(1) = ""
                        'Thread.ResetAbort()

                        vigilante1.EnableRaisingEvents = True
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
                        ssdevolucion(fileDir(2), fileDir(3), bd(0), sql)
                        My.Computer.FileSystem.DeleteFile(fileDir(2))
                        My.Computer.FileSystem.DeleteFile(fileDir(3))

                        cont = 0
                        fileDir(2) = ""
                        fileDir(3) = ""
                    End If


                End If
            End If

            'SubSet III Cobranzas
            If sSets(2) = True Then

                If ename Like "Cobranza2*" Or ename Like "CobranzaDetalle*" Or ename Like "CobranzaPago*" Then

                    If ename Like "Cobranza2*" Then
                        fileDir(4) = Trim(ruta(0) + "\Importacion\" + ename)
                        contCob = contCob + 1
                    ElseIf ename Like "CobranzaDetalle*" Then
                        fileDir(5) = Trim(ruta(0) + "\Importacion\" + ename)
                        contCob = contCob + 1
                    ElseIf ename Like "CobranzaPago*" Then
                        fileDir(6) = Trim(ruta(0) + "\Importacion\" + ename)
                        contCob = contCob + 1
                    End If

                    If contCob = 3 Then
                        vigilante1.EnableRaisingEvents = False

                        sscobranza(fileDir(4), fileDir(5), fileDir(6), bd(0), sql, ruta(0))
                        contCob = 0
                        'Llama automaticamente la exportacion

                        'Try
                        'th = New Thread(AddressOf Me.exportacionNivel1)
                        'th.Start()
                        'exportacionNivel1()
                        'Catch ex As Exception
                        'EscribirLog("Error Generando Archivos " & ex.Message, EventLogEntryType.Error)
                        'Finally
                        'th.Abort()
                        'End Try



                        fileDir(4) = ""
                        fileDir(5) = ""
                        fileDir(6) = ""
                        'Thread.ResetAbort()
                        vigilante1.EnableRaisingEvents = True
                    End If
                End If

            End If

            'SubSet IV Facturas AutoVenta
            If sSets(3) = True Then
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
                        ssautovta(fileDir(7), fileDir(8), bd(0), sql)
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))

                        cont = 0
                        fileDir(7) = ""
                        fileDir(8) = ""
                        fileDir(9) = ""
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0), sql)
                        My.Computer.FileSystem.DeleteFile(fileDir(9))
                        NC = False
                    End If
                End If
            End If

            'SubSet V Depositos
            If sSets(4) = True Then
                If ename Like "Deposito2*" Or ename Like "DepositoDetalle*" Then
                    If ename Like "Deposito2*" Then
                        fileDir(10) = ruta(0) + "\Importacion\" + ename
                        contDep = contDep + 1
                    ElseIf ename Like "DepositoDetalle*" Then
                        fileDir(11) = ruta(0) + "\Importacion\" + ename
                        contDep = contDep + 1
                    End If

                    If contDep = 2 Then
                        ssdepositos(fileDir(10), fileDir(11), bd(0), sql)
                        My.Computer.FileSystem.DeleteFile(fileDir(10))
                        My.Computer.FileSystem.DeleteFile(fileDir(11))

                        contDep = 0
                        fileDir(10) = ""
                        fileDir(11) = ""
                    End If

                End If
            End If
        End If
        Thread.Sleep(2000)
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
                        sspedidos(fileDir(0), fileDir(1), bd(1), sql, ruta(1))
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
                        ssdevolucion(fileDir(2), fileDir(3), bd(0), sql)
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
                        cobranza(fileDir(4), fileDir(5), fileDir(6), bd(0), sql)
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
                        ssautovta(fileDir(7), fileDir(8), bd(0), sql)
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0), sql)
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
                        ssdepositos(fileDir(9), fileDir(10), bd(0), sql)
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
                        sspedidos(fileDir(0), fileDir(1), bd(2), sql, ruta(2))
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
                        ssdevolucion(fileDir(2), fileDir(3), bd(0), sql)
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
                        cobranza(fileDir(4), fileDir(5), fileDir(6), bd(0), sql)
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
                        ssautovta(fileDir(7), fileDir(8), bd(0), sql)
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0), sql)
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
                        ssdepositos(fileDir(9), fileDir(10), bd(0), sql)
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
                        sspedidos(fileDir(0), fileDir(1), bd(3), sql, ruta(3))
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
                        ssdevolucion(fileDir(2), fileDir(3), bd(0), sql)
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
                        cobranza(fileDir(4), fileDir(5), fileDir(6), bd(0), sql)
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
                        ssautovta(fileDir(7), fileDir(8), bd(0), sql)
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0), sql)
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
                        ssdepositos(fileDir(9), fileDir(10), bd(0), sql)
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
                        sspedidos(fileDir(0), fileDir(1), bd(4), sql, ruta(4))
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
                        ssdevolucion(fileDir(2), fileDir(3), bd(0), sql)
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
                        cobranza(fileDir(4), fileDir(5), fileDir(6), bd(0), sql)
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
                        ssautovta(fileDir(7), fileDir(8), bd(0), sql)
                        My.Computer.FileSystem.DeleteFile(fileDir(7))
                        My.Computer.FileSystem.DeleteFile(fileDir(8))
                        cont = 0
                        For i = 0 To fileDir.GetUpperBound(0)
                            fileDir(i) = ""
                        Next
                    End If
                    If NC = True Then
                        ssNCautovta(fileDir(9), bd(0), sql)
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
                        ssdepositos(fileDir(9), fileDir(10), bd(0), sql)
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

    ''' <summary>
    ''' 
    ''' </summary>
    Private Sub exportacionNivel1()
        EscribirLog("Exportación Archivos Prioridad I iniciada a las: " & DateTime.Now, EventLogEntryType.Information)
        tiempo.Enabled = False

        If suc(0) = 1 Then

            Dim conn As New connect(bd(0), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(0) + "\Exportacion\", ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                ClienteRuta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTE RUTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Lp(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DIRECCION DESPACHO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Cls1(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls2(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls3(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls4(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 4 ", EventLogEntryType.Error)
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

            'tiempo.Enabled = True

        End If

        If suc(1) = 1 Then
            Dim conn As New connect(bd(1), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(0) + "\Exportacion\", ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                ClienteRuta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTE RUTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                Lp(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DIRECCION DESPACHO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Cls1(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls2(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls3(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls4(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 4 ", EventLogEntryType.Error)
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
            'tiempo.Enabled = True
        End If

        If suc(2) = 1 Then
            Dim conn As New connect(bd(2), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(0) + "\Exportacion\", ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                ClienteRuta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTE RUTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                Lp(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DIRECCION DESPACHO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Cls1(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls2(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls3(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls4(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 4 ", EventLogEntryType.Error)
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
            'tiempo.Enabled = True
        End If

        If suc(3) = 1 Then
            Dim conn As New connect(bd(3), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(0) + "\Exportacion\", ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                ClienteRuta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTE RUTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                Lp(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DIRECCION DESPACHO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Cls1(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls2(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls3(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls4(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 4 ", EventLogEntryType.Error)
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
            'tiempo.Enabled = True
        End If

        If suc(4) = 1 Then
            Dim conn As New connect(bd(4), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                Vendedores(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(0) + "\Exportacion\", ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                ClienteRuta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTE RUTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                Lp(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DIRECCION DESPACHO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Cls1(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls2(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls3(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls4(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 4 ", EventLogEntryType.Error)
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
            'tiempo.Enabled = True
        End If
        EscribirLog("Exportación Archivos Prioridad I Culminada a las: " & DateTime.Now, EventLogEntryType.Information)
        tiempo.Enabled = True

        tiempoII.Enabled = True
        tiempoII.Interval = timeII

    End Sub

    Private Sub tiempoII_Tick(sender As Object, e As ElapsedEventArgs)
        tiempoII.Enabled = False
        EscribirLog("Exportación Archivos Prioridad II iniciada a las: " & DateTime.Now, EventLogEntryType.Information)
        If suc(0) = 1 Then

            Dim conn As New connect(bd(0), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                MotNoVis(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotNoVta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotDev(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DE DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                incidencias(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO inCIDENCIAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            ClearMemory()

            Try
                conn.conectar()
                Vendedores(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO VENDEDORES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                TipoNegocio(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO TIPONEGOCIO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Clientes(conn, ruta(0) + "\Exportacion\", ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTES ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                ClienteRuta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLIENTE RUTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try
            Try
                conn.conectar()
                Lp(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO LISTA DE PRECIOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                UndMed(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO UNIDADMEDIDA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Sku(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO PRODUCTO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Almcn(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ALMACEN ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Docs(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DOCUMENTOS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                DespDir(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO DIRECCION DESPACHO ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Supervisor(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO SUPERVISOR ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Zona(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO ZONAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
                ClearMemory()
            End Try

            Try
                conn.conectar()
                Cls1(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 1 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls2(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 2 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls3(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 3 ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                Cls4(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO CLASIFICACION 4 ", EventLogEntryType.Error)
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
            Dim conn As New connect(bd(1), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                MotNoVis(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotNoVta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotDev(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DE DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                incidencias(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO inCIDENCIAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

        End If

        If suc(2) = 1 Then
            Dim conn As New connect(bd(2), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                MotNoVis(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotNoVta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotDev(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DE DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                incidencias(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO inCIDENCIAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

        End If

        If suc(3) = 1 Then
            Dim conn As New connect(bd(3), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                MotNoVis(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotNoVta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotDev(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DE DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                incidencias(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO inCIDENCIAS ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try
        End If

        If suc(4) = 1 Then
            Dim conn As New connect(bd(4), sql) 'aclarar instancia SQL

            Try
                conn.conectar()
                MotNoVis(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VISITA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotNoVta(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO NO VENTA ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                MotDev(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO MOTIVO DE DEVOLUCION ", EventLogEntryType.Error)
            Finally
                conn.cerrar()
            End Try

            Try
                conn.conectar()
                incidencias(conn, ruta(0) + "\Exportacion\")
            Catch ex As Exception
                EscribirLog("ERROR " & ex.Message & " GENERANDO ARCHIVO inCIDENCIAS ", EventLogEntryType.Error)
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

    Function pedidos(fileDir As String, fileDir2 As String, bd As String, inst As String) As Boolean
        tabla = txtRead(fileDir)
        tabla2 = txtRead(fileDir2)
        Dim rdo As Boolean
        Dim fact As Integer
        Dim contResPos As Integer
        Dim contResNeg As Integer

        Dim conn As New connect(bd, inst) 'aclarar instancia SQL


        Dim Tran As SqlTransaction



        For r = 0 To tabla.Rows.Count - 2
            conn.conectar()
            cmdSQL.Connection = conn.connName
            Tran = conn.connName.BeginTransaction
            cmdSQL.Transaction = Tran
            cmdSQL.CommandType = CommandType.StoredProcedure

            Try

                cmdSQL.CommandText = "pp_ins_pedidos_MERK"
                cmdSQL.Parameters.Clear()

                'ENCABEZADO
                cmdSQL.Parameters.Add("@CODPEDIDO", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@FECHA", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@FECHADESPACHO", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@CANTIDADTOTAL", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@MONTO", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@IMPUESTO", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@DESCUENTO", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@CODVENDEDOR", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@CODCLIENTE", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@PORCENTAJEDESCUENTO", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@COMENTARIOPED", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@ORDENCOMPRA", SqlDbType.VarChar)
                'PRIMER RENGLON
                cmdSQL.Parameters.Add("@CODPRODUCTO", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@CANTIDAD", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@PRECIO", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@DESCUENTOART", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@PORCENTAJEDESCUENTOART", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@IMPUESTOART", SqlDbType.VarChar)
                cmdSQL.Parameters.Add("@SERIAL", SqlDbType.VarChar)

                'VALORES PARA ENCABEZADO

                Dim numpedido As String = tabla.Rows(r)("column1")
                cmdSQL.Parameters("@CODPEDIDO").Value = tabla.Rows(r)("column1")
                cmdSQL.Parameters("@FECHA").Value = tabla.Rows(r)("column2")
                cmdSQL.Parameters("@FECHADESPACHO").Value = tabla.Rows(r)("column3")
                cmdSQL.Parameters("@CANTIDADTOTAL").Value = tabla.Rows(r)("column4")
                cmdSQL.Parameters("@MONTO").Value = tabla.Rows(r)("column5")
                cmdSQL.Parameters("@IMPUESTO").Value = tabla.Rows(r)("column6")
                cmdSQL.Parameters("@DESCUENTO").Value = tabla.Rows(r)("column7")
                cmdSQL.Parameters("@CODVENDEDOR").Value = tabla.Rows(r)("column8")
                cmdSQL.Parameters("@CODCLIENTE").Value = tabla.Rows(r)("column9")
                cmdSQL.Parameters("@PORCENTAJEDESCUENTO").Value = tabla.Rows(r)("column10")
                cmdSQL.Parameters("@COMENTARIOPED").Value = tabla.Rows(r)("column11")
                cmdSQL.Parameters("@ORDENCOMPRA").Value = tabla.Rows(r)("column12")

                Dim tbl As DataTable

                tbl = SelectDataTable2(tabla2, tabla.Rows(r)("column1"))

                'Cargar Renglon inicial del Pedido, Validar antes que haya existencia del mismo en la disponibilidad
                Dim hayExistencia As Boolean = False
                Dim s As Integer = 0
                Dim p As Integer = 0
                Dim reng As Integer = tbl.Rows.Count


                Do While hayExistencia = False And reng > s

                    Dim Total_Stock As Double = 0

                    Dim tablaStock As DataTable = Stk(tbl.Rows(s)("column2").ToString.Trim, conn, Tran, False)


                    If tablaStock.Rows.Count > 0 Then
                        Total_Stock = Dec(tablaStock.Rows(0)("stock"))
                    Else
                        Total_Stock = 0
                    End If
                    tablaStock.Clear()


                    Dim Total_art As String = Dec(tbl.Rows(s)("column3"))

                    Dim disp As Double = Total_Stock

                    If disp > 0 And Total_art <= disp Then

                        cmdSQL.Parameters("@CODPRODUCTO").Value = tbl.Rows(s)("column2")
                        cmdSQL.Parameters("@CANTIDAD").Value = tbl.Rows(s)("column3")
                        cmdSQL.Parameters("@PRECIO").Value = tbl.Rows(s)("column4")
                        cmdSQL.Parameters("@DESCUENTOART").Value = tbl.Rows(s)("column5")
                        cmdSQL.Parameters("@PORCENTAJEDESCUENTOART").Value = tbl.Rows(s)("column6")
                        cmdSQL.Parameters("@IMPUESTOART").Value = tbl.Rows(s)("column7")
                        cmdSQL.Parameters("@SERIAL").Value = tbl.Rows(s)("column8")

                        hayExistencia = True
                        fact = cmdSQL.ExecuteScalar()
                        fact = s + 1
                    Else
                        hayExistencia = False

                        EscribirLog("El articulo: " & tbl.Rows(s)("column2").ToString & " del pedido de venta: " & tbl.Rows(s)("column1") & " no pudo ser cargado por no haber suficiente Stock Disponible", EventLogEntryType.Warning)
                        s = s + 1
                    End If

                Loop
                'realizamos el alta

                If fact <> 0 Then
                    Dim res As Boolean = pedidoDetalle(tbl, conn, tran, s + 1, fact, tbl.Rows(s)("column2"))
                    If res = False Then
                        EscribirLog("El pedido " & tbl.Rows(s)("column1") & " / " & fact & " fue cargado exitosamente", EventLogEntryType.Information)
                    Else
                        EscribirLog("El pedido " & tbl.Rows(s)("column1") & " / " & fact & " fue cargado con detalles", EventLogEntryType.Information)
                    End If
                Else
                    EscribirLog("Para el pedido de venta: " & tabla.Rows(r)("column1") & " no hay disponibilidad en ninguno de sus renglones, el pedido no será cargado ", EventLogEntryType.Warning)
                End If

                Tran.Commit()
                contResPos = contResPos + 1

                'Actualizar el cobro
                cmdSQL.Connection = conn.connName
                cmdSQL.CommandType = CommandType.StoredProcedure

                cmdSQL.CommandText = "Merk_AjustesPedidos"
                cmdSQL.Parameters.Clear()

                'Inicio parametros para carga de la tabla COBROS

                cmdSQL.Parameters.Add("@Num_Pedido", SqlDbType.Char)


                cmdSQL.Parameters("@Num_Pedido").Value = numpedido

                cmdSQL.ExecuteNonQuery()

                EscribirLog("El pedido de venta # " & tabla.Rows(r)("column1") & " fue cargado con Exito", EventLogEntryType.Information)

            Catch ex As Exception
                EscribirLog("El pedido de venta: " & tabla.Rows(r)("column1") & " no pudo ser cargado por el siguiente error: " & ex.Message, EventLogEntryType.Error)
                Tran.Rollback()

                contResNeg = contResNeg + 1
            Finally
                conn.cerrar()
            End Try

        Next
        tabla.Clear()

        If contResNeg > 0 Then
            rdo = False
        Else
            rdo = True
        End If

        Return rdo

    End Function

    Function pedidoDetalle(tabla As DataTable, conn As connect, tran As SqlTransaction, ind As Integer, fact As Integer, art As String) As Boolean
        'tabla = txtRead(fileDir)
        Dim res As Boolean = False

        Dim cmdSQL1 As New SqlCommand
        cmdSQL1.Connection = conn.connName
        cmdSQL1.CommandType = CommandType.StoredProcedure
        cmdSQL1.Transaction = tran
        cmdSQL1.CommandText = "pp_ins_reng_ped_Merk"
        cmdSQL1.Parameters.Clear()

        cmdSQL1.Parameters.Add("@CODPEDIDO", SqlDbType.VarChar)
        cmdSQL1.Parameters.Add("@CODPRODUCTO", SqlDbType.VarChar)
        cmdSQL1.Parameters.Add("@CANTIDAD", SqlDbType.VarChar)
        cmdSQL1.Parameters.Add("@PRECIO", SqlDbType.VarChar)
        cmdSQL1.Parameters.Add("@DESCUENTOART", SqlDbType.VarChar)
        cmdSQL1.Parameters.Add("@PORCENTAJEDESCUENTOART", SqlDbType.VarChar)
        cmdSQL1.Parameters.Add("@IMPUESTOART", SqlDbType.VarChar)
        cmdSQL1.Parameters.Add("@SERIAL", SqlDbType.VarChar)



        Dim ctrReng As Boolean = False
        Dim nroDoc As String
        Dim ArtPrgRow As Integer = 1

        For r = ind To tabla.Rows.Count - 1
            Try

                Dim Total_Stock As Double = 0

                Dim tablaStock As DataTable = Stk(tabla.Rows(r)("column2").ToString.Trim, conn, tran, False)


                If tablaStock.Rows.Count > 0 Then
                    Total_Stock = Dec(tablaStock.Rows(0)("stock"))
                Else
                    Total_Stock = 0
                End If
                tablaStock.Clear()


                Dim Total_art As String = Dec(tabla.Rows(r)("column3"))

                Dim disp As Double = Total_Stock


                If disp > 0 And Total_art <= disp Then

                    cmdSQL1.Parameters("@CODPEDIDO").Value = tabla.Rows(r)("column1")
                    cmdSQL1.Parameters("@CODPRODUCTO").Value = tabla.Rows(r)("column2")
                    cmdSQL1.Parameters("@CANTIDAD").Value = tabla.Rows(r)("column3")
                    cmdSQL1.Parameters("@PRECIO").Value = tabla.Rows(r)("column4")
                    cmdSQL1.Parameters("@DESCUENTOART").Value = tabla.Rows(r)("column5")
                    cmdSQL1.Parameters("@PORCENTAJEDESCUENTOART").Value = tabla.Rows(r)("column6")
                    cmdSQL1.Parameters("@IMPUESTOART").Value = tabla.Rows(r)("column7")
                    cmdSQL1.Parameters("@SERIAL").Value = tabla.Rows(r)("column8")

                    cmdSQL1.ExecuteNonQuery()


                Else

                    EscribirLog("El articulo: " & tabla.Rows(r)("column2").ToString & " del pedido de venta: " & nroDoc & "/" & fact & " no pudo ser cargado por no haber suficiente Stock Disponible", EventLogEntryType.Warning)
                    res = True
                End If

            Catch ex As Exception
                EscribirLog("El articulo: " & tabla.Rows(r)("column2").ToString & " del pedido de venta: " & nroDoc & "/" & fact & " no pudo ser cargado por presentarse el siguiente error: " & ex.Message, EventLogEntryType.Warning)
                res = True
            End Try
        Next
        Return res
    End Function

    Private Sub devolucion(fileDir As String, cmdSQL As SqlCommand)
        Throw New NotImplementedException
    End Sub

    Private Sub devoluciondetalle(fileDir As String, cmdSQL As SqlCommand, conn As connect)
        Throw New NotImplementedException
    End Sub

    ''' <summary>
    ''' Carga el cobro en la BD, toma como parametros las tres rutas de los 3 archivos involucrados en el proceso de cobranza
    ''' Cobranza, CobranzaDetalleDocumento y CobranzaDetallePago
    ''' SM
    ''' </summary>
    ''' <param name="fileDir"> ruta archivo cobranza (encabezado)</param>
    ''' <param name="fileDir1"> ruta archivo detalle documentos de la cobranza</param>
    ''' <param name="fileDir2"> ruta archivo detalle pago cobranza</param>
    ''' <param name="cmdSQL"> CommandSQL </param>
    ''' <param name="conn"> Conexion SQL</param>
    ''' <param name="tran"> Transaccion SQL</param>
    Function cobranza(fileDir As String, fileDir1 As String, fileDir2 As String, bd As String, inst As String) As Boolean
        Dim res As Boolean
        Dim contResPos As Integer
        Dim contResNeg As Integer
        Dim numcobro As String

        'Limpia tablas que seran utilizadas
        tbl_gral.Clear() 'tabla encabezado
        tabla1.Clear() 'tabla detalle cobro (documentos)
        tabla2.Clear() 'tabla detalle pago (formas de pago) de esta hay que extraer las retenciones

        tbl_gral = txtRead(fileDir)
        tabla1 = txtRead(fileDir1)
        tabla2 = txtRead(fileDir2)

        Dim conn As New connect(bd, inst) 'aclarar instancia SQL


        Dim Trans As SqlTransaction


#Region "lee e inserta linea por linea registros del encabezado de la cobranza almacenados temporalmente en tbl_gral"

        For r = 0 To tbl_gral.Rows.Count - 2
            conn.conectar()
            cmdSQL.Connection = conn.connName
            Trans = conn.connName.BeginTransaction
            cmdSQL.Transaction = Trans
            cmdSQL.CommandType = CommandType.StoredProcedure

            Try

                cmdSQL.CommandText = "pp_ins_cobros_MERK"
                cmdSQL.Parameters.Clear()

                'Inicio parametros para carga de la tabla COBROS

                cmdSQL.Parameters.Add("@Num_Cobro", SqlDbType.Char)
                cmdSQL.Parameters.Add("@Codcliente", SqlDbType.Char)
                cmdSQL.Parameters.Add("@FechaCobro", SqlDbType.Char)
                cmdSQL.Parameters.Add("@MontoCobro", SqlDbType.Char)

                'Fin parametros tabla COBROS
                'Inicio parametros para la carga del primer renglon de RENG_COB

                cmdSQL.Parameters.Add("@TipDoc", SqlDbType.Char)
                cmdSQL.Parameters.Add("@DocNum", SqlDbType.Char)
                cmdSQL.Parameters.Add("@MontoCob", SqlDbType.Char)

                'Fin parametros tabla RENG_COB

                'Inicio Parametros tabla COBROS
                numcobro = tbl_gral.Rows(r)("column1")
                cmdSQL.Parameters("@Num_Cobro").Value = tbl_gral.Rows(r)("column1")
                cmdSQL.Parameters("@Codcliente").Value = tbl_gral.Rows(r)("column4")
                cmdSQL.Parameters("@FechaCobro").Value = tbl_gral.Rows(r)("column3")
                cmdSQL.Parameters("@MontoCobro").Value = tbl_gral.Rows(r)("column2")

                'Fin parametros tabla COBROS


                'Filtrado de Datatable de documentos al cobro, por numero del encabezado de cobranza que se encuentra siendo procesado al momento
                'y posterior carga del primer renglon de la misma.

                Dim tbl As DataTable

                tbl = SelectDataTable2(tabla1, tbl_gral.Rows(r)("column1"))

                'determinar cuantos registros trae la tabla filtrada
                Dim contReng As Integer = tbl.Rows.Count  'cantidad de documentos asociados al cobro

                'conversion del primer renglon a tipo de documento aceptado y recibido por el profit
                Dim tpDoc As String
                tpDoc = tbl.Rows(0)("column10")

#Region "carga del primer renglon solicitado por el sp del encabezado de cobro"
                'el renglon debe ser el 1 siempre en este proceso

                cmdSQL.Parameters("@TipDoc").Value = tbl.Rows(0)("column10")
                cmdSQL.Parameters("@DocNum").Value = tbl.Rows(0)("column2")
                cmdSQL.Parameters("@MontoCob").Value = tbl.Rows(0)("column4")

                'Recuperar numero profit de la cobranza generada para utilizarlo al cargar detalle documento y detalle pago
                Dim cob As Integer = cmdSQL.ExecuteScalar()

#End Region

#End Region

#Region "insercion de detalleCobro"
                ' en caso de existir mas documentos que cancelar o retenciones creadas que relacionar al cobro se pasa a cargar documentos detalle
                If contReng > 1 Then
                    cobranzadetalle(tabla1, cmdSQL, conn, Trans, cob, contReng)
                End If
#End Region

#Region "Insercion Detalle Tipo de Pago"

                'Variable que guardara el total de las formas de pago, monto que sera recuperado de la ejecucion de la funcion cobranzapago
                Dim totFP As Double

                totFP = cobranzapago(tabla2, cmdSQL, conn, Trans, cob, tbl_gral.Rows(r)("column5").ToString)

                'Actualizar el cobro
                cmdSQL.Connection = conn.connName
                cmdSQL.CommandType = CommandType.StoredProcedure

                cmdSQL.CommandText = "Merk_AjustesCobros"
                cmdSQL.Parameters.Clear()

                'Inicio parametros para carga de la tabla COBROS

                cmdSQL.Parameters.Add("@Num_Cobro", SqlDbType.Char)


                cmdSQL.Parameters("@Num_Cobro").Value = numcobro

                cmdSQL.ExecuteNonQuery()



#End Region
                Trans.Commit()
                contResPos = contResPos + 1

                EscribirLog("La Cobranza Merkant # " & tbl_gral.Rows(r)("column1") & " fue cargada con Exito", EventLogEntryType.Information)

            Catch ex As Exception
                Trans.Rollback()
                EscribirLog("La Cobranza Merkant # " & tbl_gral.Rows(r)("column1") & " no pudo ser procesada por el siguiente error: " & ex.Message, EventLogEntryType.Error)
                contResNeg = contResNeg + 1
            Finally
                conn.cerrar()

            End Try
        Next
        tabla.Clear()

        If contResNeg > 0 Then
            res = False
        Else
            res = True
        End If

        Return res
    End Function

    Private Sub GeneraAdelanto(tblAd As DataTable, conn As connect, tran As SqlTransaction, cob As Integer)
        'declaro objetos de base datos para insertar documento tipo AJNM mediante SP pp_ins_docum_cc_RET

        Dim CMDret As New SqlCommand
        CMDret.Connection = conn.connName
        CMDret.CommandType = CommandType.StoredProcedure
        CMDret.Transaction = tran
        CMDret.CommandText = "pp_ins_docum_cc_AJPM"
        CMDret.Parameters.Clear()

#Region "leo e inserto linea por linea las retenciones existentes"

        For r = 0 To tblAd.Rows.Count - 1

            Dim tpDoc As String
            tpDoc = tblAd.Rows(r)("Column8")

            CMDret.Parameters.AddWithValue("@Nro_Cobro", tblAd.Rows(r)("Column1"))
            CMDret.Parameters.AddWithValue("@FechaDoc", tblAd.Rows(r)("Column7"))
            CMDret.Parameters.AddWithValue("@NroOrigen", tblAd.Rows(r)("Column9"))
            CMDret.Parameters.AddWithValue("@TipoOrigen", tblAd.Rows(r)("Column8"))
            CMDret.Parameters.AddWithValue("@Monto", tblAd.Rows(r)("Column4"))

            CMDret.ExecuteNonQuery()
            'limpio parametros para reutilizarlos en caso de haber mas retenciones pendientes por cargar
            CMDret.Parameters.Clear()

        Next

#End Region
    End Sub

    ''' <summary>
    ''' Crea el Documento ANJM correspondiente a la retencion enviada por el movil
    ''' dicho documento luego sera llamado por el detalle de documento de la cobranza
    ''' </summary>
    ''' <param name="tblRet">Tabla contentiva de los documentos tipo Retencion</param>
    ''' <param name="conn"> Conexion SQL</param>
    ''' <param name="tran"> Transaccion SQL</param>
    ''' <param name="cob"> Numero cobranza implicita</param>
    Private Sub GeneraRetencion(tblRet As DataTable, conn As connect, tran As SqlTransaction, cob As Integer)

        'declaro objetos de base datos para insertar documento tipo AJNM mediante SP pp_ins_docum_cc_RET

        Dim CMDret As New SqlCommand
        CMDret.Connection = conn.connName
        CMDret.CommandType = CommandType.StoredProcedure
        CMDret.Transaction = tran
        CMDret.CommandText = "pp_ins_docum_cc_RET"
        CMDret.Parameters.Clear()

#Region "leo e inserto linea por linea las retenciones existentes"

        For r = 0 To tblRet.Rows.Count - 1

            Dim tpDoc As String
            tpDoc = tblRet.Rows(r)("Column8")

            CMDret.Parameters.AddWithValue("@Nro_Cobro", tblRet.Rows(r)("Column1"))
            CMDret.Parameters.AddWithValue("@FechaDoc", tblRet.Rows(r)("Column7"))
            CMDret.Parameters.AddWithValue("@NroOrigen", tblRet.Rows(r)("Column9"))
            CMDret.Parameters.AddWithValue("@TipoOrigen", tblRet.Rows(r)("Column8"))
            CMDret.Parameters.AddWithValue("@Monto", tblRet.Rows(r)("Column4"))

            CMDret.ExecuteNonQuery()
            'limpio parametros para reutilizarlos en caso de haber mas retenciones pendientes por cargar
            CMDret.Parameters.Clear()

        Next

#End Region

    End Sub

    ''' <summary>
    ''' Inserta renglon por renglon los documentos de cobro asociados al encabezado de cobro procesado en SUB cobranza
    ''' </summary>
    ''' <param name="tabla"> Datatable de los documentos a cargar</param>
    ''' <param name="cmdSQL"> Command SQL</param>
    ''' <param name="conn"> Conexion SQL</param>
    ''' <param name="tran"> Transaccion SQL</param>
#Region "Leo e inserto registros subsiguientes de documentos en el cobro"
    Private Sub cobranzadetalle(tabla As DataTable, cmdSQL As SqlCommand, conn As connect, tran As SqlTransaction, cob As Integer, contreg As Integer)

        cmdSQL.CommandText = "pp_ins_reng_cob_MERK"
        cmdSQL.Parameters.Clear()

        'valido si hay mas de un renglon en el detalle de documentos al cobro, en caso positivo procede a leerlos e insertarlos uno por uno
        'desde la posicion nro 2 (recordemos que ya tuvo que haber sido cargado el primero)
        If contreg > 1 Then

            For r = 1 To tabla.Rows.Count - 2
                'limpio parametros en caso de que hayan mas registros que ingresar y deban ser reutilizados
                cmdSQL.Parameters.Clear()

                cmdSQL.Parameters.AddWithValue("@Num_Cobro", (tabla.Rows(r)("column1")))
                cmdSQL.Parameters.AddWithValue("@TipDoc", tabla.Rows(r)("column10"))
                cmdSQL.Parameters.AddWithValue("@DocNum", (tabla.Rows(r)("column2")))
                cmdSQL.Parameters.AddWithValue("@MontoCob", (tabla.Rows(r)("column4")))

                cmdSQL.ExecuteNonQuery()
            Next
        End If

        tabla.Clear()
    End Sub
#End Region

    ''' <summary>
    ''' Cargar detalle del pago de la cobranza que esta siendo procesada
    ''' </summary>
    ''' <param name="tabla"> tabla filtrada </param>
    ''' <param name="cmdSQL"> commandSql </param>
    ''' <param name="conn"> Conexion SQL</param>
    ''' <param name="tran"> Transaccion SQL</param>
    ''' <param name="cob"> Numero de cobro procesado</param>
    Function cobranzapago(tablaTP As DataTable, cmdSQL As SqlCommand, conn As connect, tran As SqlTransaction, cob As Integer, codVen As String)

        cmdSQL.CommandText = "pp_ins_reng_tip_MERK"
        cmdSQL.Parameters.Clear()

        cmdSQL.Parameters.Add("@Num_Cobro", SqlDbType.Char)
        cmdSQL.Parameters.Add("@NumDoc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@TipDoc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@MontoDoc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@BancoO", SqlDbType.Char)
        cmdSQL.Parameters.Add("@FechaDoc", SqlDbType.Char)
        cmdSQL.Parameters.Add("@TipoOrig", SqlDbType.Char)
        cmdSQL.Parameters.Add("@NroOrigen", SqlDbType.Char)


        For r = 0 To tablaTP.Rows.Count - 2

            cmdSQL.Parameters("@Num_Cobro").Value = tablaTP.Rows(r)("column1")
            cmdSQL.Parameters("@NumDoc").Value = tablaTP.Rows(r)("column2")
            cmdSQL.Parameters("@TipDoc").Value = tablaTP.Rows(r)("column3")
            cmdSQL.Parameters("@MontoDoc").Value = tablaTP.Rows(r)("column4")
            cmdSQL.Parameters("@BancoO").Value = tablaTP.Rows(r)("column6")
            cmdSQL.Parameters("@FechaDoc").Value = tablaTP.Rows(r)("column7")
            cmdSQL.Parameters("@TipoOrig").Value = tablaTP.Rows(r)("column8")
            cmdSQL.Parameters("@NroOrigen").Value = tablaTP.Rows(r)("column9")

            cmdSQL.ExecuteNonQuery()
        Next
        tablaTP.Clear()
        Return 0
    End Function
#End Region



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

    Private Sub deposito(fileDir As String, fileDir1 As String, cmdSQL As SqlCommand)
        tabla = txtRead(fileDir)
        tabla1 = txtRead(fileDir1) 'Tabla de Deposito Detalle

        For r = 0 To tabla.Rows.Count - 2
            cmdSQL.CommandText = "dbo.insDeposito"
            cmdSQL.Parameters.Clear()

            cmdSQL.Parameters.Add("@dep_num", SqlDbType.Int)
            cmdSQL.Parameters.Add("@deposito", SqlDbType.Char)
            cmdSQL.Parameters.Add("@fecha", SqlDbType.SmallDateTime)
            cmdSQL.Parameters.Add("@movi", SqlDbType.Int)
            cmdSQL.Parameters.Add("@bancoDep", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@cod_caja", SqlDbType.Char)
            cmdSQL.Parameters.Add("@movie", SqlDbType.Int)
            cmdSQL.Parameters.Add("@total_efec", SqlDbType.Decimal)
            cmdSQL.Parameters.Add("@total_cheq", SqlDbType.Decimal)
            cmdSQL.Parameters.Add("@total_tarj", SqlDbType.Decimal)
            cmdSQL.Parameters.Add("@che_dev", SqlDbType.Int)
            cmdSQL.Parameters.Add("@cta_egre", SqlDbType.Char)
            cmdSQL.Parameters.Add("@feccom", SqlDbType.SmallDateTime)
            cmdSQL.Parameters.Add("@numcom", SqlDbType.Int)
            cmdSQL.Parameters.Add("@cta_cont01", SqlDbType.Char)
            cmdSQL.Parameters.Add("@cta_cont02", SqlDbType.Char)
            cmdSQL.Parameters.Add("@cta_cont03", SqlDbType.Char)
            cmdSQL.Parameters.Add("@dis_cen", SqlDbType.Text)
            cmdSQL.Parameters.Add("@moneda", SqlDbType.Char)
            cmdSQL.Parameters.Add("@tasa", SqlDbType.Decimal)
            cmdSQL.Parameters.Add("@campo1", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@campo2", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@campo3", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@campo4", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@campo5", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@campo6", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@campo7", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@campo8", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@co_us_in", SqlDbType.Char)
            cmdSQL.Parameters.Add("@fe_us_in", SqlDbType.DateTime)
            cmdSQL.Parameters.Add("@co_us_mo", SqlDbType.Char)
            cmdSQL.Parameters.Add("@fe_us_mo", SqlDbType.DateTime)
            cmdSQL.Parameters.Add("@co_us_el", SqlDbType.Char)
            cmdSQL.Parameters.Add("@fe_us_el", SqlDbType.DateTime)
            cmdSQL.Parameters.Add("@revisado", SqlDbType.Char)
            cmdSQL.Parameters.Add("@trasnfe", SqlDbType.Char)
            cmdSQL.Parameters.Add("@co_sucu", SqlDbType.Char)
            cmdSQL.Parameters.Add("@AUX01", SqlDbType.Decimal)
            cmdSQL.Parameters.Add("@AUX02", SqlDbType.VarChar)
            cmdSQL.Parameters.Add("@next_number", SqlDbType.Int)
            cmdSQL.Parameters("@dep_num").Value = 0
            cmdSQL.Parameters("@deposito").Value = tabla.Rows(r)("Column3").ToString
            cmdSQL.Parameters("@fecha").Value = FechaSDT(tabla.Rows(r)("Column5"))
            cmdSQL.Parameters("@movi").Value = 0
            cmdSQL.Parameters("@bancoDep").Value = tabla.Rows(r)("Column4").ToString
            cmdSQL.Parameters("@cod_caja").Value = ""
            cmdSQL.Parameters("@movie").Value = 0
            cmdSQL.Parameters("@total_efec").Value = 0
            cmdSQL.Parameters("@total_cheq").Value = Dec(tabla.Rows(r)("Column6"))
            cmdSQL.Parameters("@total_tarj").Value = 0
            cmdSQL.Parameters("@che_dev").Value = 0
            cmdSQL.Parameters("@cta_egre").Value = "134101"
            cmdSQL.Parameters("@feccom").Value = FechaSDT(tabla.Rows(r)("Column5"))
            cmdSQL.Parameters("@numcom").Value = 0
            cmdSQL.Parameters("@cta_cont01").Value = ""
            cmdSQL.Parameters("@cta_cont02").Value = ""
            cmdSQL.Parameters("@cta_cont03").Value = ""
            cmdSQL.Parameters("@dis_cen").Value = ""
            cmdSQL.Parameters("@moneda").Value = "BSF"
            cmdSQL.Parameters("@tasa").Value = 1
            cmdSQL.Parameters("@campo1").Value = tabla.Rows(r)("Column1").ToString
            cmdSQL.Parameters("@campo2").Value = tabla.Rows(r)("Column2").ToString
            Dim codCja As String = tabla.Rows(r)("Column2").ToString
            cmdSQL.Parameters("@campo3").Value = ""
            cmdSQL.Parameters("@campo4").Value = ""
            cmdSQL.Parameters("@campo5").Value = ""
            cmdSQL.Parameters("@campo6").Value = ""
            cmdSQL.Parameters("@campo7").Value = ""
            cmdSQL.Parameters("@campo8").Value = ""
            cmdSQL.Parameters("@co_us_in").Value = "999"
            cmdSQL.Parameters("@fe_us_in").Value = FechaSDT(tabla.Rows(r)("Column5"))
            cmdSQL.Parameters("@co_us_mo").Value = ""
            cmdSQL.Parameters("@fe_us_mo").Value = FechaSDT(tabla.Rows(r)("Column5"))
            cmdSQL.Parameters("@co_us_el").Value = ""
            cmdSQL.Parameters("@fe_us_el").Value = FechaSDT(tabla.Rows(r)("Column5"))
            cmdSQL.Parameters("@revisado").Value = ""
            cmdSQL.Parameters("@trasnfe").Value = ""
            cmdSQL.Parameters("@co_sucu").Value = "01"
            cmdSQL.Parameters("@AUX01").Value = 0
            cmdSQL.Parameters("@AUX02").Value = ""
            cmdSQL.Parameters("@next_number").Value = 0

            Dim depNum As Integer
            depNum = cmdSQL.ExecuteScalar()

            Dim tablaDepDet As DataTable = SelectDataTable2(tabla1, tabla.Rows(r)("Column1").ToString)
            depositoDetalle(tablaDepDet, cmdSQL, depNum, codCja)
        Next

        tabla.Clear()

    End Sub

    Private Sub depositoDetalle(tabla As DataTable, cmdSQL As SqlCommand, depNum As Integer, codCja As String)

        cmdSQL.Parameters.Clear()
        cmdSQL.CommandText = "dbo.InsDepositoDetalle"


        cmdSQL.Parameters.Add("@dep_num", SqlDbType.Int)
        cmdSQL.Parameters.Add("@reng_num", SqlDbType.Int)
        cmdSQL.Parameters.Add("@codigo", SqlDbType.Char)
        cmdSQL.Parameters.Add("@mov_afec", SqlDbType.Int)
        cmdSQL.Parameters.Add("@mov_gene", SqlDbType.Int)
        cmdSQL.Parameters.Add("@forma_pag", SqlDbType.Char)
        cmdSQL.Parameters.Add("@fecha", SqlDbType.SmallDateTime)
        cmdSQL.Parameters.Add("@doc_num", SqlDbType.Char)
        cmdSQL.Parameters.Add("@descrip", SqlDbType.Char)
        cmdSQL.Parameters.Add("@monto", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@origen", SqlDbType.Char)
        cmdSQL.Parameters.Add("@cob_pag", SqlDbType.Int)
        cmdSQL.Parameters.Add("@banc_tarj", SqlDbType.Char)
        cmdSQL.Parameters.Add("@comision", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@impuesto", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@tipo_plazo", SqlDbType.Char)
        cmdSQL.Parameters.Add("@AUX01", SqlDbType.Decimal)
        cmdSQL.Parameters.Add("@AUX02", SqlDbType.Char)

        Dim reng As Integer = 1
        For r = 0 To tabla.Rows.Count - 1

            cmdSQL.Parameters("@dep_num").Value = depNum

            cmdSQL.Parameters("@reng_num").Value = reng
            reng = reng + 1
            cmdSQL.Parameters("@codigo").Value = codCja
            cmdSQL.Parameters("@mov_afec").Value = 0
            cmdSQL.Parameters("@mov_gene").Value = 0
            cmdSQL.Parameters("@forma_pag").Value = "CH"
            cmdSQL.Parameters("@fecha").Value = DateTime.Now
            cmdSQL.Parameters("@doc_num").Value = tabla.Rows(r)("Column4").ToString
            cmdSQL.Parameters("@descrip").Value = ""
            cmdSQL.Parameters("@monto").Value = 0
            cmdSQL.Parameters("@origen").Value = ""
            cmdSQL.Parameters("@cob_pag").Value = 0
            cmdSQL.Parameters("@banc_tarj").Value = ""
            cmdSQL.Parameters("@comision").Value = 0
            cmdSQL.Parameters("@impuesto").Value = 0
            cmdSQL.Parameters("@tipo_plazo").Value = "1"
            cmdSQL.Parameters("@AUX01").Value = 0
            cmdSQL.Parameters("@AUX02").Value = ""



            cmdSQL.ExecuteNonQuery()
        Next

        tabla.Clear()

    End Sub



#Region "MétodosExtracción"
    Public Sub Vendedores(ByVal Conex As connect, Ruta As String)
        '////////////////////////////////////////////////
        '//////////////////Vendedor.txt//////////////////
        '////////////////////////////////////////////////

        'adp = New SqlDataAdapter("select co_ven,ven_des,campo1,campo2,condic from Vendedor where co_ven Not Like '%sup%'", Conex.connName)
        adp = New SqlDataAdapter("Merk_Vendedores", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        tbl_gral.Clear()

        adp.Fill(tbl_gral)


        Dim ArrVen(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Ven1 As String = tbl_gral.Rows(r)("codvendedor").ToString.Trim
            Ven1 = Delimitador(6, Ven1)
            Dim Ven2 As String = tbl_gral.Rows(r)("nombre").ToString.Trim
            Ven2 = Delimitador(60, Ven2)
            Dim Ven3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim Ven3 As String
            If Ven3_Bool = True Then
                Ven3 = "0"
            Else
                Ven3 = "1"
            End If
            Ven3 = Delimitador(1, Ven3)
            Dim Ven4 As String = tbl_gral.Rows(r)("codzonaerp").ToString.Trim
            Ven4 = Delimitador(20, Ven4)
            Dim Ven5 As String = tbl_gral.Rows(r)("codrutaerp").ToString.Trim
            Ven5 = Delimitador(6, Ven5)
            Dim Ven6 As String = tbl_gral.Rows(r)("codsupervisor").ToString.Trim
            Ven6 = Delimitador(20, Ven6)
            Dim Ven7 As String = tbl_gral.Rows(r)("tiporuta").ToString.Trim
            Ven7 = Delimitador(1, Ven7)
            Dim Ven8 As String = tbl_gral.Rows(r)("refbancaria").ToString.Trim
            Ven8 = Delimitador(20, Ven8)

            Dim FilaVen As String
            FilaVen = Ven1 + vbTab + Ven2 + vbTab + Ven3 + vbTab + Ven4 + vbTab + Ven5 + vbTab + Ven6 + vbTab + Ven7 + vbTab + Ven8

            ArrVen(r) = FilaVen
        Next

        'genera Archivo Vendedores

        txt(Ruta & "Vendedor.txt", tbl_gral.Rows.Count, ArrVen)
        tbl_gral.Clear()

    End Sub

    Public Sub ClienteRuta(ByVal Conex As connect, Ruta As String)
        '////////////////////////////////////////////////
        '//////////////////Vendedor.txt//////////////////
        '////////////////////////////////////////////////

        'adp = New SqlDataAdapter("select co_ven,ven_des,campo1,campo2,condic from Vendedor where co_ven Not Like '%sup%'", Conex.connName)
        adp = New SqlDataAdapter("Merk_ClienteRuta", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        tbl_gral.Clear()

        adp.Fill(tbl_gral)


        Dim ArrVen(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Ven1 As String = tbl_gral.Rows(r)("codcliente").ToString.Trim
            Ven1 = Delimitador(15, Ven1)
            Dim Ven2 As String = tbl_gral.Rows(r)("codvendedor").ToString.Trim
            Ven2 = Delimitador(60, Ven2)

            Dim FilaVen As String
            FilaVen = Ven1 + vbTab + Ven2

            ArrVen(r) = FilaVen
        Next

        'genera Archivo Vendedores

        txt(Ruta & "clienteruta.txt", tbl_gral.Rows.Count, ArrVen)
        tbl_gral.Clear()

    End Sub

    Public Sub TipoNegocio(ByVal Conex As connect, Ruta As String)
        '////////////////////////////////////////////////
        '//////////////////TipoNegocio.txt////////////
        '////////////////////////////////////////////////

        'adp = New SqlDataAdapter("select co_seg,seg_des from segmento", Conex.connName)
        adp = New SqlDataAdapter("Merk_TipoNegocio", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrTp(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Tp1 As String = ((tbl_gral.Rows(r)("codtiponegocio")).ToString).Trim
            Tp1 = Delimitador(6, Tp1)
            Dim Tp2 As String = ((tbl_gral.Rows(r)("descripcion").ToString).Trim)
            Tp2 = Delimitador(60, Tp2)
            Dim Tp3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim Tp3 As String
            If Tp3_Bool = True Then
                Tp3 = "0"
            Else
                Tp3 = "1"
            End If
            Tp3 = Delimitador(1, Tp3)

            Dim FilaTp As String = Tp1 + vbTab + Tp2 + vbTab + Tp3
            ArrTp(r) = FilaTp
        Next


        'Genera archivo TipoNegocio
        txt(Ruta & "TipoNegocio.txt", tbl_gral.Rows.Count, ArrTp)
        tbl_gral.Clear()
    End Sub

    Public Sub Clientes(ByVal Conex As connect, Ruta As String, Ruta2 As String)
        adp = New SqlDataAdapter("Merk_Clientes", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrTp(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Cl1 As String = ((tbl_gral.Rows(r)("codcliente")).ToString).Trim
            Cl1 = Delimitador(15, Cl1)
            Dim Cl2 As String = ((tbl_gral.Rows(r)("nombre").ToString).Trim)
            Cl2 = Delimitador(60, Cl2)
            Dim Cl3 As String = ((tbl_gral.Rows(r)("telefono1").ToString).Trim)
            Cl3 = Delimitador(30, Cl3)
            Dim Cl4 As String = ((tbl_gral.Rows(r)("telefono2").ToString).Trim)
            Cl4 = Delimitador(30, Cl4)
            Dim Cl5 As String = ((tbl_gral.Rows(r)("direccion1").ToString).Trim)
            Cl5 = Delimitador(255, Cl5)
            Dim Cl6 As String = ((tbl_gral.Rows(r)("NumeroTributario1").ToString).Trim)
            Cl6 = Delimitador(20, Cl6)
            Dim Cl7_DT As DateTime = tbl_gral.Rows(r)("fechastatus")
            Dim Cl7 As String
            Cl7_DT = DateTime.Parse(Cl7_DT)
            Cl7 = Cl7_DT.ToString("yyyyMMdd")
            Dim Cl8_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim Cl8 As String
            If Cl8_Bool = True Then
                Cl8 = "0"
            Else
                Cl8 = "1"
            End If
            Cl8 = Delimitador(1, Cl8)
            Dim Cl9 As String = Delimitador(19, Format(tbl_gral.Rows(r)("limitecredito"), "###0.00"))
            Cl9 = Replace(Cl9, ",", ".")
            Dim Cl10 As String = ((tbl_gral.Rows(r)("codtipoaccion").ToString).Trim)
            Cl10 = Delimitador(10, Cl10)
            Dim Cl11 As String = Delimitador(19, Format(tbl_gral.Rows(r)("saldo"), "###0.00"))
            Cl11 = Replace(Cl11, ",", ".")
            Dim Cl12 As String = ((tbl_gral.Rows(r)("codvendedor").ToString).Trim)
            Cl12 = Delimitador(6, Cl12)
            Dim Cl13 As String
            Cl13 = tbl_gral.Rows(r)("diascredito")
            Cl13 = Delimitador(3, Cl13)
            Dim Cl14 As String = ((tbl_gral.Rows(r)("nombrepropietario").ToString).Trim)
            Cl14 = Delimitador(50, Cl14)
            Dim Cl15 As String = ((tbl_gral.Rows(r)("estado").ToString).Trim)
            Cl15 = Delimitador(30, Cl15)
            Dim Cl16 As String = ((tbl_gral.Rows(r)("ciudad").ToString).Trim)
            Cl16 = Delimitador(30, Cl16)
            Dim Cl17 As String = ((tbl_gral.Rows(r)("codtiponegocio").ToString).Trim)
            Cl17 = Delimitador(100, Cl17)
            Dim Cl18 As String = ((tbl_gral.Rows(r)("formapago").ToString).Trim)
            Cl18 = Delimitador(1, Cl18)
            Dim Cl19 As String = ((tbl_gral.Rows(r)("razonsocial").ToString).Trim)
            Cl19 = Delimitador(60, Cl19)
            Dim Cl20 As String = ((tbl_gral.Rows(r)("correo").ToString).Trim)
            Cl20 = Delimitador(50, Cl20)
            Dim Cl21 As String = ((tbl_gral.Rows(r)("codlistaprecioerp").ToString).Trim)
            Cl21 = Delimitador(10, Cl21)
            Dim Cl22 As String = ((tbl_gral.Rows(r)("codigoescaneo").ToString).Trim)
            Cl22 = Delimitador(20, Cl22)

            Dim FilaTp As String = Cl1 + vbTab + Cl2 + vbTab + Cl3 + vbTab + Cl4 + vbTab + Cl5 + vbTab + Cl6 + vbTab + Cl7 + vbTab + Cl8 + vbTab + Cl9 + vbTab + Cl10 + vbTab + Cl11 + vbTab + Cl12 + vbTab + Cl13 + vbTab + Cl14 + vbTab + Cl15 + vbTab + Cl16 + vbTab + Cl17 + vbTab + Cl18 + vbTab + Cl19 + vbTab + Cl20 + vbTab + Cl21 + vbTab + Cl22
            ArrTp(r) = FilaTp
            FilaTp = Replace(FilaTp, ".", "")
            FilaTp = Replace(FilaTp, ",", ".")
        Next

        'Genera archivo Cliente
        txt(Ruta & "Cliente.txt", tbl_gral.Rows.Count, ArrTp)
        'Genera archivo PlanCliente
        'txt(Ruta2 & "ClientePLanificacion.txt", tbl_gral.Rows.Count, ArrCp)

        tbl_gral.Clear()
        tablaTrans.Clear()
    End Sub

    Private Sub MotNoVis(ByVal Conex As connect, Ruta As String)
        '////////////////////////////////////////////////
        '//////////////////MotivoNoVisita.txt////////////
        '////////////////////////////////////////////////
        adp = New SqlDataAdapter("Merk_MotivoNoVisita", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrMnV(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Mnv1 As String = tbl_gral.Rows(r)("codmotivonovisita").ToString
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
        adp = New SqlDataAdapter("Merk_MotivoNoVenta", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrMnVta(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Mnv1 As String = tbl_gral.Rows(r)("CODMOTIVONOVENTA").ToString
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
        adp = New SqlDataAdapter("Merk_MotivoDevolucion", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrDev(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim mdev1 As String = tbl_gral.Rows(r)("CODMOTIVOdevolucion").ToString
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
        adp = New SqlDataAdapter("Merk_VisitaMotivo", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim Arrinc(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim inc1 As String = tbl_gral.Rows(r)("CodVisitaMotivo").ToString
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
        adp = New SqlDataAdapter("Merk_Clasificacion1", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)

        tbl_gral.Clear()
        adp.Fill(tbl_gral)

        Dim ArrCl1(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Cl2_1 As String = tbl_gral.Rows(r)("codclasificacion1").ToString
            Cl2_1 = Delimitador(18, Cl2_1)
            Dim Cl2_2 As String = tbl_gral.Rows(r)("descripcion").ToString
            Cl2_2 = Delimitador(50, Cl2_2)
            Dim Cl2_3 As String = "1"
            Cl2_3 = Delimitador(1, Cl2_3)




            Dim FilaCl2 As String = Cl2_1 + vbTab + Cl2_2 + vbTab + Cl2_3
            ArrCl1(r) = FilaCl2
        Next

        'Genera Archivo Clasificacion1
        txt(Ruta & "Clasificacion1.txt", tbl_gral.Rows.Count, ArrCl1)

        tbl_gral.Clear()
    End Sub

    Private Sub Cls2(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("Merk_Clasificacion2 ", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrCl3(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Cl3_1 As String = tbl_gral.Rows(r)("codclasificacion2").ToString
            Cl3_1 = Delimitador(18, Cl3_1)
            Dim Cl3_2 As String = tbl_gral.Rows(r)("codclasificacion1").ToString
            Cl3_2 = Delimitador(50, Cl3_2)
            Dim Cl3_3 As String = "1"
            Cl3_3 = Delimitador(1, Cl3_3)
            Dim Cl3_4 As String = tbl_gral.Rows(r)("descripcion").ToString
            Cl3_4 = Delimitador(18, Cl3_4)



            Dim FilaCl3 As String = Cl3_1 + vbTab + Cl3_4 + vbTab + Cl3_2 + vbTab + Cl3_3
            ArrCl3(r) = FilaCl3
        Next

        'Genera Archivo Clasificacion 3
        txt(Ruta & "Clasificacion2.txt", tbl_gral.Rows.Count, ArrCl3)
        tbl_gral.Clear()
    End Sub

    Private Sub Cls3(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("Merk_Clasificacion3 ", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrCl3(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Cl3_1 As String = tbl_gral.Rows(r)("codclasificacion3").ToString
            Cl3_1 = Delimitador(18, Cl3_1)
            Dim Cl3_2 As String = tbl_gral.Rows(r)("codclasificacion2").ToString
            Cl3_2 = Delimitador(50, Cl3_2)
            Dim Cl3_3 As String = "1"
            Cl3_3 = Delimitador(1, Cl3_3)
            Dim Cl3_4 As String = tbl_gral.Rows(r)("descripcion").ToString
            Cl3_4 = Delimitador(18, Cl3_4)



            Dim FilaCl3 As String = Cl3_1 + vbTab + Cl3_4 + vbTab + Cl3_2 + vbTab + Cl3_3
            ArrCl3(r) = FilaCl3
        Next

        'Genera Archivo Clasificacion 3
        txt(Ruta & "Clasificacion3.txt", tbl_gral.Rows.Count, ArrCl3)
        tbl_gral.Clear()
    End Sub

    Private Sub Cls4(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("Merk_Clasificacion4 ", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrCl3(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Cl3_1 As String = tbl_gral.Rows(r)("codclasificacion4").ToString
            Cl3_1 = Delimitador(18, Cl3_1)
            Dim Cl3_2 As String = tbl_gral.Rows(r)("codclasificacion3").ToString
            Cl3_2 = Delimitador(50, Cl3_2)
            Dim Cl3_3 As String = "1"
            Cl3_3 = Delimitador(1, Cl3_3)
            Dim Cl3_4 As String = tbl_gral.Rows(r)("descripcion").ToString
            Cl3_4 = Delimitador(18, Cl3_4)



            Dim FilaCl3 As String = Cl3_1 + vbTab + Cl3_4 + vbTab + Cl3_2 + vbTab + Cl3_3
            ArrCl3(r) = FilaCl3
        Next

        'Genera Archivo Clasificacion 3
        txt(Ruta & "Clasificacion4.txt", tbl_gral.Rows.Count, ArrCl3)
        tbl_gral.Clear()
    End Sub

    Private Sub Lp(ByVal Conex As connect, Ruta As String)

        adp = New SqlDataAdapter("Merk_ListaP", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)


        Dim ArrLp(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Lp1 As String = tbl_gral.Rows(r)("codlistaprecio").ToString.Trim
            Lp1 = Delimitador(10, Lp1)
            Dim Lp2 As String = tbl_gral.Rows(r)("codproducto").ToString.Trim
            Lp2 = Delimitador(30, Lp2)
            Dim Lp3 As String = Delimitador(21, Format(tbl_gral.Rows(r)("preciocompra"), "###0.00"))
            Dim Lp4 As String = Delimitador(21, Format(tbl_gral.Rows(r)("precioventa"), "###0.00"))
            Dim Lp5 As String = tbl_gral.Rows(r)("codunidadmedida")
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
        adp = New SqlDataAdapter("Merk_Unidades", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrUm(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Um1 As String = (tbl_gral.Rows(r)("codunidadmedida").ToString).Trim
            Um1 = Delimitador(20, Um1)
            Dim Um2 As String = (tbl_gral.Rows(r)("descripcion").ToString).Trim
            Um2 = Delimitador(50, Um2)
            Dim Um3_Bool As Boolean = tbl_gral.Rows(r)("status").ToString
            Dim Um3 As String
            If Um3_Bool = True Then
                Um3 = "1"
            Else
                Um3 = "0"
            End If
            Um3 = Delimitador(1, Um3)
            Dim FilaUm As String = Um1 + vbTab + Um2 + vbTab + Um3
            ArrUm(r) = FilaUm

        Next

        'genera Archivo UnidadMedida
        txt(Ruta & "UnidadMedida.txt", tbl_gral.Rows.Count, ArrUm)

        tbl_gral.Clear()
    End Sub

    Private Sub Sku(ByVal Conex As connect, Ruta As String)
        Dim Consulta As String = "Merk_Producto"
        adp = New SqlDataAdapter(Consulta, Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrPr(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Pr1 As String = tbl_gral.Rows(r)("codproducto").ToString.Trim
            Pr1 = Delimitador(30, Pr1)
            Dim Pr2 As String = tbl_gral.Rows(r)("descripcioncorta").ToString.Trim
            Pr2 = Delimitador(35, Pr2)
            Dim Pr3_bool As Boolean = tbl_gral.Rows(r)("status")
            Dim Pr3 As String
            If Pr3_bool = False Then
                Pr3 = "1"
            Else
                Pr3 = "0"
            End If
            Pr3 = Delimitador(1, Pr3)
            Dim Pr4_int As Integer = tbl_gral.Rows(r)("aplicaiva")
            Dim Pr4 As String = Pr4_int.ToString
            Pr4 = Delimitador(1, Pr4)
            Dim Pr5 As String = tbl_gral.Rows(r)("tipoproducto").ToString.Trim
            Pr5 = Delimitador(1, Pr5)
            Dim Pr6_int As Integer = tbl_gral.Rows(r)("unidades").ToString.Trim
            Dim Pr6 As String
            Pr6 = Delimitador(4, Pr6_int)
            Dim Pr7 As String = Delimitador(21, Format(tbl_gral.Rows(r)("volumen"), "###0.00"))
            Dim Pr8 As String = tbl_gral.Rows(r)("categoria1").ToString.Trim
            Pr8 = Delimitador(50, Pr8)
            Dim Pr9 As String = tbl_gral.Rows(r)("categoria2").ToString.Trim
            Pr9 = Delimitador(50, Pr9)
            Dim Pr10 As String = tbl_gral.Rows(r)("categoria3").ToString.Trim
            Pr10 = Delimitador(50, Pr10)
            Dim Pr11 As String = tbl_gral.Rows(r)("categoria4").ToString.Trim
            Pr11 = Delimitador(50, Pr11)
            Dim Pr12 As String = tbl_gral.Rows(r)("codunidadmedida").ToString.Trim
            Pr12 = Delimitador(20, Pr12)
            Dim Pr13_Bool As Boolean = tbl_gral.Rows(r)("esvacio").ToString
            Dim Pr13 As String
            If Pr13_Bool = True Then
                Pr13 = "1"
            Else
                Pr13 = "0"
            End If
            Pr13 = Delimitador(1, Pr13)
            Dim Pr14 As String = tbl_gral.Rows(r)("codproductovacio").ToString.Trim
            Pr14 = Delimitador(20, Pr14)
            Dim Pr15 As String = Delimitador(21, Format(tbl_gral.Rows(r)("cantidadvacio"), "###0.00"))
            Dim Pr16 As String = Delimitador(21, Format(tbl_gral.Rows(r)("porcentajeieps"), "###0.00"))


            Dim FilaPr As String
            FilaPr = Pr1 + vbTab + Pr2 + vbTab + Pr3 + vbTab + Pr4 + vbTab + Pr5 + vbTab + Pr6 + vbTab + Pr7 + vbTab + Pr8 + vbTab + Pr9 + vbTab + Pr10 +
                 vbTab + Pr11 + vbTab + Pr12 + vbTab + Pr13 + vbTab + Pr14 + vbTab + Pr15 + vbTab + Pr16
            FilaPr = Replace(FilaPr, ",", ".")
            ArrPr(r) = FilaPr
        Next

        'genera Archivo Producto
        txt(Ruta & "Producto.txt", tbl_gral.Rows.Count, ArrPr)

        tbl_gral.Clear()
    End Sub

    Private Sub Bco(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("Merk_Banco", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrBco(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Bco1 As String = tbl_gral.Rows(r)("codbanco").ToString.Trim
            Bco1 = Delimitador(6, Bco1)
            Dim Bco2 As String = tbl_gral.Rows(r)("nombre").ToString.Trim
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
        adp = New SqlDataAdapter("Merk_Historia", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrHis(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim His1 As String = tbl_gral.Rows(r)("coddocumento").ToString.Trim
            His1 = Delimitador(30, His1)
            Dim His2 As String = tbl_gral.Rows(r)("tipodocumento")
            His2 = Delimitador(2, His2)
            Dim His3 As String = tbl_gral.Rows(r)("codcliente").ToString.Trim
            His3 = Delimitador(15, His3)
            Dim His4_DT As DateTime = tbl_gral.Rows(r)("fechaemision")
            Dim His4 As String
            His4_DT = DateTime.Parse(His4_DT)
            His4 = His4_DT.ToString("yyyyMMdd")
            Dim His5_Doub As Double = tbl_gral.Rows(r)("total")
            Dim His5 As String
            His5 = Delimitador(21, Format(His5_Doub, "##,##0.00"))
            Dim His6_Doub As Double = tbl_gral.Rows(r)("descuento")
            Dim His6 As String
            His6 = Delimitador(21, Format(His6_Doub, "##,##0.00"))
            Dim His7_Doub As Double = tbl_gral.Rows(r)("impuesto")
            Dim His7 As String
            His7 = Delimitador(21, Format(His7_Doub, "##,##0.00"))
            Dim His8_Bool As Boolean = tbl_gral.Rows(r)("codstatushistoria").ToString
            Dim His8 As String
            If His8_Bool = True Then
                His8 = "1"
            Else
                His8 = "0"
            End If
            His8 = Delimitador(30, His8)


            Dim His9_DT As DateTime = tbl_gral.Rows(r)("fechastatus")
            Dim His9 As String
            His9_DT = DateTime.Parse(His9_DT)
            His9 = His9_DT.ToString("yyyyMMdd")

            Dim FilaHis As String
            FilaHis = His1 + vbTab + His2 + vbTab + His3 + vbTab + His4 + vbTab + His5 + vbTab + His6 + vbTab + His7 + vbTab + His8 + vbTab + His9
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
        adp = New SqlDataAdapter("Merk_HistoriaDetalle ", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrDH(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Dh1 As String = tbl_gral.Rows(r)("coddocumento").ToString.Trim
            Dh1 = Delimitador(30, Dh1)

            Dim Dh2 As String = tbl_gral.Rows(r)("tipodocumento")
            Dh2 = Delimitador(2, Dh2)

            Dim Dh3 As String = tbl_gral.Rows(r)("codproducto").ToString.Trim
            Dh3 = Delimitador(30, Dh3)

            Dim Dh4 As String = tbl_gral.Rows(r)("codunidadmedida").ToString.Trim
            Dh4 = Delimitador(20, Dh4)

            Dim Dh5_Doub As Double = tbl_gral.Rows(r)("cantidad")
            Dim Dh5 As String = Delimitador(21, Format(Dh5_Doub, "##,##0.00"))

            Dim Dh6_Doub As Double = tbl_gral.Rows(r)("precio")
            Dim Dh6 As String = Delimitador(21, Format(Dh6_Doub, "##,##0.00"))

            Dim Dh7_Doub As Double = tbl_gral.Rows(r)("total")
            Dim Dh7 As String = Delimitador(21, Format(Dh7_Doub, "##,##0.00"))

            Dim Dh8_Doub As Double = tbl_gral.Rows(r)("descuento")
            Dim Dh8 As String = Delimitador(21, Format(Dh8_Doub, "##,##0.00"))

            Dim Dh9_Doub As Double = tbl_gral.Rows(r)("impuesto")
            Dim Dh9 As String = Delimitador(21, Format(Dh9_Doub, "##,##0.00"))

            Dim Dh10 As String = tbl_gral.Rows(r)("coddocumentoasociado").ToString.Trim
            Dh10 = Delimitador(30, Dh10)

            Dim Dh11 As String = tbl_gral.Rows(r)("tipodocumentoasociado").ToString.Trim
            Dh11 = Delimitador(30, Dh11)


            Dh10 = Delimitador(30, Dh10)
            Dh11 = Delimitador(2, Dh11)

            Dim FilaDh As String = Dh1 + vbTab + Dh2 + vbTab + Dh3 + vbTab + Dh4 + vbTab + Dh5 + vbTab + Dh6 + vbTab +
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
        adp = New SqlDataAdapter("Merk_Almacen", Conex.connName)
        adp.SelectCommand.CommandType = CommandType.StoredProcedure
        adp.Fill(tbl_gral)

        Dim ArrAlm(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Alm1 As String = tbl_gral.Rows(r)("codalmacen").ToString.Trim
            Alm1 = Delimitador(20, Alm1)
            Dim Alm2 As String = tbl_gral.Rows(r)("codproducto").ToString.Trim
            Alm2 = Delimitador(30, Alm2)
            Dim Alm3 As String = tbl_gral.Rows(r)("codunidadmedida").ToString.Trim
            Alm3 = Delimitador(20, Alm3)
            Dim Alm4_Doub As Double = tbl_gral.Rows(r)("cantidad")
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
        adp = New SqlDataAdapter("Merk_Documentos", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)
        Dim ArrDoc(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Doc1 As String = tbl_gral.Rows(r)("coddocumentoerp").ToString.Trim
            Doc1 = Delimitador(14, Doc1)
            Dim Doc2 As String = tbl_gral.Rows(r)("codclienteerp").ToString.Trim
            Doc2 = Delimitador(15, Doc2)
            Dim Doc3 As String = tbl_gral.Rows(r)("tipodocumento").ToString.Trim
            Doc3 = Delimitador(3, Doc3)
            Dim Doc4 As String
            Doc4 = tbl_gral.Rows(r)("cancelado")
            Doc4 = Delimitador(1, Doc4)
            Dim Doc5_DT As DateTime = tbl_gral.Rows(r)("fechaemision")
            Dim Doc5 As String
            Doc5_DT = DateTime.Parse(Doc5_DT)
            Doc5 = Doc5_DT.ToString("yyyyMMdd")
            Dim Doc6_DT As DateTime = tbl_gral.Rows(r)("fechavencimiento")
            Dim Doc6 As String
            Doc6_DT = DateTime.Parse(Doc6_DT)
            Doc6 = Doc6_DT.ToString("yyyyMMdd")
            Dim Doc7 As String = Delimitador(21, Format(tbl_gral.Rows(r)("saldo"), "###0.00"))
            Dim Doc8 As String = Delimitador(21, Format(tbl_gral.Rows(r)("total"), "###0.00"))
            Dim Doc9_Bool As Boolean = tbl_gral.Rows(r)("anulado").ToString
            Dim Doc9 As String
            If Doc9_Bool = True Then
                Doc9 = "0"
            Else
                Doc9 = "1"
            End If

            Dim FilaDoc As String = Doc1 + vbTab + Doc2 + vbTab + Doc3 + vbTab + Doc4 + vbTab + Doc5 + vbTab + Doc6 + vbTab + Doc7 +
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
        adp = New SqlDataAdapter("MerkCliDespacho", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrDD(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim DD1 As String = tbl_gral.Rows(r)("codclienteerp").ToString.Trim
            DD1 = Delimitador(15, DD1)
            Dim DD2 As String = tbl_gral.Rows(r)("direccion").ToString.Trim
            DD2 = Delimitador(155, DD2)
            Dim DD3 As String = tbl_gral.Rows(r)("prioridad").ToString.Trim
            DD3 = Delimitador(10, DD3)
            Dim DD4 As String = tbl_gral.Rows(r)("coddireccionerp").ToString.Trim
            DD4 = Delimitador(30, DD1)
            Dim DD5 As String
            DD5 = tbl_gral.Rows(r)("direccionfiscal")
            DD5 = Delimitador(1, DD5)

            Dim FilaDD As String
            FilaDD = DD1 + vbTab + DD2 + vbTab + DD3 + vbTab + DD4 + vbTab + DD5

            ArrDD(r) = FilaDD
        Next

        'genera Archivo DireccionesDespacho
        txt(Ruta & "DireccionesDespacho.txt", tbl_gral.Rows.Count, ArrDD)

        tbl_gral.Clear()
    End Sub

    Private Sub Supervisor(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("Merk_Supervisor", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrSup(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Sup1 As String = tbl_gral.Rows(r)("codsupervisorerp").ToString.Trim
            Sup1 = Delimitador(20, Sup1)
            Dim Sup2 As String = tbl_gral.Rows(r)("nombre").ToString.Trim
            Sup2 = Delimitador(50, Sup2)
            Dim Sup3Bool As Boolean = tbl_gral.Rows(r)("status")
            Dim Sup3 As String
            Select Case Sup3Bool
                Case True
                    Sup3 = "0"
                Case False
                    Sup3 = "1"
            End Select
            Sup3 = Delimitador(1, Sup3)
            Dim Sup4 As String = tbl_gral.Rows(r)("password").ToString.Trim
            Sup4 = Delimitador(30, Sup4)


            Dim FilaSup As String
            FilaSup = Sup1 + vbTab + Sup2 + vbTab + Sup3

            ArrSup(r) = FilaSup
        Next

        'genera Archivo DireccionesDespacho
        txt(Ruta & "Supervisor.txt", tbl_gral.Rows.Count, ArrSup)

        tbl_gral.Clear()
    End Sub

    Private Sub Zona(ByVal Conex As connect, Ruta As String)
        adp = New SqlDataAdapter("Merk_Zona", Conex.connName)
        cmdBld = New SqlCommandBuilder(adp)
        adp.Fill(tbl_gral)

        Dim ArrZon(tbl_gral.Rows.Count) As String

        For r = 0 To tbl_gral.Rows.Count - 1
            Dim Zon1 As String = tbl_gral.Rows(r)("codzonaerp").ToString.Trim
            Zon1 = Delimitador(20, Zon1)
            Dim Zon2 As String = tbl_gral.Rows(r)("descripcion").ToString.Trim
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

    Function SelectDataTable2(ByVal dt As DataTable, ByVal filter As String) As DataTable
        Dim row As DataRow()
        Dim dtNew As DataTable
        ' copy table structure
        dtNew = dt.Clone()
        ' sort and filter data
        row = dt.Select("Column1" & "=" & "'" & filter & "'")
        ' fill dtNew with selected rows
        For Each dr As DataRow In row
            dtNew.ImportRow(dr)
        Next
        ' return filtered dt
        Return dtNew
    End Function

    Function SelectDataTable3(ByVal dt As DataTable, ByVal filter As String) As DataTable
        Dim row As DataRow()
        Dim dtNew As DataTable
        ' copy table structure
        dtNew = dt.Clone()
        ' sort and filter data
        row = dt.Select("Column3" & "=" & "'" & filter & "'")
        ' fill dtNew with selected rows
        For Each dr As DataRow In row
            dtNew.ImportRow(dr)
        Next
        ' return filtered dt
        Return dtNew
    End Function

    Function SelectDataTable4(ByVal dt As DataTable, filter As String) As DataTable
        Dim row As DataRow()
        Dim dtNew As DataTable
        ' copy table structure
        dtNew = dt.Clone()
        ' sort and filter data
        row = dt.Select("(Column3 <> 'Retencion' OR Column3 <> 'Nota de Credito' OR Column3 <> 'Nota de Debito') AND (Column1 = '" & filter & "')")
        ' fill dtNew with selected rows
        For Each dr As DataRow In row
            dtNew.ImportRow(dr)
        Next
        ' return filtered dt
        Return dtNew
    End Function

    Function SelectDataTable5(ByVal dt As DataTable, ByVal filter As String) As DataTable
        Dim row As DataRow()
        Dim dtNew As DataTable
        ' copy table structure
        dtNew = dt.Clone()
        ' sort and filter data
        row = dt.Select("Column10" & "=" & "'" & filter & "'")
        ' fill dtNew with selected rows
        For Each dr As DataRow In row
            dtNew.ImportRow(dr)
        Next
        ' return filtered dt
        Return dtNew
    End Function

    Private Sub Modificar_Stock(ByVal art As String, cant As Double, ByVal conn As connect, ByVal tran As SqlTransaction)
        Dim CMD2 As New SqlCommand
        CMD2.Connection = conn.connName
        CMD2.CommandType = CommandType.StoredProcedure
        CMD2.Transaction = tran
        CMD2.CommandText = "ActualizaStock"
        CMD2.Parameters.Clear()


        'Parametros
        'CMD2.Parameters.Add("@sCo_Alma", SqlDbType.Char)
        CMD2.Parameters.Add("@sCo_Art", SqlDbType.Char)
        CMD2.Parameters.Add("@deCantidad", SqlDbType.Decimal)


        'Valores
        'CMD2.Parameters("@sCo_Alma").Value = "PPAL" 'almacen que se esta modificando
        CMD2.Parameters("@sCo_Art").Value = art
        CMD2.Parameters("@deCantidad").Value = cant


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

    Function ReemplazaGuion(ByVal expr As String) As String
        Dim Res As String

        Res = Replace(expr, "-", "")
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
        CMD2.CommandText = "Merk_VerificarStock"
        CMD2.Parameters.Clear()

        CMD2.Parameters.Add("@coArt", SqlDbType.VarChar)


        CMD2.Parameters("@coArt").Value = filtro

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
        Dim Origen As String = "Interace Merkant"
        'Escribimos en los Registros de Aplicación
        Dim Elog As EventLog
        Elog = New EventLog("Application", Maquina, Origen)
        Elog.WriteEntry(Texto_Evento, tipo_entrada, 100, CType(50, Short))
        Elog.Close()
        Elog.Dispose()
    End Sub

    Private Sub vigilante1_Error(sender As Object, e As ErrorEventArgs) Handles vigilante1.[Error]

        ' Show that an error has been detected.
        'Console.WriteLine("The FileSystemWatcher has detected an error")

        ' Give more information if the error is due to an internal buffer overflow.
        If TypeOf e.GetException Is InternalBufferOverflowException Then
            ' This can happen if Windows is reporting many file system events quickly 
            ' and internal buffer of the  FileSystemWatcher is not large enough to handle this
            ' rate of events. The InternalBufferOverflowException error informs the application
            ' that some of the file system events are being lost.
            EscribirLog(
                "The file system watcher experienced an internal buffer overflow: " _
                + e.GetException.Message, EventLogEntryType.Error)
        End If
    End Sub


End Class
