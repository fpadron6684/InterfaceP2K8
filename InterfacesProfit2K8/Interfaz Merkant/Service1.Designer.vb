Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InterfazMerkant
    Inherits System.ServiceProcess.ServiceBase

    'UserService reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' Punto de entrada principal del proceso
    <MTAThread()> _
    <System.Diagnostics.DebuggerNonUserCode()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' Puede que más de un servicio de NT se ejecute con el mismo proceso. Para agregar
        ' otro servicio a este proceso, cambie la siguiente línea para
        ' crear un segundo objeto de servicio. Por ejemplo,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New InterfazMerkant}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Requerido por el Diseñador de componentes
    Private components As System.ComponentModel.IContainer

    ' NOTA: el Diseñador de componentes requiere el siguiente procedimiento
    ' Se puede modificar utilizando el Diseñador de componentes.  
    ' No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.vigilante1 = New System.IO.FileSystemWatcher()
        Me.vigilante2 = New System.IO.FileSystemWatcher()
        Me.vigilante3 = New System.IO.FileSystemWatcher()
        Me.vigilante4 = New System.IO.FileSystemWatcher()
        Me.vigilante5 = New System.IO.FileSystemWatcher()
        CType(Me.vigilante1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vigilante2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vigilante3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vigilante4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vigilante5, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'vigilante1
        '
        Me.vigilante1.EnableRaisingEvents = True
        '
        'vigilante2
        '
        Me.vigilante2.EnableRaisingEvents = True
        '
        'vigilante3
        '
        Me.vigilante3.EnableRaisingEvents = True
        '
        'vigilante4
        '
        Me.vigilante4.EnableRaisingEvents = True
        '
        'vigilante5
        '
        Me.vigilante5.EnableRaisingEvents = True
        '
        'InterfazMerkant
        '
        Me.ServiceName = "InterfazMerkant"
        CType(Me.vigilante1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vigilante2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vigilante3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vigilante4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vigilante5, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Friend WithEvents vigilante1 As System.IO.FileSystemWatcher
    Friend WithEvents vigilante2 As System.IO.FileSystemWatcher
    Friend WithEvents vigilante3 As System.IO.FileSystemWatcher
    Friend WithEvents vigilante4 As System.IO.FileSystemWatcher
    Friend WithEvents vigilante5 As System.IO.FileSystemWatcher

End Class
