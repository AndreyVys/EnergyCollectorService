<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class EnergyCollectorService
    Inherits System.ServiceProcess.ServiceBase

    'UserService переопределяет метод Dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' Главная точка входа процесса
    <MTAThread()>
    <System.Diagnostics.DebuggerNonUserCode()>
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' В одном процессе может выполняться несколько служб NT. Для добавления
        ' службы в процесс измените следующую строку,
        ' чтобы создавался второй объект службы. Например,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New EnergyCollectorService}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Является обязательной для конструктора компонентов
    Private components As System.ComponentModel.IContainer

    ' Примечание: следующая процедура является обязательной для конструктора компонентов
    ' Для ее изменения используйте конструктор компонентов.  
    ' Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.SerialPort = New System.IO.Ports.SerialPort(Me.components)
        Me.EventLog1 = New System.Diagnostics.EventLog()
        Me.ModbusPort = New System.IO.Ports.SerialPort(Me.components)
        CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'EventLog1
        '
        Me.EventLog1.Log = "EnergyCollector"
        Me.EventLog1.Source = "EnergyCollector"
        '
        'ModbusPort
        '
        Me.ModbusPort.BaudRate = 115200
        '
        'EnergyCollectorService
        '
        Me.ServiceName = "EnergyCollectorService"
        CType(Me.EventLog1, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Friend WithEvents SerialPort As IO.Ports.SerialPort
    Friend WithEvents EventLog1 As EventLog
    Friend WithEvents ModbusPort As IO.Ports.SerialPort
End Class
