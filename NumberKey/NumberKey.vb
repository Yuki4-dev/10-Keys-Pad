
Public Class NumberKey

    Private _PutLiteral As String = String.Empty
    Public Property KeySendType As NumberKeySendType = NumberKeySendType.SENDKEYSTROKE

    Public ReadOnly Property PutLiteral As String
        Get
            Return _PutLiteral
        End Get
    End Property

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim WS_EX_NOACTIVATE = &H8000000
            Dim p As CreateParams = MyBase.CreateParams
            p.ExStyle = p.ExStyle Or WS_EX_NOACTIVATE
            Return p
        End Get
    End Property

    Public Event Input(sender As Object, e As EventArgs)
    Public Event BackSpaceInput(sender As Object, e As EventArgs)
    Public Event ClearInput(sender As Object, e As EventArgs)
    Public Event EnterInput(sender As Object, e As EventArgs)

    Public Sub New()
        InitializeComponent()

        For Each b In TableLayoutPanel1.Controls.OfType(Of Button)()
            AddHandler b.Click, AddressOf InputKey
        Next

        AddHandler BackSpaceButton.LongPress, AddressOf OnClearInput
    End Sub

    Private Sub NumberKey_Load(sender As Object, e As EventArgs) Handles Me.Load

        ComandLineArgsInput()
        ChangeTextSize()
    End Sub

    Private Sub Form_Closed(sender As Object, e As EventArgs) Handles Me.Closed

        My.MySettings.Default.NuberKeySize = Size
        My.MySettings.Default.NumberKeyFont = Font
    End Sub
    Private Sub Size_Changed(sender As Object, e As EventArgs) Handles Me.SizeChanged
        ChangeTextSize()
    End Sub

    Private Sub ComandLineArgsInput()

        Dim cmdArgs As String() = System.Environment.GetCommandLineArgs()
        Select Case cmdArgs.Length
            Case 2
                Me.Height = cmdArgs(1)
            Case 3
                Me.Size = New Size(cmdArgs(2), cmdArgs(1))
            Case 5
                Me.Size = New Size(cmdArgs(2), cmdArgs(1))
                Dim dispHeight As Integer = Screen.FromControl(Me).Bounds.Height - 100
                Dim dispWidth As Integer = Screen.FromControl(Me).Bounds.Width - 100
                If dispHeight > cmdArgs(4) AndAlso dispWidth > cmdArgs(3) Then
                    Me.Location = New Point(cmdArgs(3), cmdArgs(4))
                End If
        End Select

    End Sub

    Private Sub InputKey(sender As Object, e As EventArgs)

        Me.TopMost = True

        Dim pushButtonText As String = DirectCast(sender, Button).Tag
        Try
            If Not KeyActionInput(pushButtonText) Then
                OnLiteralInput(pushButtonText)
            End If
        Catch ex As Exception
            MessageBox.Show($"キーアクションを実行できませんでした。: {pushButtonText}")
        End Try

        Me.TopMost = False
    End Sub

    Private Function KeyActionInput(ByVal inputStr As String) As Boolean

        Select Case inputStr
            Case BackSpaceButton.Tag
                OnBackSpaceInput()
            Case EnterButton.Tag
                OnEnterInput()
            Case Else
                Return False
        End Select

        Return True
    End Function

    Private Sub ChangeTextSize()

        Dim fontSize = MyBase.Height / 17
        Dim enterFontSize = MyBase.Height / 25
        SuspendLayout()
        Font = New Font(Font.FontFamily, fontSize, Font.Style, Font.Unit, CType(128, Byte))
        EnterButton.Font = New Font(Font.FontFamily, enterFontSize, Font.Style, Font.Unit, CType(128, Byte))
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Protected Sub OnLiteralInput(ByVal inputStr As String)

        If KeySendType = NumberKeySendType.SENDKEYSTROKE Then
            SendKeys.Send("{" + inputStr + "}")
        ElseIf KeySendType = NumberKeySendType.ONEVENT Then
            _PutLiteral = inputStr
            RaiseEvent Input(Me, New EventArgs())
        End If
    End Sub

    Protected Sub OnBackSpaceInput()

        If KeySendType = NumberKeySendType.SENDKEYSTROKE Then
            SendKeys.Send("{BACKSPACE}")
        ElseIf KeySendType = NumberKeySendType.ONEVENT Then
            RaiseEvent BackSpaceInput(Me, New EventArgs())
        End If
    End Sub

    Protected Sub OnEnterInput()

        If KeySendType = NumberKeySendType.SENDKEYSTROKE Then
            SendKeys.Send("{ENTER}")
        ElseIf KeySendType = NumberKeySendType.ONEVENT Then
            RaiseEvent EnterInput(Me, New EventArgs())
        End If
    End Sub

    Protected Sub OnClearInput()

        If KeySendType = NumberKeySendType.SENDKEYSTROKE Then
            SendKeys.SendWait("^a")
            Threading.Thread.Yield()
            'SendKeys.Send("{DEL}")
            SendKeys.Send("{BACKSPACE}")
        ElseIf KeySendType = NumberKeySendType.ONEVENT Then
            RaiseEvent ClearInput(Me, New EventArgs())
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_CAPTURECHANGED = &H215     'ウィンドウキャプチャ
        Const WM_NCMOUSELEAVE = &H2A2       'マウス離れる

        If m.Msg = WM_CAPTURECHANGED Then
            Me.TopMost = True
        ElseIf m.Msg = WM_NCMOUSELEAVE Then
            Me.TopMost = False
        End If
        MyBase.WndProc(m)
    End Sub

    Public Enum NumberKeySendType
        ONEVENT
        SENDKEYSTROKE
    End Enum

End Class


