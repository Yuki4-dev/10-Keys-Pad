
Public Class LongPressButton
    Inherits Button

    Public Property LongPressTimeMsec = 1000
    Public Event LongPress As EventHandler
    Private TokenSource As Threading.CancellationTokenSource = Nothing

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)

        MyBase.WndProc(m)
        Const WM_LBUTTONDOWN = &H201    '左ボタンクリック開始
        Const WM_LBUTTONUP = &H202      '左ボタンクリック終了
        Const WM_LBUTTONDBLCLK = &H203  '左ボタンダブルクリック
        Select Case m.Msg
            Case WM_LBUTTONDOWN
                LongPressDispatch()
            Case WM_LBUTTONUP
                TokenSource?.Cancel()
                TokenSource = Nothing
            Case WM_LBUTTONDBLCLK
                TokenSource?.Cancel()
                TokenSource = Nothing
        End Select
    End Sub

    Private Async Sub LongPressDispatch()

        TokenSource = New Threading.CancellationTokenSource()
        Dim token As Threading.CancellationToken = TokenSource.Token
        Await Task.Factory.StartNew(Sub()
                                        Threading.Thread.Sleep(LongPressTimeMsec)
                                        If Not token.IsCancellationRequested Then
                                            Invoke(Sub() OnLongPress())
                                        End If
                                    End Sub, token)
    End Sub

    Protected Sub OnLongPress()

        RaiseEvent LongPress(Me, New EventArgs())
    End Sub

End Class


