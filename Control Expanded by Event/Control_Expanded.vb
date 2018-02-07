    'restore
    Private Sub Test_uc1_MouseLeave(sender As Object, e As EventArgs) Handles Test_uc1.MouseLeave
        Dim senderUserControl = CType(sender, UserControl)
        senderUserControl.Width = 93
        senderUserControl.Height = 39
    End Sub

    'expand
    Private Sub Test_uc1_MouseClick(sender As Object, e As MouseEventArgs) Handles Test_uc1.MouseClick
        Dim senderUserControl = CType(sender, UserControl)
        senderUserControl.Width = 217
        senderUserControl.Height = 79
    End Sub
