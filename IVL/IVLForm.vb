Imports NationalInstruments.NI4882

Public Class IVLForm
    Private Sub IVLForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.FolderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer
        If System.IO.File.Exists("S:\IVL\B-function.txt") Then
            GetBValues("S:\IVL\B-function.txt")
        End If
        Me.VListBox.Items.Add("0")
        Me.VListBox.SetSelected(0, True)


        ' FindGPIB()
        'If GPIBPresent Then
        'SplashForm.Show()
        'SplashForm.Label1.Text = "Checking for Agilent source"
        'SplashForm.Refresh()
        'FindAgilent()
        'SplashForm.Label1.Text = "Checking for Newport"
        'SplashForm.Refresh()
        'FindNewport()
        'SplashForm.Label1.Text = "Checking for Siglent source"
        'SplashForm.Refresh()
        'FindSiglent()
        'SplashForm.Label1.Text = "Checking for Gaussmeter"
        'SplashForm.Refresh()

        'SplashForm.Hide()

        'Else
        'Me.TabControl1.TabPages.Remove(Magnetoresistance)
        'End If



    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close() 'Exits programme
    End Sub







    Private Sub MRButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MRButton.Click
        UltraMR()
    End Sub





    Private Sub VLowBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles VLowBox.KeyPress
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) OrElse e.KeyChar = "-" Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub VHighBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles VHighBox.KeyPress
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) OrElse e.KeyChar = "-" Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub VStepBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles VStepBox.KeyPress
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub


    Private Sub ComplianceIBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComplianceIBox.KeyPress
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub ComplianceVBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComplianceVBox.KeyPress
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub



    Private Sub VLowBox_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles VLowBox.MouseClick
        Me.VLowBox.SelectAll()
    End Sub



    Private Sub VHighBox_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles VHighBox.MouseClick
        Me.VHighBox.SelectAll()
    End Sub



    Private Sub VStepBox_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles VStepBox.MouseClick
        Me.VStepBox.SelectAll()
    End Sub



    Private Sub ComplianceVBox_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ComplianceVBox.MouseClick
        ComplianceVBox.SelectAll()
    End Sub



    Private Sub ComplianceIBox_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ComplianceIBox.MouseClick
        ComplianceIBox.SelectAll()
    End Sub

    Private Sub SourceVoltageBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SourceVoltageBox.KeyPress
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) OrElse e.KeyChar = "-" Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub SourceVoltageBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SourceVoltageBox.LostFocus
        If Val(Me.SourceVoltageBox.Text) = 0 Then Me.SourceVoltageBox.Text = "0"
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If Me.ScanButton.Enabled = False And _
       TabControl1.SelectedTab Is Magnetoresistance Then
            TabControl1.SelectedTab = IVL
        End If
        If TabControl1.SelectedTab Is IVL Then
            Me.CancelButton = AbortButton
        ElseIf TabControl1.SelectedTab Is Magnetoresistance Then
            Me.CancelButton = AbortMRButton
        End If
    End Sub





    Private Sub SelectDirectoryButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectDirectoryButton.Click
        If Me.FolderBrowserDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            UserDirectory = Me.FolderBrowserDialog.SelectedPath
        End If

    End Sub

    Private Sub SelectBFieldsButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectBFieldsButton.Click
        OpenFileDialog.Filter = "TXT Files (*.txt)|*.txt"
        If OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            GetBValues(OpenFileDialog.FileName)
        End If
    End Sub

    Private Sub SelectVButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectVButton.Click
        OpenFileDialog.Filter = "TXT Files (*.txt)|*.txt"
        If OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            GetVvalues(OpenFileDialog.FileName)
        End If
    End Sub



    Private Sub BFieldAccuracyBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub



    Private Sub GaussPBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub GaussIBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub MeasureDelayBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) OrElse e.KeyChar = CChar(ChrW(Keys.Delete)) OrElse e.KeyChar = CChar(ChrW(Keys.Back)) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub






    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Me.FolderBrowserDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            UserDirectory = Me.FolderBrowserDialog.SelectedPath
        End If
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs)
        If Me.FolderBrowserDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            UserDirectory = Me.FolderBrowserDialog.SelectedPath
        End If
    End Sub




    Private Sub Button7_Click(sender As Object, e As EventArgs)
        If Me.FolderBrowserDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            UserDirectory = Me.FolderBrowserDialog.SelectedPath
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        B2_trans()
    End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If Me.FolderBrowserDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            UserDirectory = Me.FolderBrowserDialog.SelectedPath
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        B1_trans()
    End Sub

    Private Sub Button4_Click_1(sender As System.Object, e As System.EventArgs)
        B_I()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        B_t2()
    End Sub

    

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs)
        NewUltraMR()
    End Sub

    Private Sub Button7_Click_1(sender As System.Object, e As System.EventArgs)
        NewUltraMR2()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        NewUltraMR3()
    End Sub

    Private Sub Button10_Click(sender As System.Object, e As System.EventArgs) Handles Button10.Click
        NewUltraMR4()
    End Sub

    Private Sub Button11_Click(sender As System.Object, e As System.EventArgs) Handles Button11.Click
        NewUltraMR5()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        NewUltraMR6()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs)
        NewUltraMR7()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        NewUltraMR8()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        NewUltraMR9()
    End Sub

    Private Sub Button16_Click(sender As System.Object, e As System.EventArgs) Handles Button16.Click
        NewUltraMR10()
    End Sub

    Private Sub Button17_Click(sender As System.Object, e As System.EventArgs) Handles Button17.Click
        NewUltraMR11()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        NewUltraMR12()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        NewUltraMR13()
    End Sub

    Private Sub Button20_Click(sender As System.Object, e As System.EventArgs) Handles Button20.Click
        NewUltraMR14()
    End Sub

    
    Private Sub Button21_Click(sender As System.Object, e As System.EventArgs) Handles Button21.Click
        NewUltraMR15()
    End Sub

    Private Sub Button22_Click(sender As System.Object, e As System.EventArgs) Handles Button22.Click
        NewUltraMR16()
    End Sub

    Private Sub Button23_Click(sender As System.Object, e As System.EventArgs) Handles Button23.Click
        NewUltraMR17()
    End Sub

    Private Sub Button24_Click(sender As System.Object, e As System.EventArgs) Handles Button24.Click
        NewUltraMR18()
    End Sub

    Private Sub Button25_Click(sender As System.Object, e As System.EventArgs)
        TEST_ZMEAS()
    End Sub

    Private Sub Button26_Click(sender As System.Object, e As System.EventArgs) Handles Button26.Click
        NewUltraMR19()
    End Sub

    Private Sub Button27_Click(sender As System.Object, e As System.EventArgs) Handles Button27.Click
        B_IRelationTest()
    End Sub

    Private Sub Button25_Click_1(sender As System.Object, e As System.EventArgs) Handles Button25.Click
        NewUltraMR20()
    End Sub

    Private Sub Button28_Click(sender As System.Object, e As System.EventArgs) Handles Button28.Click
        NewUltraMR21()
    End Sub

    Private Sub Button29_Click(sender As System.Object, e As System.EventArgs) Handles Button29.Click
        NewUltraMR22()
    End Sub

    Private Sub Button30_Click(sender As System.Object, e As System.EventArgs) Handles Button30.Click
        NewUltraMR23()
    End Sub

    Private Sub Button31_Click(sender As System.Object, e As System.EventArgs) Handles Button31.Click
        NewUltraMR24()
    End Sub

    Private Sub Button32_Click(sender As System.Object, e As System.EventArgs) Handles Button32.Click
        NewUltraMR25()
    End Sub

    Private Sub Button33_Click(sender As System.Object, e As System.EventArgs) Handles Button33.Click
        NewUltraMR26()
    End Sub

    Private Sub Button34_Click(sender As System.Object, e As System.EventArgs) Handles Button34.Click
        LargeMR()
    End Sub

    Private Sub Button35_Click(sender As System.Object, e As System.EventArgs) Handles Button35.Click
        NewUltraMR27()
    End Sub

    Private Sub Button36_Click(sender As System.Object, e As System.EventArgs) Handles Button36.Click
        Key2612ON()
    End Sub

    Private Sub Button4_Click_2(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Bsmllcoil_I()
    End Sub

    Private Sub Button37_Click(sender As System.Object, e As System.EventArgs) Handles Button37.Click
        NewUltraMR29()
    End Sub

    Private Sub Button38_Click(sender As System.Object, e As System.EventArgs) Handles Button38.Click
        NewUltraMR30()
    End Sub

    Private Sub Button39_Click(sender As System.Object, e As System.EventArgs)
        NewUltraMR31()
    End Sub

    Private Sub Button40_Click(sender As System.Object, e As System.EventArgs)
        NewUltraMR32()
    End Sub

    Private Sub Button41_Click(sender As System.Object, e As System.EventArgs)
        NewUltraMR33()
    End Sub

    Private Sub Button42_Click(sender As System.Object, e As System.EventArgs)
        NewUltraMR34()
    End Sub

 

    Private Sub Button44_Click(sender As System.Object, e As System.EventArgs) Handles Button44.Click
        test150818()
    End Sub

    Private Sub Button45_Click(sender As System.Object, e As System.EventArgs) Handles Button45.Click
        test170818()
    End Sub

    Private Sub Button46_Click(sender As System.Object, e As System.EventArgs) Handles Button46.Click
        DeviceCond(2)
    End Sub

    Private Sub Button43_Click(sender As System.Object, e As System.EventArgs) Handles Button43.Click
        NewUltraMR35()
    End Sub

    Private Sub Button7_Click_2(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        IVLmeas()
    End Sub

    Private Sub Button8_Click_1(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        NewUltraMR36()
    End Sub
End Class
