Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub SplitButton_Click(sender As Object, e As EventArgs) Handles SplitButton.Click
        Dim stoppingpoint As Integer
        Dim tempstring As String
        Dim prefix As String
        Dim MoreString As Boolean

        stoppingpoint = 0
        MoreString = True

        prefix = "\em " 'prefix can be changed here

        tempstring = txtbxInput.Text 'puts input box text into a variable
        While MoreString = True

            If Len(tempstring) < (500 - Len(prefix)) Then   'if the tempstring will be less than when it needs to be cut off, 
                stoppingpoint = Len(tempstring)             'assign stoppingpoint as length of tempstring so will just write the rest of it
                MoreString = False                          'assign loop variable to false so it stops.
            Else
                stoppingpoint = InStrRev(tempstring.Substring(0, (500 - Len(prefix))), " ") 'Find the last space so the last word isn't chopped in half.
            End If

            txtbxOutput.Text = txtbxOutput.Text & vbCrLf & prefix & tempstring.Substring(0, stoppingpoint) 'keep what's in the output box and write the next line as a new line.
            tempstring = tempstring.Remove(0, stoppingpoint) 'remove that line from the variable so it's not written again.


        End While
    End Sub
End Class
