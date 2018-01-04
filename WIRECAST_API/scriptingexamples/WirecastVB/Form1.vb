Public Class Form1
    Inherits System.Windows.Forms.Form



#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(208, 224)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 24)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Go"
        '
        'ListBox1
        '
        Me.ListBox1.Location = New System.Drawing.Point(8, 8)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(272, 186)
        Me.ListBox1.TabIndex = 1
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(8, 224)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(88, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Speaker Mute"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Function GetWirecast() As Object
        Try
            Return GetObject(Class:="Wirecast.Application")
        Catch
            Return Nothing
        End Try
    End Function

    Private Function GetDocument(ByVal doc_idx) As Object
        Dim objWirecast As Object
        Dim objDoc As Object

        objWirecast = GetWirecast()
        If (Not objWirecast Is Nothing) Then
            objDoc = objWirecast.DocumentByIndex(doc_idx)
        End If

        Return objDoc
    End Function

    Private Function GetLayer(ByVal doc_idx, ByVal layer_name) As Object
        Dim objDoc As Object
        Dim objLayer As Object

        objDoc = GetDocument(doc_idx)
        If (Not objDoc Is Nothing) Then
            objLayer = objDoc.LayerByName(layer_name)
        End If

        Return objLayer
    End Function


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim objLayer As Object
        Dim objDocument As Object

        objDocument = GetDocument(1)
        objLayer = GetLayer(1, "normal")

        If (objDocument Is Nothing) Then
            MsgBox("Please Launch Wirecast before opening this Example")
        End If

        If (Not objLayer Is Nothing) Then
            Dim count As Integer
            Dim incr As Integer

            count = objLayer.ShotCount()
            For incr = 1 To count
                Dim shot_id As Integer

                shot_id = objLayer.ShotIDByIndex(incr)
                If (shot_id <> 0) Then
                    Dim shot As Object
                    shot = objDocument.ShotByShotID(shot_id)
                    If (Not shot Is Nothing) Then
                        Dim name As String
                        name = shot.Name
                        ListBox1.Items.Add(name)
                    End If
                End If
            Next incr
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim objLayer As Object
        objLayer = GetLayer(1, "normal")

        Dim shot_idx As Integer
        shot_idx = ListBox1.SelectedIndex() + 1
        If (shot_idx > 0) And (Not objLayer Is Nothing) Then
            Dim shot_id As Integer

            shot_id = objLayer.ShotIDByIndex(shot_idx)
            objLayer.ActiveShotID = shot_id
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim objDoc As Object
        objDoc = GetDocument(1)
        If (Not objDoc Is Nothing) Then
            If (objDoc.AudioMutedToSpeaker <> 0) Then
                objDoc.AudioMutedToSpeaker = 0
            Else
                objDoc.AudioMutedToSpeaker = 1
            End If
        End If
    End Sub
End Class
