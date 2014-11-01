<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.excelButton = New System.Windows.Forms.Button()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.runscriptButton = New System.Windows.Forms.Button()
        Me.wordButton = New System.Windows.Forms.Button()
        Me.OpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "Get Excel file!"
        Me.OpenFileDialog1.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm|All files|*.*"
        '
        'excelButton
        '
        Me.excelButton.Location = New System.Drawing.Point(235, 336)
        Me.excelButton.Name = "excelButton"
        Me.excelButton.Size = New System.Drawing.Size(131, 23)
        Me.excelButton.TabIndex = 2
        Me.excelButton.Text = "Choose Excel Input"
        Me.excelButton.UseVisualStyleBackColor = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(12, 12)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(355, 318)
        Me.RichTextBox1.TabIndex = 2
        Me.RichTextBox1.Text = resources.GetString("RichTextBox1.Text")
        '
        'runscriptButton
        '
        Me.runscriptButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.runscriptButton.Location = New System.Drawing.Point(15, 367)
        Me.runscriptButton.Name = "runscriptButton"
        Me.runscriptButton.Size = New System.Drawing.Size(130, 23)
        Me.runscriptButton.TabIndex = 4
        Me.runscriptButton.Text = "Run Script"
        Me.runscriptButton.UseVisualStyleBackColor = True
        '
        'wordButton
        '
        Me.wordButton.Location = New System.Drawing.Point(235, 365)
        Me.wordButton.Name = "wordButton"
        Me.wordButton.Size = New System.Drawing.Size(131, 23)
        Me.wordButton.TabIndex = 3
        Me.wordButton.Text = "Choose Word Template"
        Me.wordButton.UseVisualStyleBackColor = True
        '
        'OpenFileDialog2
        '
        Me.OpenFileDialog2.FileName = "Get Word file!"
        Me.OpenFileDialog2.Filter = "Word files|*.docx;*.doc|All files|*.*"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(378, 397)
        Me.Controls.Add(Me.wordButton)
        Me.Controls.Add(Me.runscriptButton)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.excelButton)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "EzLetter Creator"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents excelButton As System.Windows.Forms.Button
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents runscriptButton As System.Windows.Forms.Button
    Friend WithEvents wordButton As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog2 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument

End Class
