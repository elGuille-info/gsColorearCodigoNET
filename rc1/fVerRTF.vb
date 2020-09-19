'------------------------------------------------------------------------------
' Ver el código RTF y el resultado                                  (08/Feb/07)
'
' Cambios para que se muestre el RTF después de hacer cambios       (10/Sep/20)
' también se mostrará el texto coloreado al pulsar en RTF
'
' ©Guillermo 'guille' Som, 2007, 2020
'------------------------------------------------------------------------------

Public Class fVerRTF

    Private m_Texto As String
    ''' <summary>
    ''' Asignar el texto en formato RTF a mostrar
    ''' </summary>
    Public Property Texto() As String
        Get
            Return m_Texto
        End Get
        Set(ByVal value As String)
            m_Texto = value
            Me.RichTextBox1.Rtf = value
        End Set
    End Property

    Private Sub btnText_Click(sender As System.Object,
                              e As System.EventArgs) Handles btnText.Click
        Me.RichTextBox1.Text = m_Texto ' Texto
    End Sub

    Private Sub btnRTF_Click(sender As System.Object,
                             e As System.EventArgs) Handles btnRTF.Click
        Try
            Me.RichTextBox1.Rtf = Me.RichTextBox1.Text ' Texto
            m_Texto = RichTextBox1.Rtf
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try
    End Sub
End Class