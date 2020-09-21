'------------------------------------------------------------------------------
' Clase parcial para asignar las imágenes a los menús e iconos      (12/Sep/20)
' ya que el diseñador de la preview no lo maneja bien en el .Designer
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Linq

Partial Public Class fColorear

    Private Sub AsignarImagenes()
        Me.mnuSintaxColorear.Image = New System.Drawing.Bitmap(".\Resources\ColorHS.png")
        Me.mnuFicAbrir.Image = New System.Drawing.Bitmap(".\Resources\openHS.png")
        Me.mnuFicGuardar.Image = New System.Drawing.Bitmap(".\Resources\saveHS.png")
        Me.mnuFicNavegar.Image = New System.Drawing.Bitmap(".\Resources\Web.png")
        Me.mnuFicAcerca.Image = New System.Drawing.Bitmap(".\Resources\Info.png")
        Me.mnuFicSalir.Image = New System.Drawing.Bitmap(".\Resources\XP_Cerrar.png")
        Me.mnuEdiDeshacer.Image = New System.Drawing.Bitmap(".\Resources\Edit_UndoHS.png")
        Me.mnuEdiCortar.Image = New System.Drawing.Bitmap(".\Resources\CutHS.png")
        Me.mnuEdiCopiar.Image = New System.Drawing.Bitmap(".\Resources\CopyHS.png")
        Me.mnuEdiPegar.Image = New System.Drawing.Bitmap(".\Resources\PasteHS.png")
        Me.btnColorear.Image = New System.Drawing.Bitmap(".\Resources\ColorHS.png")
        Me.mnuSintax_dotNet.Image = New System.Drawing.Bitmap(".\Resources\Hoja_visual_studio.png")
        Me.mnuSintax_VB.Image = New System.Drawing.Bitmap(".\Resources\Hoja_vb.png")
        Me.mnuSintax_CS.Image = New System.Drawing.Bitmap(".\Resources\Hoja_cs.png")
        Me.mnuSintax_Java.Image = New System.Drawing.Bitmap(".\Resources\Hoja_Java.png")
        Me.mnuSintax_FSharp.Image = New System.Drawing.Bitmap(".\Resources\Hoja_FSharp.png")
        Me.mnuSintax_IL.Image = New System.Drawing.Bitmap(".\Resources\Hoja_visual_studio.png")
        Me.mnuSintax_CPP.Image = New System.Drawing.Bitmap(".\Resources\Hoja_cpp.png")
        Me.mnuSintax_Pascal.Image = New System.Drawing.Bitmap(".\Resources\Hoja_p.png")
        Me.mnuSintax_SQL.Image = New System.Drawing.Bitmap(".\Resources\Hoja_SQL.png")
        Me.mnuSintax_VB6.Image = New System.Drawing.Bitmap(".\Resources\Hoja_VB6.png")
        Me.mnuSintax_XML.Image = New System.Drawing.Bitmap(".\Resources\XMLFileHS.png")
        Me.tsbColorear.Image = New System.Drawing.Bitmap(".\Resources\ColorHS.png")
        Me.tsbAbrir.Image = New System.Drawing.Bitmap(".\Resources\openHS.png")
        Me.tsbGuardar.Image = New System.Drawing.Bitmap(".\Resources\saveHS.png")
        Me.tsbCortar.Image = New System.Drawing.Bitmap(".\Resources\CutHS.png")
        Me.tsbCopiar.Image = New System.Drawing.Bitmap(".\Resources\CopyHS.png")
        Me.tsbPegar.Image = New System.Drawing.Bitmap(".\Resources\PasteHS.png")
        Me.tsbDeshacer.Image = New System.Drawing.Bitmap(".\Resources\Edit_UndoHS.png")
        Me.tsbSintax.Image = New System.Drawing.Bitmap(".\Resources\Hoja_colores.png")
        Me.tsbNavegar.Image = New System.Drawing.Bitmap(".\Resources\Web.png")
        Me.tsbAcerca.Image = New System.Drawing.Bitmap(".\Resources\Info.png")
        Me.tsbSalir.Image = New System.Drawing.Bitmap(".\Resources\XP_Cerrar.png")

    End Sub

End Class
