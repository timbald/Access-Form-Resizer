Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim frmResize As New clsFormResize

Private Sub Form_Load()
    frmResize.init Me.Form
End Sub

Private Sub Form_Resize()
    'Place Holder
End Sub
