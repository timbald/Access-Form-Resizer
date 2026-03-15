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
    frmResize.ResizeNow 'If this is a subform and you notice any issues when maximising and restoring, use this line of code to force the resize code to fire once more
End Sub
