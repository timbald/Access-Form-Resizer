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

Private Sub Label1_DblClick(Cancel As Integer)

    DoCmd.MoveSize Me.Form.WindowLeft * 0.9, Me.Form.WindowTop * 0.9, Me.Form.WindowWidth * 1.1, Me.Form.WindowHeight * 1.1

End Sub
