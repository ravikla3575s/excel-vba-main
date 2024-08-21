VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Pharmacist
Option Explicit

Private pEmployeeNumber As Long
Private pPharmacistName As String
Private pWorkHour As Single

' 社員番号プロパティ (Long型)
Public Property Get EmployeeNumber() As Long
    EmployeeNumber = pEmployeeNumber
End Property

Public Property Let EmployeeNumber(ByVal value As Long)
    pEmployeeNumber = value
End Property

' 薬剤師氏名プロパティ
Public Property Get PharmacistName() As String
    PharmacistName = pPharmacistName
End Property

Public Property Let PharmacistName(ByVal value As String)
    pPharmacistName = value
End Property

' 週あたり勤務時間プロパティ (Single型)
Public Property Get WorkHour() As Single
    WorkHour = pWorkHour
End Property

Public Property Let WorkHour(ByVal value As Single)
    pWorkHour = value
End Property
