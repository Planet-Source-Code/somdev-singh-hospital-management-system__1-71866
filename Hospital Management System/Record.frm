VERSION 5.00
Begin VB.MDIForm Record 
   BackColor       =   &H8000000F&
   Caption         =   "Hospital Management System"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8550
   Icon            =   "Record.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "MDIForm1"
   Picture         =   "Record.frx":5ADA
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuAuthentication 
      Caption         =   "&Authentication"
      Begin VB.Menu mnuUserManagement 
         Caption         =   "User Management"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEmployee 
      Caption         =   "Employee"
      Begin VB.Menu mnuEmployeeDetails 
         Caption         =   "Employee Details"
      End
      Begin VB.Menu mnuPaymentDetails 
         Caption         =   "Employee Payment Details"
      End
      Begin VB.Menu mnuPaymentSlip 
         Caption         =   "Employee Payment Slip"
      End
   End
   Begin VB.Menu mnuInquiry 
      Caption         =   "&Inquiry"
      Begin VB.Menu mnuAvailable 
         Caption         =   "Beds Availability"
      End
      Begin VB.Menu mnuFaculies 
         Caption         =   "Faculties"
      End
      Begin VB.Menu mnuFacilities 
         Caption         =   "Facilities"
      End
      Begin VB.Menu mnuContact 
         Caption         =   "Contact"
      End
   End
   Begin VB.Menu mnuRegistration 
      Caption         =   "&Registration"
      Begin VB.Menu mnuAdInt 
         Caption         =   "Admission Intimation"
      End
      Begin VB.Menu mnuDisInt 
         Caption         =   "Discharge Intimation"
      End
   End
   Begin VB.Menu Transaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuTransInt 
         Caption         =   "Transfer Intimation"
      End
   End
   Begin VB.Menu mnuBilling 
      Caption         =   "&Billing"
      Begin VB.Menu mnuOther 
         Caption         =   "Test Bill"
      End
      Begin VB.Menu tbill 
         Caption         =   "Treatment Bill"
      End
      Begin VB.Menu mnuDisBill 
         Caption         =   "Discharge Bill"
      End
   End
   Begin VB.Menu mnuRecord 
      Caption         =   "R&ecord"
      Begin VB.Menu mnuAllRecord 
         Caption         =   "All Record"
      End
   End
   Begin VB.Menu Info 
      Caption         =   "Re&port"
      Begin VB.Menu Report1 
         Caption         =   "Total Entry Report"
      End
      Begin VB.Menu Report2 
         Caption         =   "Treatment Transaction Report"
      End
      Begin VB.Menu ptransaction 
         Caption         =   "Patient's Transaction"
      End
      Begin VB.Menu Report4 
         Caption         =   "Employee Datails "
      End
      Begin VB.Menu Report5 
         Caption         =   "Employee Payment Slip"
      End
   End
   Begin VB.Menu mnuutilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnucalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuwordpad 
         Caption         =   "WordPad"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Hel&p"
      Begin VB.Menu mnuInstructions 
         Caption         =   "Instructions"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Us"
      End
   End
End
Attribute VB_Name = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddBill_Click()
'AdmissionBill.Show
End Sub

Private Sub AddInt_Click()

End Sub


Private Sub mnuAddBill_Click()

End Sub

Private Sub mnuAddInt_Click()
AdmissionInt.Show
End Sub

'Private Sub mnuAdBill_Click()
'AdmissionMemo.Show

'AdmissionMemo.Text1.Text = ""
'AdmissionMemo.Text2.Text = ""
'AdmissionMemo.Text3.Text = ""
'AdmissionMemo.Text4.Text = ""

'AdmissionMemo.Text6.Text = ""
'AdmissionMemo.Text7.Text = ""
'AdmissionMemo.Text8.Text = ""

'AdmissionMemo.Text10.Text = ""
'AdmissionMemo.Text11.Text = ""
'AdmissionMemo.Text12.Text = ""
'End Sub

Private Sub MDIForm_Load()
'Form2.Show
'Width = 12000
'Height = 9000
End Sub

Private Sub mnuAbout_Click()
Aus.Show
End Sub

Private Sub mnuAdInt_Click()
Call UNLOADFORMS
AdmissionInt.Show
End Sub



Private Sub mnuAllRecord_Click()
Call UNLOADFORMS
Oldrecord.Show
End Sub

Private Sub mnuAvailable_Click()
Call UNLOADFORMS
Availability.Show

End Sub

Private Sub Contact_Click()
'Load Contact
Contact.Show
End Sub

Private Sub DisInt_Click()
DischargeInt.Show
End Sub

Private Sub Exit_Click()
End
End Sub

'Private Sub mnuBioChe_Click()
'BioChemical.Show
'End Sub

'Private Sub mnuBioTech_Click()
'BioTechnical.Show
'End Sub

Private Sub mnuBloodB_Click()
TestTransaction.Show
End Sub

Private Sub mnucalculator_Click()
Shell "calc.exe"
End Sub





Private Sub mnuContact_Click()
Call UNLOADFORMS
Contact.Show
End Sub

Private Sub mnuDisBill_Click()
Call UNLOADFORMS
End Sub

'Private Sub mnuCRecord_Click()
'Currentrecord.Show
'End Sub

'Private Sub mnuDisBill_Click()
'DischargeBill.Show
'End Sub

Private Sub mnuDisInt_Click()
Call UNLOADFORMS
DischargeInt.Show
End Sub

Private Sub mnuEmployeeDetails_Click()
Call UNLOADFORMS
EmployeeDetails.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuFacilities_Click()
Call UNLOADFORMS
Facilities.Show
End Sub



Private Sub mnuFaculies_Click()
Call UNLOADFORMS
Faculty.Show
End Sub



Private Sub mnuInstructions_Click()
Instructions.Show
End Sub

Private Sub mnuLogOff_Click()
MsgBox "Are you Sure to log off"
frmSecurity.Show
End Sub

Private Sub mnuOther_Click()
Call UNLOADFORMS
TestTransaction.Show
End Sub



Private Sub mnuPaymentDetails_Click()
Call UNLOADFORMS
Emp_PaymentDetails.Show
End Sub

Private Sub mnuPaymentSlip_Click()
Call UNLOADFORMS
PaymentSlip.Show
End Sub

Private Sub mnuTransInt_Click()
Call UNLOADFORMS
TransferInt.Show
TransferInt.Frame1.Enabled = False
'TransferInt.Command1.Visible = False
'TransferInt.Command2.Visible = False
'TransferInt.Label13.Visible = True

'TransferInt.Command5.Visible = False
'TransferInt.Command6.Visible = False

'TransferInt.Text3.Visible = True
'TransferInt.Text3.Enabled = True
'TransferInt.Text3.SetFocus
End Sub



Private Sub r_Click()

End Sub

Private Sub mnuUserManagement_Click()
Call UNLOADFORMS
frmSecurity.Show
End Sub

Private Sub mnuwordpad_Click()
Shell "wordpad.exe"
End Sub

Private Sub ptransaction_Click()
Call UNLOADFORMS
DataReport3.Show
End Sub

Private Sub Report1_Click()
Call UNLOADFORMS
DataReport1.Show
End Sub

Private Sub Report2_Click()
Call UNLOADFORMS
DataReport2.Show
End Sub

Private Sub Report4_Click()
Call UNLOADFORMS
DataReport4.Show
End Sub

Private Sub Report5_Click()
Call UNLOADFORMS
DataReport5.Show
End Sub

Private Sub tbill_Click()
Call UNLOADFORMS
TreatmentBill.Show
End Sub

'Private Sub TreatmentReport_Click()
'TestReport.Show
'End Sub
Public Sub UNLOADFORMS()

End Sub
