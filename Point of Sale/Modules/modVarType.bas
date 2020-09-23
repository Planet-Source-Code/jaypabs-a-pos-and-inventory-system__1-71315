Attribute VB_Name = "modVarType"
''*****************************************************************
'' File Name:
'' Purpose:
'' Required Files:
''
'' Programmer: Philip V. Naparan   E-mail: philipnaparan@yahoo.com
'' Date Created:
'' Last Modified:
'' Modified By:
'' Credits: NONE, ALL CODES ARE CODED BY Philip V. Naparan
''*****************************************************************

'Some of the code are being modified in order to fit with my Point of Sale and Inventory System program.
'For more source code please visit my website at http://www.sourcecodester.com/

'Variable structure for user
Public Type USER_INFO
    USER_PK As Long
    USER_NAME As String
    USER_ISADMIN As Boolean
End Type

'Enumerator for form state
Public Enum FormState
    adStateAddMode = 0
    adStateEditMode = 1
    adStatePopupMode = 2
End Enum

Public Type BUSINESS_INFO
    BUSINESS_NAME As String
    BUSINESS_ADDRESS As String
    BUSINESS_CONTACT_INFO As String
End Type
