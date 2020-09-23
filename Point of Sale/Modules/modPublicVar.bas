Attribute VB_Name = "modPublicVar"
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

Option Explicit

'Public InvalidDB                    As Boolean
Public CurrUser                     As USER_INFO
Public DBPath                       As String
Public Enc                          As New clsBlowfish
Public CurrBiz                      As BUSINESS_INFO

Public CN                           As New Connection

Public ReportType                   As String
