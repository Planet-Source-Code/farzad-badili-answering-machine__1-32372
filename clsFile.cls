VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*A little out of place, simple object just returns a collection of .wav files.
Public Function GetMessages() As Collection

Dim oColl As New Collection
Dim myString As String

    myString = Dir(App.path & "\Messages\*.wav")
    Do While myString <> ""
        oColl.Add myString
        myString = Dir
    Loop
    
    Set GetMessages = oColl

End Function
