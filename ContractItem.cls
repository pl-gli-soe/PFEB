VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContractItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public is_active_in_another_words_have_p_on_a_field As Boolean
Public is_valid As Boolean
Public is_future As Boolean
Public is_past As Boolean
Public has_friends As Boolean
Public expires_this_year As Boolean
Public has_future_successor As Boolean
Public is_successor As Boolean


Public eff As Date
Public exp As Date
Public buyer_code As String
Public status_description As String

Public plt As String
Public pn As String
Public duns As String


Public referencja As Range
