Attribute VB_Name = "EnumModule"
' FORREST SOFTWARE
' Copyright (c) 2017 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Public Enum E_JAKI_WZORZEC_ZOSTAL_SPELNIONY

    
    E_OTHER = -1

    E_NO_MATCH = 0
    
    E_03_PATTERN = 3
    E_07_PATTERN = 7
    E_11_PATTERN = 11
    
    
    
    E_TSK_3_PRE = 101
    E_TSK_3_MID = 102
    E_TSK_3_END = 103
    E_TSK_2_PRE = 104
    E_TSK_2_END = 105
    
    E_TSK_ALONE = 106
    
    
    E_CGH_3_PRE = 1001
    E_CGH_3_MID = 1002
    E_CGH_3_END = 1003
    E_CGH_2_PRE = 1004
    E_CGH_2_END = 1005
    
    E_CGH_ALONE = 1006
    
End Enum

Public Enum E_JAKI_OPERTOR
    E_EQUAL = 1
    E_NOT_EQUAL
    E_LIKE
    E_NOT_LIKE
    E_ZA_OPERATOR
    E_FLAG_OPERATOR
    E_CGH_LOGIC
    E_RECV_OPERATOR
    E_FLAG_OPERATOR_OSEA
    E_FLAG_OPERATOR_OSEA_F
End Enum


Public Enum E_INPUT_LIST
    INPUT_PLT = 1
    INPUT_PN
    INPUT_DUNS
    INPUT_A
    INPUT_CONT_NUM
    INPUT_DATE_ADDED
    INPUT_EFF
    INPUT_EXP
    INPUT_DESC
    INPUT_SUPPLIER_NM
    INPUT_BYR_CODE
    INPUT_OVERCOMING
End Enum

