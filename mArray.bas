Attribute VB_Name = "mArray"
Option Explicit

'VB stores arrays and tables as 'safe arrays' and therefore we can access the descriptor
'thru a simple routine to extract all the information about them:

'whether it is (re-)dimensioned or undimensioned
'number of dimensions
'size of each table element
'characteristics and features
'lBound, uBound and number of elements in each dimension

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MaxDims As Long = 8  'max number of dimensions; modify if you need more
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Declare Function hTable Lib "msvbvm50.dll" Alias "VarPtr" (Table() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Enum ArrayFeatures
    FADF_AUTO = &H1             'Array is allocated on the stack
    FADF_STATIC = &H2           'Array is statically allocated
    FADF_EMBEDDED = &H4         'Array is embedded in a structure
    FADF_FIXEDSIZE = &H10       'Array may not be resized or reallocated
    FADF_BSTR = &H100           'An array of BSTRs
    FADF_UNKNOWN = &H200        'An array of IUnknown*
    FADF_DISPATCH = &H400       'An array of IDispatch*
    FADF_VARIANT = &H800        'An array of VARIANTs
    FADF_RESERVED = &HFFFFF0E8  'Bits reserved for future use
End Enum
#If False Then 'keep capitalization
Private FADF_AUTO, FADF_STATIC, FADF_EMBEDDED, FADF_FIXEDSIZE, FADF_BSTR, FADF_UNKNOWN, FADF_DISPATCH, FADF_VARIANT, FADF_RESERVED
#End If

Private Type SAFEARRAYBOUND
    NumElements As Long
    LBound      As Long
    UBound      As Long
End Type

Public Type SAFEARRAYDESCRIPTOR
    NumDims     As Integer      'number of dimensions
    Features    As Integer      'feature bits
    ElementSize As Long         'size of one element
    Locks       As Long         'number of locks
    PtrToData   As Long         'pointer to first element
    Bounds(1 To MaxDims) _
           As SAFEARRAYBOUND    'number of elements and lbound/ubound for each dimension
End Type

Public Function GetArrayDescriptor(ByVal hTable As Long) As SAFEARRAYDESCRIPTOR

  'param hTable must point to a pointer which in turn points to the array descriptor
  '
  'you get the hTable parameter by calling hTable(your_table_name)
  'hTable is in fact a disguise for the VarPtr function (which unfortunately does not
  'accept tables() )
  '
  'so the function call for this function should look like this:
  '
  'GetArrayDescriptor(hTable(your_table_name)) 'returns SAFEARRAYDESCRIPTOR for your_table_name
  '
  'one little drawback though:
  'apparently VB does not store variable (redimmable) tables of variable length strings
  'as safearrays, so this set of routines does not work with this kind of tables.
  '
  'it's okay however with fixed (non-redimmable) tables of variable length strings
  'and with variable (redimmable) tables of fixed length strings

  Dim PtrToDesc As Long
  Dim i         As Long

    CopyMemory PtrToDesc, ByVal hTable, 4
    If PtrToDesc Then
        With GetArrayDescriptor
            CopyMemory .NumDims, ByVal PtrToDesc, 16 'get the first 16 bytes (NumDims..PtrToData)
            If .NumDims <= MaxDims Then 'to prevent out of range indexing
                PtrToDesc = PtrToDesc + 16 'adjust pointer
                For i = .NumDims To 1 Step -1 'in reverse order; the m.s. dimension is rightmost
                    CopyMemory .Bounds(i), ByVal PtrToDesc, 8 'get Number of Elements and LBound
                    PtrToDesc = PtrToDesc + 8 'adjust pointer
                    With .Bounds(i)
                        .UBound = .LBound + .NumElements - 1 'calculate UBound
                    End With '.BOUNDS(I)
                Next i
            End If
        End With 'GETARRAYDESCRIPTOR
    End If

End Function

Public Function IsDimmed(ByVal hTable As Long) As Boolean

  'hTable points to a pointer which in turn points to the array desriptor

  'you get the hTable parameter by calling hTable(your_table_name) so the
  'function call should look like this:

  'If IsDimmed(hTable(your_table_name)) Then ...

    IsDimmed = GetArrayDescriptor(hTable).NumDims

End Function

':) Ulli's VB Code Formatter V2.17.3 (2004-Jul-23 11:12) 48 + 56 = 104 Lines
