Attribute VB_Name = "m_CONST"
Option Explicit

Public Const MARKER_CONTENT = "$Content$"

Public Const CELLNAME_VHTML_TYPE = "User.vhtml_Type"
Public Const CELLNAME_VHTML_LINK = "User.vhtml_Link"
Public Const CELLNAME_HTML_PATTERN = "User.TextPattern"
Public Const CELLNAME_GLOBAL_TEMPLATE = "User.vhtml_GlobalTemplate"
Public Const CELLNAME_CONTENT = "Prop.Content"
Public Const CELLNAME_ATTRS = "Prop.Attrs"
Public Const CELLNAME_MARKER = "Prop.Marker"

Public Const EDIT_CONTENT = 0
Public Const EDIT_ATTRS = 1

Public Const EMPTY_TEMPLATE_PATTERN = "$GENERAL_CONTENT$"
Public Const EMPTY_TEMPLATE_TITLE = "$GENERAL_TITLE$"

Public Enum vhtml_types
    vt_Template = 0
    vt_Block = 1
    vt_Single = 2
    vt_String = 3
    vt_Frame = 4
End Enum
