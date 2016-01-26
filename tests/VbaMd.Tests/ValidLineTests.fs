module Vbamd.ValidLineTests

open VbaMd.ValidLine
open Xunit

[<Theory>]
[<InlineDataAttribute("001. Public Function AFunc() As String")>]
[<InlineDataAttribute("002. Option Explicit")>]
[<InlineDataAttribute("003. Private Function BFunc(ByRef p_strMessage As String) As Integer")>]
[<InlineDataAttribute("004. Public Function CFunc(ByRef p_strMessage As String, ByVal l_lngLevel as Long) As Variant")>]
[<InlineDataAttribute("005. Function DFunc() As String")>]
[<InlineDataAttribute("006. Function EFunc(ByRef p_strMessage As String) As Integer")>]
[<InlineDataAttribute("007. Function FFunc(ByRef p_strMessage As String, ByVal l_lngLevel as Long) As Integer")>]
[<InlineDataAttribute("008. End Function")>]
[<InlineDataAttribute("009. End Sub")>]
[<InlineDataAttribute("010. Public Sub ASub() As String")>]
[<InlineDataAttribute("011. Private Sub BSub(ByRef p_strMessage As String)")>]
[<InlineDataAttribute("012. Public Sub CSub(ByRef p_strMessage As String, ByVal l_lngLevel as Long)")>]
[<InlineDataAttribute("013. Sub DSub() As String")>]
[<InlineDataAttribute("014. Sub ESub(ByRef p_strMessage As String)")>]
[<InlineDataAttribute("015. Sub FSub(ByRef p_strMessage As String, ByVal l_lngLevel as Long)")>]
[<InlineDataAttribute("016. Public Function IsEmptyElement(ByRef p_objXElement As MSXML2.IXMLDOMElement) As Boolean")>]
let ``good code lines that should return true`` value = 
    let actual = isValidLine value
    Assert.True(actual)
    
[<Theory>]
[<InlineDataAttribute("001. Public Function GFuncEra( _")>]
[<InlineDataAttribute("002. Private Function HFunc(ByRef p_strMessage As String,")>]
[<InlineDataAttribute("003. Public")>]
[<InlineDataAttribute("004. Dim a As String")>]
[<InlineDataAttribute("005. x = x + y")>]
[<InlineDataAttribute("006. Public Sub GSub( _")>]
[<InlineDataAttribute("007. Private Sub HSub(ByRef p_strMessage As String,")>]
[<InlineDataAttribute("008. Private")>]
let ``bad code lines that should return false`` value = 
    let actual = isValidLine value
    Assert.False(actual)
     