module VbaMd.MainTests 

// yaun - yet another unused nuget

open System.IO
open VbaMd.Main
open Xunit
open ApprovalTests
open ApprovalTests.Reporters
open System.Runtime.CompilerServices

[<Fact>]
let ``FileInfo().Directory returns the directory without the path separator`` () =
    let testvalue = @"C:\mark\excel\iati-xl2xml\src\xl2xml.xlsm"
    let expected = @"C:\mark\excel\iati-xl2xml\src"
    let actual = FileInfo(testvalue).Directory
    Assert.Equal(expected, actual.FullName)


/// test getValidLines in main - pass an array of lines, in other words a little vba module here in an array
let moduleAsArray_testdata = [|  @"'####Module : MXMLHelper";
                            @"'#####Type : Module";
                            @"'#####Description : Functions to work with aspects of XML";
                            @"'***";
                            @"Option Explicit";
                            @"Option Private Module";
                            @"";
                            @"Private Const MODULENAME As String = ""MXMLHelper""";
                            @"";
                            @"'***";
                            @"Public Function IsEmptyElement(ByRef p_objXElement As MSXML2.IXMLDOMElement) As Boolean";
                            @"'***";
                            @"'>Description : Tests if an element is named but otherwise empty";
                            @"'>Parameters :";
                            @"'> > 1. p_objXElement : the element to test";
                            @"'>";
                            @"'>Returns : true if the element is named and empty, false otherwise";
                            @"'>Dependencies: none";
                            @"'>Usage :";
                            @"'```";
                            @"'Dim oMyElement as new MSXML2.IXMLDOMElement";
                            @"'Debug.Assert(false = IsEmptyElement(oMyElement))";
                            @"'```";
                            @"    Const METHODNAME As String = ""IsEmptyElement""";
                            @"    IsEmptyElement = False";
                            @"    If (p_objXElement.Attributes.Length = 0 And Not (p_objXElement.HasChildNodes()) And p_objXElement.nodeTypedValue = vbNullString) Then IsEmptyElement = True";
                            @"    ";
                            @"End Function";
                            @"";
                            @"'***";
                            @"Public Function ConvertToSpecialCharacter(ByRef p_strIn As String) As String";
                            @"'***";
                            @"'>Description : converts the characters that need xml escaping";
                            @"'>Parameters :";
                            @"'> > 1. p_strIn : the raw string to perform the replacements";
                            @"'>";
                            @"'>Returns : the xml escaped string";
                            @"'>Dependencies: none";
                            @"'>Notes : makes the following conversions: & to &amp;, < to &lt;, > to &gt;, double quote to &quot;, single quote to &apos;,";
                            @"'>Usage :";
                            @"'```";
                            @"'Dim sRawString as String : s = ""<hello>""";
                            @"'Debug.Assert(s  = ""&lt;hello&gt;"")";
                            @"'```";
                            @"    Const METHODNAME As String = ""ConvertToSpecialCharacter""";
                            @"    ConvertToSpecialCharacter = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(p_strIn, Chr(34), ""&quot;""), ""&"", ""&amp;""), ""'"", ""&apos;""), ""<"", ""&lt;""), "">"", ""&gt;""), vbCrLf, """"), vbCr, """"), vbLf, """"), vbBack, """"), vbFormFeed, """"), vbNewLine, """"), vbNullChar, """"), vbNullString, """"), vbTab, """"), vbVerticalTab, """"))";
                            @"";
                            @"End Function"
                        |]

[<Fact>]
let ``getValidLines returns a non-empty array`` = 
    let actual = getValidLines moduleAsArray_testdata
    Assert.NotEmpty(actual);

[<Fact>]
let ``getValidLines with test data returns 37 lines`` = 
    let actual = getValidLines moduleAsArray_testdata
    Assert.True(37 = actual.Length)

[<Theory>]
[<InlineDataAttribute(@"####Module : MXMLHelper  ")>]
[<InlineDataAttribute(@"#####Type : Module  ")>]
[<InlineDataAttribute(@"#####Description : Functions to work with aspects of XML  ")>]
[<InlineDataAttribute(@"***  ")>]
[<InlineDataAttribute(@"Option Explicit  ")>]
[<InlineDataAttribute(@"Option Private Module  ")>]
[<InlineDataAttribute(@"***  ")>]
[<InlineDataAttribute(@"Public Function IsEmptyElement(ByRef p_objXElement As MSXML2.IXMLDOMElement) As Boolean  ")>]
[<InlineDataAttribute(@"***  ")>]
[<InlineDataAttribute(@">Description : Tests if an element is named but otherwise empty  ")>]
[<InlineDataAttribute(@">Parameters :  ")>]
[<InlineDataAttribute(@"> > 1. p_objXElement : the element to test  ")>]
[<InlineDataAttribute(@">  ")>]
[<InlineDataAttribute(@">Returns : true if the element is named and empty, false otherwise  ")>]
[<InlineDataAttribute(@">Dependencies: none  ")>]
[<InlineDataAttribute(@">Usage :  ")>]
[<InlineDataAttribute(@"```  ")>]
[<InlineDataAttribute(@"Dim oMyElement as new MSXML2.IXMLDOMElement  ")>]
[<InlineDataAttribute(@"Debug.Assert(false = IsEmptyElement(oMyElement))  ")>]
[<InlineDataAttribute(@"```  ")>]
[<InlineDataAttribute(@"End Function  ")>]
[<InlineDataAttribute(@"***  ")>]
[<InlineDataAttribute(@"Public Function ConvertToSpecialCharacter(ByRef p_strIn As String) As String  ")>]
[<InlineDataAttribute(@"***  ")>]
[<InlineDataAttribute(@">Description : converts the characters that need xml escaping  ")>]
[<InlineDataAttribute(@">Parameters :  ")>]
[<InlineDataAttribute(@"> > 1. p_strIn : the raw string to perform the replacements  ")>]
[<InlineDataAttribute(@">  ")>]
[<InlineDataAttribute(@">Returns : the xml escaped string  ")>]
[<InlineDataAttribute(@">Dependencies: none  ")>]
[<InlineDataAttribute(@">Notes : makes the following conversions: & to &amp;, < to &lt;, > to &gt;, double quote to &quot;, single quote to &apos;,  ")>]
[<InlineDataAttribute(@">Usage :  ")>]
[<InlineDataAttribute(@"```  ")>]
[<InlineDataAttribute("Dim sRawString as String : s = \"<hello>\"  ")>]
[<InlineDataAttribute("Debug.Assert(s  = \"&lt;hello&gt;\")  ")>]
[<InlineDataAttribute("```  ")>]
[<InlineDataAttribute("End Function  ")>]
let ``getValidLines contain all these rows`` value = 
    let actual = getValidLines moduleAsArray_testdata
    Assert.Contains(actual, fun a -> a = value.ToString())

/// do an approval test on a markdown file produced by the entire process
[<Literal>]
let testMdFile = __SOURCE_DIRECTORY__ + @"\testdata\MExcelHelper.md"
[<Literal>]
let testBasFile = __SOURCE_DIRECTORY__ + @"\testdata\MExcelHelper.bas"
[<Literal>]
let parent = "xl2xml.xlsm"

////********************************************************************//
////Can't get this to run with FAKE.  Will run through the xunit.console.runner.exe
////in the solution tools directory so commented out here for the build
//[<Fact>]
//[<UseReporterAttribute(typedefof<Reporters.TortoiseTextDiffReporter>)>]
//[<MethodImpl(MethodImplOptions.NoInlining)>]
//let ``BasFileGeneration`` () =
//    // do
//    parseFile testBasFile parent
//    // verify
//    Approvals.VerifyFile(testMdFile)
////********************************************************************//