module VbaMd.ContinuedLineTests 

open System
open System.IO
open VbaMd.ContinuedLine
open Xunit
        
// this is a walkthrough of the processing for this dataset
let testdata_LineContinuationInTheMiddle = ["001. Public Sub MyFunc(ByVal _";
                                            "002. argName1 as String, ByVal argName2 as String, _";
                                            "003. ByVal argName3 as String)";
                                            "004. Dim s as String"]
// add an index and tuple it then filter by endswith a space and underscore
let test_ContinuationInMiddle_Expected1 = [(0,"001. Public Sub MyFunc(ByVal _");
                                           (1,"002. argName1 as String, ByVal argName2 as String, _")]
// create a new entries with the tupled text with space and underscore removed then the next text from the original array trimmed and the original list appended
let test_ContinuationInMiddle_Expected2 = ["001. Public Sub MyFunc(ByVal 002. argName1 as String, ByVal argName2 as String, _";
                                           "002. argName1 as String, ByVal argName2 as String, 003. ByVal argName3 as String)"]
// add an index and tuple it then filter by endswith a space and underscore
let test_ContinuationInMiddle_Expected3 = [(0,"001. Public Sub MyFunc(ByVal 002. argName1 as String, ByVal argName2 as String, _")]
// create a new entries with the tupled text with space and underscore removed then the next text from the original array trimmed and the original list appended
let test_ContinuationInMiddle_Expected4 = ["001. Public Sub MyFunc(ByVal 002. argName1 as String, ByVal argName2 as String, 002. argName1 as String, ByVal argName2 as String, 003. ByVal argName3 as String)"]
// add an index and tuple it then filter by endswith a space and underscore
let test_ContinuationInMiddle_Expected5 = []
// append original list to the function result, unwinding the recursion stack
let test_ContinuationInMiddle_Expected6 = ["001. Public Sub MyFunc(ByVal 002. argName1 as String, ByVal argName2 as String, 002. argName1 as String, ByVal argName2 as String, 003. ByVal argName3 as String)";
                                           "001. Public Sub MyFunc(ByVal 002. argName1 as String, ByVal argName2 as String, _";
                                           "002. argName1 as String, ByVal argName2 as String, 003. ByVal argName3 as String)"]
// append original list to the function result, unwinding the recursion stack
let test_ContinuationInMiddle_Expected7 = ["001. Public Sub MyFunc(ByVal 002. argName1 as String, ByVal argName2 as String, 002. argName1 as String, ByVal argName2 as String, 003. ByVal argName3 as String)";
                                           "001. Public Sub MyFunc(ByVal 002. argName1 as String, ByVal argName2 as String, _";
                                           "002. argName1 as String, ByVal argName2 as String, 003. ByVal argName3 as String)";
                                           "001. Public Sub MyFunc(ByVal _";
                                           "002. argName1 as String, ByVal argName2 as String, _";
                                           "003. ByVal argName3 as String)";
                                           "004. Dim s as String"]


[<Theory>]
[<InlineDataAttribute("001. Public Sub MyFunc(ByVal 002. argName1 as String, ByVal argName2 as String, 002. argName1 as String, ByVal argName2 as String, 003. ByVal argName3 as String)")>]
[<InlineDataAttribute("001. Public Sub MyFunc(ByVal 002. argName1 as String, ByVal argName2 as String, _")>]
[<InlineDataAttribute("002. argName1 as String, ByVal argName2 as String, 003. ByVal argName3 as String)")>]
[<InlineDataAttribute("001. Public Sub MyFunc(ByVal _")>]
[<InlineDataAttribute("002. argName1 as String, ByVal argName2 as String, _")>]
[<InlineDataAttribute("003. ByVal argName3 as String)")>]
[<InlineDataAttribute("004. Dim s as String")>]
let ``ContinuationInMiddle with 4 lines returns these specific lines`` value =
    let actual = collapseLineContinuations testdata_LineContinuationInTheMiddle
    Assert.Contains(actual, fun a -> a = value.ToString())

[<Fact>]
let ``ContinuationInMiddle with 4 lines returns 7 Lines`` () =
    let actual = collapseLineContinuations testdata_LineContinuationInTheMiddle
    Assert.True(7 = actual.Length)



//// I would like to somehow get this to work
//// because I think it will allow me to confirm the order of elements in the list
//// is what I expect it to be.  I think the tests above don't confirm order.  I'll have to come back to it though
//let ``ContinuationInMiddle Use of contains`` () = 
//    let actual = collapseLineContinuations test_ContinuationInMiddle_CollapsesToSingleLineWithDuplicates
////    let contains = [| <@ new Action<_>(fun (f:string) -> ignore(f.Length = 74)) @>;
////                     <@ new Action<_>(fun (f:string) -> ignore(f.Length = 28)) @>;
////                     <@ new Action<_>(fun (f:string) -> ignore(f.Length = 28)) @>;
////                     <@ new Action<_>(fun (f:string) -> ignore(f.Length = 28)) @>;
////                     <@ new Action<_>(fun (f:string) -> ignore(f.Length = 28)) @>;
////                     <@ new Action<_>(fun (f:string) -> ignore(f.Length = 28)) @>;
////                     <@ new Action<_>(fun (f:string) -> ignore(f.Length = 28)) @> |]
//    let contains = [|
//                    fun (f:string) -> ignore(f.Length = 74);
//                    fun (f:string) -> ignore(f.Length = 28);
//                    fun (f:string) -> ignore(f.Length = 28);
//                    fun (f:string) -> ignore(f.Length = 28);
//                    fun (f:string) -> ignore(f.Length = 28);
//                    fun (f:string) -> ignore(f.Length = 28);
//                    fun (f:string) -> ignore(f.Length = 28) 
//                    |]
//    Assert.Collection(actual, contains)

/// collapseLineContinuations test strings
let testdata_LineContinuationAtEnd = ["010. Public Sub MyFunc(ByVal _"; "011. argName1 as String, ByVal argName2 as String, _"; "012. ByVal argName3 as String, _";"013. ByVal argName4 as String, _"]

[<Fact>]
let ``ContinuationAtEnd ThrowsAndCatchesOutOfRangeException NoCollapse`` () =
    let actual = collapseLineContinuations testdata_LineContinuationAtEnd
    Assert.True(4 = actual.Length)

[<Theory>]
[<InlineDataAttribute("010. Public Sub MyFunc(ByVal _")>]
[<InlineDataAttribute("011. argName1 as String, ByVal argName2 as String, _")>]
[<InlineDataAttribute("012. ByVal argName3 as String, _")>]
[<InlineDataAttribute("013. ByVal argName4 as String, _")>]
let ``ContinuationAtEnd ThrowsAndCatchesOutOfRangeException OriginalListRemains`` value =
    let actual = collapseLineContinuations testdata_LineContinuationAtEnd
    Assert.Contains(actual, fun a -> a = value.ToString())


/// tests for stripDuplicateTokens
let test_ContinuationInMiddle_Expected8 = ["011. Public Sub MyFunc2(ByVal 012. argName1 as String, 012. argName1 as String, 013. ByVal argName2 as String)";
                                           "011. Public Sub MyFunc2(ByVal 012. argName1 as String, _";
                                           "012. argName1 as String, 013. ByVal argName2 as String)";
                                           "010. End Function"
                                           "011. Public Sub MyFunc2(ByVal _";
                                           "012. argName1 as String, _";
                                           "013. ByVal argName2 as String)"]

[<Theory>]
[<InlineDataAttribute("011. Public Sub MyFunc2(ByVal argName1 as String, ByVal argName2 as String)")>]
[<InlineDataAttribute("011. Public Sub MyFunc2(ByVal argName1 as String, _")>]
[<InlineDataAttribute("012. argName1 as String, ByVal argName2 as String)")>]
[<InlineDataAttribute("010. End Function")>]
[<InlineDataAttribute("011. Public Sub MyFunc2(ByVal _")>]
[<InlineDataAttribute("012. argName1 as String, _")>]
[<InlineDataAttribute("013. ByVal argName2 as String)")>]
let ``stripDuplicateTokens removes all but first token`` value =    
    let actual  = stripDuplicateTokens test_ContinuationInMiddle_Expected8
    Assert.Contains(actual, fun a -> a = value.ToString())