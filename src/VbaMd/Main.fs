module VbaMd.Main

open System
open System.IO
open System.Collections.Generic
open System.Text.RegularExpressions

//let parseSourceDirectory (inFile : string) (outDir : string) = 
//
//    let srcDir = DirectoryArg(inFile)
//     
//    let allLines = 
//        try
//            File.ReadAllLines(inFile) 
//        with
//            e -> printfn "An exception was thrown reading the file: %s" e.Message ; [| |]
//    let relevantLines = getValidLines allLines
//    File.WriteAllLines(outDir + "test.md", relevantLines)
//    ()
//            
//    /// get all the valid lines
//    let getValidLines (lines : string[]) = 
//
//        /// helper function to combine a jagged array in to a sequence
//        let collapse (arr : seq<string list>) = 
//            seq {
//                for ar in arr do
//                    for i in 0..List.length(ar) - 1 do
//                        yield ar.[i]
//            }
//
//        let numbered  = lines |> Array.mapi(fun i f -> (i+1).ToString("d3")  + ". " + f) |> Array.toSeq
//        let windows = 
//            numbered 
//            |> Seq.windowed(5)  // make the window lists
//            |> Seq.map (fun f -> ContinuedLine.collapseLineContinuations (f |> Array.toList))  // collapse the continuations
//        
//        let allLines = collapse windows    // put them all back together again
//        let uniqueLines = allLines |> Seq.distinct // remove all duplicate lines
//        let uniqueTokensOnly = ContinuedLine.stripDuplicateTokens uniqueLines    // clean the lines containing duplicate text as a result of the line continuation collapse
//        uniqueTokensOnly 
//            |> Seq.toArray
//            |> Array.map(fun f -> Regex.Replace(f, @"\d{3}\.\s(?<!^\d{3}\.\s)", ""))        // get rid of embedded line numbers
//            |> Array.filter(fun f -> ValidLine.isValidLine f)  // get rid of invalid lines
//            |> Array.map(fun f -> f.Replace("'", "")) // get rid of comment symbols
//            |> Array.map(fun f -> Regex.Replace(f, @"\d{3}\.\s",""))        // strip the leading numbers 
//            |> Array.map(fun f -> f.PadRight(f.Length + 2, ' '))    // add two spaces to the end for the markdown new line 

/// <summary>
/// Helper function to get the valid lines from the source lines.
/// </summary>
let getValidLines (lines : string[]) = 

    /// helper function to combine a jagged array in to a sequence
    let collapse (arr : seq<string list>) = 
        seq {
            for ar in arr do
                for i in 0..List.length(ar) - 1 do
                    yield ar.[i]
        }

    let numbered  = lines |> Array.mapi(fun i f -> (i+1).ToString("d3")  + ". " + f) |> Array.toSeq
    let windows = 
        numbered 
        |> Seq.windowed(5)  // make the window lists
        |> Seq.map (fun f -> ContinuedLine.collapseLineContinuations (f |> Array.toList))  // collapse the continuations
        
    let allLines = collapse windows    // put them all back together again
    let uniqueLines = allLines |> Seq.distinct // remove all duplicate lines
    let uniqueTokensOnly = ContinuedLine.stripDuplicateTokens uniqueLines    // clean the lines containing duplicate text as a result of the line continuation collapse
    uniqueTokensOnly 
        |> Seq.toArray
        |> Array.map(fun f -> Regex.Replace(f, @"\d{3}\.\s(?<!^\d{3}\.\s)", ""))        // get rid of embedded line numbers
        |> Array.filter(fun f -> ValidLine.isValidLine f)  // get rid of invalid lines
        |> Array.map(fun f -> f.Replace("'", "")) // get rid of comment symbols
        |> Array.map(fun f -> Regex.Replace(f, @"\d{3}\.\s",""))        // strip the leading numbers 
        |> Array.map(fun f -> f.PadRight(f.Length + 2, ' '))    // add two spaces to the end for the markdown new line 

/// <summary>
/// Process the source file.
/// </summary>
[<CompiledName("ParseFile")>]
let parseFile (sourcefile : string) (parent : string) = 
    //TODO: what is the accepted return from this, can I return a tuple, what are the library guidelines?
    let source = FileInfo(sourcefile)
    let moduleName = source.Name.Substring(0, source.Name.Length - 4)   // get rid of the extension
    let outDir = source.Directory

    let allLines = 
        try
            File.ReadAllLines(sourcefile) 
        with
            e -> printfn "An exception was thrown reading the file: %s" e.Message ; [| |]
    //TODO: pass the parent in to this method so it appears somewhere in the top of the markdown file
    let relevantLines = getValidLines allLines
    try
        File.WriteAllLines(outDir.FullName + @"\" + moduleName + @".md", relevantLines)
    with 
        e -> printfn "An exception was thrown writing the file: %s" e.Message ; ()