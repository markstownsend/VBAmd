module VbaMd.ContinuedLine 

open System
open System.Text.RegularExpressions

/// <summary>
/// Creates the smallest number of continued lines from the list of lines potentially containing line continuations.
/// </summary>
let rec collapseLineContinuations (lines : string list)  = 

    /// combine a line continuation and drop the trailing continuation character sequence from the first argument
    let combineLine (first : string) (second : string) = 
        Regex.Replace(first, @" _$", " ") + second.Trim()

    let collapsing = 
        lines
            |> List.mapi(fun i f ->  (i, f))
            |> List.filter(fun f -> (snd f).EndsWith(" _"))

    match collapsing with
    | [] -> lines
    | _   -> try 
                (collapseLineContinuations (collapsing |> List.map( fun f -> combineLine (snd f) (lines.Item( (fst f) + 1))))) @ lines
             with   // once you reach outside the list for the rest of the continued line then any subsequent continuation is in the 
                    // next window so you have done all you can do in this window so return it.
                    // TODO:Is there a better way than catching the IndexOutOfBounds exception, afterall that takes some time to do
                | :? ArgumentException as ex -> printfn "Index out of bounds exception: %s" ex.Message ; lines
                | _ as ex -> printfn "%s" (ex.ToString()) ; lines
        
/// <summary>
/// Recognizes and removes duplicate line number 'tokens' (the result of a previous operation in the pipeline).
/// </summary>
let stripDuplicateTokens (rawLines : string seq) =

    /// breaks down the string with duplicate tokens and builds it back up without the duplicates
    let removeDuplicates (line : string) =
        // identify the possible duplicates
        let tokens = Regex.Split(line, @"(\d{3}\.\s)") |> Array.filter(fun f -> f.Trim().Length > 0)
        let counttokens = tokens.Length
        let possibleDuplicateTokens = seq {
            for i in 0 .. 2 .. counttokens - 1 do
                yield (tokens.[i], tokens.[i+1])
        }
        // reconstruct the string removing the duplicates
        let noDuplicateTokens = 
            possibleDuplicateTokens 
                |> Seq.distinctBy(fun f -> (fst f))
                |> Seq.fold(fun (a, b) (m, n) -> (a + m, b + n)) ("","") 
        tokens.[0] + (snd noDuplicateTokens)

    seq {
        for line in rawLines do
            let tokenized = Regex.Matches(line, "(\d{3}\.\s)")
            if tokenized.Count > 0 then
                match tokenized.Count with
                | 1 -> yield line
                | _ -> yield removeDuplicates line
            else
                yield line
    }