module VbaMd.ValidLine 

open System
open System.Text.RegularExpressions
        
[<Literal>]
let ValidCommentLinePattern = @"^\d{3}\.\s*'[`\*#\>\-\s\w]+"
[<Literal>]
let ValidMethodOpeningPattern = @"^\d{3}\.\s*(Public|Private|)\s?(Sub|Function)\s{1}([\w]+)(\([\w\s\.,]*\))(?!.*(_|Const|End|For|Function|Global|If|Public|Private|Sub|With))(\sAs\s+\w+[\.\w]*)?"
[<Literal>]
let ValidMethodClosingPattern = @"^\d{3}\.\s*(End (Sub|Function)){1}"
[<Literal>]
let ValidOptionPattern = @"^\d{3}\.\s*Option{1}.+"
//[<Literal>]
//let SimpleRegex = @""

/// helper function to pick valid comments and valid code lines
let isValidLine (line : string) = 

    /// has to be a complete match of the entire line, not just the enclosing capturing classes
    /// for the code line to be considered a match
    let (|IsRequiredMethodLine|_|) (pattern: string) (input : string) = 
        let result = Regex.Match(input, pattern)
        if result.Success then
            if result.Value.Length = input.Length then   // no partial matches
                Some "match"
            else
                None
        else
            None

    /// comments essentially match the first tick and have no capturing classes
    let (|IsRequiredCommentLine|_|) (pattern: string) (input : string) = 
        let result = Regex.Match(input, pattern)
        if result.Success then
            Some "match"
        else
            None

    /// declarations (options, consts, module level) essentially match the option keyword and have no capturing classes
    let (|IsRequiredDeclarationLine|_|) (pattern: string) (input : string) = 
        let result = Regex.Match(input, pattern)
        if result.Success then
            Some "match"
        else
            None

     
    match line with
        | IsRequiredMethodLine ValidMethodOpeningPattern "match"
            -> true
        | IsRequiredMethodLine ValidMethodClosingPattern "match"
            -> true 
        | IsRequiredCommentLine ValidCommentLinePattern "match"
            -> true 
        | IsRequiredDeclarationLine ValidOptionPattern "match"
            -> true
        | _ -> false
