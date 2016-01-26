/// <summary>
/// Provides the functionality to collapse continued lines.  Continued lines are specified in VBA with the character combination ' _'.
/// </summary>
[<AutoOpen>]
module public VbaMd.ContinuedLine 

/// <summary>
/// Creates the smallest number of continued lines from the list of lines potentially containing line continuations.
/// </summary>
/// <remarks>
/// The line continuation character sequence in vba is ' _'.  This function collapses the lines that are continued 
/// if there are no continuations or if it errors it just returns the list that was passed, if there are
/// continuations it collapses the continuations in to a single line removing the continuation characters.  If there are
/// multiple continuations in the window the single line will contain duplication that will be removed in a subsequent step.
/// If the continuations span multiple windows they will be collapsed successively and a subsequent step will qualify whether 
/// the collapsing operation produced a valid code line and that line will be picked if it did.  If a line is continued more than 
/// 4 times it will result in the code line being dropped from the documentation.  4 was chosen arbitrarily but it 
/// impairs readability to have multiple continuations of the same line so fix that in your code! :)
/// Finally, and more of a note to self, this implementation uses an exception handler to trap the attempt to reach below the list
/// for the (conceptual) next line.  In other words when the last line in the list has a continuation.  The exception is thrown away.
/// This is a costly mechanism for responding to this circumstance which could happen quite alot given the window (list of lines) 
/// is arbitrarily constructed.  I'd like to think of a better way to do this.  At some point.
/// </remarks>
/// <param name="lines">The list of continued lines.</param>
/// <returns>A smaller list of source lines with all line continuations removed.</returns>
val collapseLineContinuations : lines:string list -> string list
            

/// <summary>
/// Recognizes and removes duplicate line number 'tokens' (the result of a previous operation in the pipeline).
/// </summary>
/// <remarks>The token is the line number sequence: 'nnn. '.  A duplicate is considered the second instance of the token pattern
/// in the line.  In other words: '001. ' is the first token in a hypothetical line then the subsequent tokens
/// '002. ' and '003. ' are both duplicate tokens.  The actual number that the token represents is not parsed.</remarks>
/// <param name="rawLines"> The sequence of lines possibly containing duplicate tokens.</param>
val stripDuplicateTokens : rawLines:seq<string> -> seq<string> 