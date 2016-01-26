/// <summary>
/// Provides the functionality to identify a line the parser considers valid.
/// </summary>
/// <remarks>
/// The only valid lines are: comments, method opening statements, method closing, option statements, property statements, module level constant and variable declarations
/// TODO: have to trap the use of embedded underscores like Ba_nana and _Banana and Banana_ can't be preceeded by a space
/// </remarks>
[<AutoOpen>]
module public VbaMd.ValidLine  

/// <summary>
/// Helper function to pick valid comments and valid code lines
/// </summary>
val isValidLine : line:string -> bool 
