/// <summary>
/// API module for the library.
/// </summary>
module public VbaMd.Main

/// <summary>
/// Process the source files in the entire directory.
/// </summary>
/// <remarks>
/// Iterate over the files in the given directory searching for the *.bas, *.cls and *.frm files.
/// All source files and the single spreadsheet ('container' or 'parent') application are expected to be in the same directory. 
/// If there is more than one spreadsheet application (*.xls, *.xlsm, *.xla) it will assume the first one is the parent (this may not be what you want).
/// If the spreadsheet application is missing it will consider the source files to be standalone and pick the directory name as the parent.
/// So if you have some utilities that you save and import in to new workbooks they can be documented too.
/// Files are saved with the same filename except with the *.md extension.
/// </remarks>
//[<CompiledName("ParseSourceDirectory")>]
//val parseSourceDirectory : inDir:string -> outDir:string -> unit


/// <summary>
/// Process the source file.
/// </summary>
/// <remarks>
/// Process the single source file and save the markdown to the source directory
/// </remarks>
[<CompiledName("ParseFile")>]
val parseFile : sourceFile:string -> parent:string -> unit


/// <summary>
/// Helper function to get the valid lines from the source lines.
/// </summary>
val getValidLines : lines:string[] -> string[]