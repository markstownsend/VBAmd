(*** hide ***)
#I "../../bin"


(**
Documenting VBA with Markdown syntax
====================================

This script extracts method comments from your VBA source files and turns them into markdown files.  Those markdown files can be subsequently processed with [FSharp.Formatting](http://tpetricek.github.io/FSharp.Formatting) or rendered as is.
The library does not extract your VBA source files from the spreadsheet, there are other tools that do that, [CodeCleaner](http://www.appspro.com/Utilities/CodeCleaner.htm) being one and [RubberDuck](http://www.rubberduck-vba.com) another or 
if you'd like to look at a complete F# solution using FAKE then look at the build script on this [project](https://github.com/wateraid/iati-xl2xml).

This script relies on the fact your code files have not been edited outside the VBA IDE.  If you edit your code files outside the VBA IDE (for example in a text editor) you could introduce syntax errors and these
will likely defeat the parsing in this script, especially around method opening statements, method closing statements and comment single ticks.

*)


(**
Get the comments
----------------
The crux of this library is simply the single tick that indicates the line is a comment in VBA.  The library will get any comments that are bound by the start of the file and the start of a method
or any comments that are bound by the start of a method and the end of a method (function or sub) declaration.  For example:

    [lang=vba]
    1. -start of file-
    .. optional empty lines
    ..
    n. ' some comment
    n1 [Public|Private|] Function|Sub methodName [()] [As SomeType] [_] 
    n2     any continuation if the preceeding line had a continuation character

also

    [lang=vba]
    1. [Public|Private|] Function|Sub methodName [()] [As SomeType] [_]
    2.     any continuation if the preceeding line had a continuation character
    .. optional empty lines
    ..
    n. ' some comment
    n1.' some other comment
    n2. End [Function|Sub]
    

This method will not get inline comments, only the header and method blocks.  This is by design, the inline
comments tend to have a very narrow context for example referring only to the current line or next line of code and that doesn't need to be surfaced up to the public documentation. 

*)

#r @"..\..\src\VbaMd\bin\Debug\VbaMd.dll"
open VbaMd.Main
parseFile @"somevbafile.bas" "someworkbook.xlsx"

(**
Some more info
*)
