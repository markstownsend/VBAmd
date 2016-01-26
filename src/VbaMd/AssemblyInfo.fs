namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("VbaMd")>]
[<assembly: AssemblyProductAttribute("VbaMd")>]
[<assembly: AssemblyDescriptionAttribute("Extracts markdown from your VBA source files.")>]
[<assembly: AssemblyVersionAttribute("1.0")>]
[<assembly: AssemblyFileVersionAttribute("1.0")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "1.0"
