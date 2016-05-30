namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("VbaMd")>]
[<assembly: AssemblyProductAttribute("VbaMd")>]
[<assembly: AssemblyDescriptionAttribute("Extracts markdown from your VBA source files.")>]
[<assembly: AssemblyVersionAttribute("0.0.1")>]
[<assembly: AssemblyFileVersionAttribute("0.0.1")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "0.0.1"
    let [<Literal>] InformationalVersion = "0.0.1"
