namespace VbaMd.Args

open System
open System.IO


/// <summary>
/// Represents one of the arguments, as a directory on the current system.
/// </summary>
/// <remarks>
/// This needs to be an existing directory on the local system to which the application has write permissions.
/// </remarks>
[<Class>]
type internal DirectoryArg(directory) = 

    [<Literal>]
    let PossibleHosts = @"*.xlsm, *.xls, *.xla"
    [<Literal>]
    let PossibleSources = @"*.bas, *.cls, *.frm"
    
    let _fullPath = directory
    let mutable _isValid = false
    let mutable _parentDirectory = String.Empty
    let mutable _hostApp = String.Empty

    do match Directory.Exists(_fullPath) with
            | false ->  printfn "You have not entered a valid directory on the local machine.  The script will terminate now." |> ignore
                        _isValid <- false
            | true ->   _isValid <- true

    do match _isValid with
        | true -> _parentDirectory <- Directory.GetParent(_fullPath).Name
        | false -> _parentDirectory <- directory
    
    do  try
            _hostApp <- Directory.GetFiles(directory, PossibleHosts) |> Array.toList |> List.head
        with
            e -> printfn "Could not find a host application.  Taking the directory path as the host instead." ;  _hostApp <- _parentDirectory

    member internal this.IsValid with get () = _isValid
    member internal this.FullPath with get () = _fullPath
    member internal this.Parent with get () = _parentDirectory
    member internal this.HostApp with get () = _hostApp
    member internal this.GetFiles with get () =
                                        try
                                            Directory.GetFiles(_fullPath, PossibleSources)
                                        with
                                            e -> printfn "An exception was thrown enumerating the files in the source directory: %s" e.Message ; [| |]

