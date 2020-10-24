namespace voiceroid_remote_controller_library

open System
open System.Runtime.InteropServices
open System.Text
open System.Diagnostics
open System.Collections.Generic
open System.Threading

module SendMessageWithString =
    [<DllImport("user32.dll")>]
    extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, String IParam)

module SendMessageWithIntPtr =
    [<DllImport("user32.dll")>]
    extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr IParam)

module Say =

    type EnumWindowsDelegate = delegate of IntPtr * IntPtr -> bool

    [<DllImport("user32.dll")>]
    extern bool EnumChildWindows(IntPtr hWnd, EnumWindowsDelegate lpEnumFunc, IntPtr lparam);

    [<DllImport("user32.dll", CharSet=CharSet.Auto)>]
    extern int GetWindowTextLength(HandleRef hWnd)

    [<DllImport("user32.dll", CharSet=CharSet.Auto)>]
    extern int GetWindowText(HandleRef hWnd, StringBuilder lpString, int nMaxCount)

    [<DllImport("user32.dll", CharSet=CharSet.Auto)>]
    extern int GetClassName(HandleRef hWnd, StringBuilder lpClassName, int nMaxCount)

    type TalkOnVoiceroid() =

        let WM_SETTEXT = 0x000C
        let BM_CLICK = 0x00F5


        let childList = new List<IntPtr>()

        let mut exeFilePathList = new List<string>()

        let enumWindowCallBack(handle: IntPtr) (iParam: IntPtr): bool = 
            childList.Add(handle)
            true

        let mutable isCanPlay = false
        let mutable firstTextFieldHandle = IntPtr.Zero
        let mutable firstPlayButtonHandle = IntPtr.Zero

        let firstTextAreaString = "FirstTextArea"
        let firstPlayButtonString = "FirstPlayButton"
        let richTextAreaClassName = "WindowsForms10.RichEdit20W"
        let playButtonWindowTitle = "再生"

        member this.IsCanPlay 
            with get(): bool = isCanPlay
            and private set (tmp: bool) = isCanPlay <- tmp

        member this.FirstTextFieldHandle
            with get(): IntPtr = firstTextFieldHandle
            and private set (tmp: IntPtr) = firstTextFieldHandle <- tmp

        member this.FirstPlayButtonHandle
            with get(): IntPtr = firstPlayButtonHandle
            and private set (tmp: IntPtr) = firstPlayButtonHandle <- tmp

        member private this.SetWindowHandle(handle: IntPtr) =

            EnumChildWindows(handle, new EnumWindowsDelegate(enumWindowCallBack), IntPtr.Zero) |> ignore

            let mut tmp = IntPtr.Zero
            let mut count = 0
            let mut isPlayed = false
            let mut isFirst: bool = true

            let rec searchAndSetWindowHandle (l: IntPtr list) (resultList: (string * IntPtr) list) (previous: IntPtr) (isFirstText: bool) (isFirstPlayButton: bool) : (string * IntPtr) list =
                match l with
                | head :: tail -> 
                    let capacity = GetWindowTextLength(HandleRef(this, head)) * 2
                    let textStringBuilder = StringBuilder(capacity)
                    let getWindowTextResult = GetWindowText(HandleRef(this, head), textStringBuilder, textStringBuilder.Capacity)

                    let classNameCapacity = 2048 * 2
                    let classNameStringBuilder = StringBuilder(classNameCapacity)
                    let getClassNameResult = GetClassName(HandleRef(this, head), classNameStringBuilder, classNameCapacity)
                    let className = classNameStringBuilder.ToString()

                    let windowText = textStringBuilder.ToString()

                    // printfn "classname: %s, window title: %s, hwnd: %d" className windowText head

                    if className.Contains(richTextAreaClassName) && isFirstText then
                        searchAndSetWindowHandle tail (resultList @ [(firstTextAreaString, head)]) head false isFirstPlayButton
                    else
                        if windowText.Contains(playButtonWindowTitle) && isFirstPlayButton then
                            searchAndSetWindowHandle tail (resultList @ [(firstPlayButtonString, head)]) head isFirstText false
                        else
                            searchAndSetWindowHandle tail resultList head isFirstText isFirstPlayButton
                | [] -> resultList

            searchAndSetWindowHandle (Seq.toList childList) [] IntPtr.Zero true true


        member this.Say(text: string) =
            if this.IsCanPlay then 
                SendMessageWithString.SendMessage(this.FirstTextFieldHandle, WM_SETTEXT, IntPtr.Zero, text) |> ignore
                SendMessageWithIntPtr.SendMessage(this.FirstPlayButtonHandle, BM_CLICK, IntPtr.Zero, IntPtr.Zero) |> ignore
                true
            else
                false


        member private this.SearchProcess(processArray: Process[], targetWindowTitle: string): IntPtr option = 
            if processArray.Length < 1 then
                None
            else
                let rec returnValue (l: Process list) (result: IntPtr) =
                    match l with
                    | head :: tail ->
                        if head.MainWindowTitle = targetWindowTitle then
                            returnValue [] head.MainWindowHandle
                        else
                            returnValue tail result
                    | [] -> if result <> IntPtr.Zero then Some result else None
                
                returnValue (Seq.toList processArray) IntPtr.Zero

        member public this.Init(targetVoiceroidWindowTitle: string, exePath: string) =
            let processArray = Process.GetProcessesByName("VOICEROID")
                

            let searchProcessResult = this.SearchProcess(processArray, targetVoiceroidWindowTitle)

            if searchProcessResult.IsSome then
                let result = this.SetWindowHandle(searchProcessResult.Value)
                for x, y in result do
                    printfn "%s, %d" x y
                    if x = firstTextAreaString then
                        this.FirstTextFieldHandle <- y
                    else
                        if x = firstPlayButtonString then
                            this.FirstPlayButtonHandle <- y

                if this.FirstTextFieldHandle <> IntPtr.Zero && this.FirstPlayButtonHandle <> IntPtr.Zero then
                    this.IsCanPlay <- true
                else
                    this.IsCanPlay <- false
            else
                printfn "missing target software"
                printfn "starting %s" exePath
                let p = Process.Start(exePath)

                Thread.Sleep(5000)

                if p.MainWindowHandle <> IntPtr.Zero then
                    this.Init(targetVoiceroidWindowTitle, exePath)
                else
                    printfn "failed start %s" exePath
