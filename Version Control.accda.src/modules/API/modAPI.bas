Attribute VB_Name = "modAPI"
'---------------------------------------------------------------------------------------
' Module    : modAPI
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : This module exposes a set of VCS tools to other projects.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("API")

' MCP Debug log file path (set when first used)
Private m_strMCPDebugLogPath As String

' True while this project is dispatching a call on to the installed add-in. A refusal that
' arrives while it is set came back through that dispatch into the project that issued it,
' which is a different failure from a genuinely nested call and needs to say so.
' RedirectTargetIsSelf exists to make that unreachable; this reports it if it ever is not.
Private m_blnRedirecting As Boolean

' Why an API entry point turned a call away. The two cases need opposite advice, and the
' wrong one costs the caller the whole diagnosis -- see RefuseReentrantCall.
Private Enum eRefusalReason
    rrNested = 0            ' A call arrived from inside another API call
    rrSelfDispatch = 1      ' A dispatch to the installed add-in came back to its sender
End Enum

' Returned by API (and embedded in the APIAsync JSON) when a call arrives while another
' is still in progress. Published as a prefix rather than an error number because the
' refusal cannot be raised -- see RefuseReentrantCall. External callers (the MCP server
' in particular) match on this to report a refusal instead of treating it as data.
Public Const API_REFUSED_PREFIX As String = "VCS_API_REFUSED: "


' Note, some enums are listed here when they are directly exposed
' through the Options and VCS classes. (Allowing them to be used externally)
' Use hard-coded enum index values to preserve user settings in existing projects
' if new values are added in the future.

' Control the interaction mode. (This one stays here rather than in modConstants because
' SetInteractionMode below exposes it to external callers.)
Public Enum eInteractionMode
    eimNormal = 0
    eimSilent = 1
    eimNonInteractive = 2   ' Suppress UI. Unresolved decisions block; they are not approved.
End Enum

' Formats used when exporting table data.
Public Enum eTableDataExportFormat
    etdNoData = 0
    etdTabDelimited = 1
    etdXML = 2
    [_Last]
End Enum

' Sanitize levels used for sanitizing general and color elements in source files.
Public Enum eSanitizeLevel
    eslNone = 0     ' Don't sanitize anything.
    eslMinimal = 1  ' Sanitize minimal things like GUIDs.
    eslStandard = 2 ' Remove non-critical elements that cause VCS noise between builds.
    eslExtended = 3 ' Remove as much as possible. May have possible negative effects.
    [_Last]         ' Placeholder for the end of the list.
End Enum

' Controls whether connection strings with credentials are stored in a .env file.
' Export: Auto = only when UID/PWD detected; Always = all connections; Never = keep in source.
' Build: Auto = prompt to save new credentials; Always = save silently; Never = never write .env.
Public Enum eUseEnvConnections
    uecAuto = 0     ' Export when credentials detected; build prompts before saving.
    uecAlways = 1   ' Always store connection strings in .env file.
    uecNever = 2    ' Leave connection strings in source; never write .env on build.
    [_Last]
End Enum


'---------------------------------------------------------------------------------------
' Procedure : HandleRibbonCommand
' Author    : Adam Waller
' Date      : 2/28/2022
' Purpose   : Handle an incoming command from the TwinBasic ribbon COM add-in.
'           : (Allows us to keep all the logic between the XML ribbon file and the
'           :  Access add-in.)
'---------------------------------------------------------------------------------------
'
Public Function HandleRibbonCommand(strCommand As String, Optional strArgument As String) As Boolean
    ' The function is called by Application.Run which can be re-entrant but we really
    ' don't want it to be since that'd cause errors. To avoid this, we will ignore any
    ' commands while the current command is running.
    Static IsRunning As Boolean
    Dim varResult As Variant
    Dim dResult As Dictionary

    LogUnhandledErrors
    On Error GoTo ErrHandler

    If IsRunning Then
        ' Ignore the re-entry; do NOT go to clean-up.
        Exit Function
    End If

    If APIDispatchBusy Then Exit Function
    If ManualCommandRefused Then Exit Function
    IsRunning = True

    ' Make sure we are not attempting to run this from the current database when making
    ' changes to the add-in itself. (It will re-run the command through the add-in.)
    If RunningOnLocal() Then
        varResult = RunInAddIn("HandleRibbonCommand", True, strCommand, strArgument)
        GoTo CheckResult
    End If

    Dim cPriorRequest As clsOperationRequest
    Set cPriorRequest = IncomingOperationRequest
    Set IncomingOperationRequest = CreateOperationRequest(eosUserInterface)

    ' If a function is not found, this will throw an error. It is up to the ribbon
    ' designer to ensure that the control IDs match public procedures in the VCS
    ' (clsVersionControl) class module.
    ' For example, to run VCS.Export, the ribbon button ID should be named "btnExport"

    ' Trim off control ID prefix when calling command
    If Len(strArgument) Then
        varResult = CallByName(VCS, Mid(strCommand, 4), VbMethod, strArgument)
    Else
        varResult = CallByName(VCS, Mid(strCommand, 4), VbMethod)
    End If

CheckResult:
    HandleRibbonCommand = True
    If VarType(varResult) = vbString Then
        If Left$(CStr(varResult), 1) = "{" Then
            Set dResult = ParseJson(CStr(varResult))
            If dResult.Exists("success") Then HandleRibbonCommand = CBool(dResult("success"))
        End If
    End If
CleanUp:
    Set IncomingOperationRequest = cPriorRequest
    IsRunning = False
    Exit Function

ErrHandler:
    ' An error occurred so we need to make it available for further attempts
    ' but do not handle the error.
    Set IncomingOperationRequest = cPriorRequest
    IsRunning = False

    ' Re-throw
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function


'---------------------------------------------------------------------------------------
' Procedure : HandleTestAssertion
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Entry point for TestAssert calls from user projects. Must live in a
'           : standard module so Application.Run can reach it.
'---------------------------------------------------------------------------------------
'
Public Function HandleTestAssertion(ByVal blnCondition As Boolean, _
                                     Optional ByVal varContext As Variant) As Boolean
    If TestRunner.State <> etrsRunning Then
        HandleTestAssertion = False
        Exit Function
    End If
    HandleTestAssertion = True
    TestRunner.RecordAssertion blnCondition, varContext
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetCancelPollForTest
' Author    : Josh
' Date      : 09/30/2026
' Purpose   : Test seam for user projects. The running operation sees an MCP cancel
'           : request on the Nth cancel check from now, so a test can stop its own run
'           : through the path a real request takes. Pass 0 to disarm. Must live in a
'           : standard module so Application.Run can reach it.
'---------------------------------------------------------------------------------------
'
Public Sub SetCancelPollForTest(ByVal lngCancelOnCheck As Long)
    Operation.SetCancelPollForTest lngCancelOnCheck
End Sub


'---------------------------------------------------------------------------------------
' Procedure : VCS
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Wrapper for the VCS class, providing easy API access to VCS functions.
'           : *NOTE* that this class is not persisted. This allows us to wrap up and
'           : remove any object references after the call completes.
'---------------------------------------------------------------------------------------
'
Public Function VCS() As clsVersionControl
    Set VCS = New clsVersionControl
End Function


'---------------------------------------------------------------------------------------
' Procedure : Preload
' Author    : Adam Waller
' Date      : 2/22/2025
' Purpose   : A simple generic function that does nothing in particular, but allows
'           : us to call it from external applications (i.e. Ribbon) to ensure that
'           : the add-in has been loaded into Microsoft Access.
'---------------------------------------------------------------------------------------
'
Public Function Preload()
    ' Hello, Add-In!
End Function


'---------------------------------------------------------------------------------------
' Procedure : APICapabilities
' Author    : Josh
' Date      : 10/01/2026
' Purpose   : Name the API behaviours this build supports, so a client can refuse a
'           : request before it starts. Clients call this directly through
'           : Application.Run ("<add-in path>.APICapabilities"), not through API. On
'           : an add-in that predates it, Access then refuses the call with a COM
'           : error (2517, procedure not found). Asking API for a method the add-in
'           : lacks instead stops on a modal "Run-time error 438" dialog, which holds
'           : Access until a person dismisses it.
'           : A rebuild does not change the version number, so the version cannot
'           : answer this. Returns JSON: {"success":true,"capabilities":[...]}
'           :   build_as_paths: BuildAs(source, output) builds without pickers and
'           :   reports output_path on its completion callback.
'---------------------------------------------------------------------------------------
'
Public Function APICapabilities() As String

    Dim dResult As Dictionary

    Set dResult = New Dictionary
    dResult.Add "success", True
    dResult.Add "capabilities", Array("build_as_paths")
    APICapabilities = ConvertToJson(dResult)

End Function


'---------------------------------------------------------------------------------------
' Procedure : WorkerCallback
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : A public callback endpoint for worker script operations to check back in
'           : after completion.
'---------------------------------------------------------------------------------------
'
Public Function WorkerCallback(strKey As String, Optional varParams As Variant)
    If RunningOnLocal Then
        RunInAddIn "WorkerCallback", False, strKey, varParams
    Else
        Worker.ReturnWorker strKey, varParams
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetInteractionMode
' Author    : Adam Waller
' Date      : 6/24/2023
' Purpose   : Control the types of UI interaction. (For example, you might set the
'           : interaction mode to silent during an automated build.)
'---------------------------------------------------------------------------------------
'
Public Function SetInteractionMode(intMode As eInteractionMode) As String
    SetInteractionMode = SelectInteractionMode(Operation, intMode)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunningOnLocal
' Author    : Adam Waller
' Date      : 5/18/2022
' Purpose   : Returns true if the code is running in the current database instead of
'           : the add-in database.
'---------------------------------------------------------------------------------------
'
Private Function RunningOnLocal() As Boolean
    RunningOnLocal = (StrComp(CurrentProject.FullName, CodeProject.FullName, vbTextCompare) = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RedirectTargetIsSelf
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Returns true when a redirect to the installed add-in would land back in the
'           : project that is already running.
'           :
'           : `RunningOnLocal` only reports that the current database and the running code
'           : are the same file. That is true of a development copy open as the current
'           : database, where redirecting to the installed add-in is the whole point -- and
'           : equally true of the *installed* add-in open as the current database, where the
'           : redirect target is the same file. `Application.Run` then re-enters the entry
'           : point while the outer call still holds its `Static IsRunning`, so every call
'           : came back as a re-entrant refusal blaming nesting that never happened.
'           : Nothing is supposed to open the install as the current database, but a user
'           : who double-clicks the file does exactly that, and a misleading refusal sends
'           : whoever reads it looking for a caller that does not exist.
'           :
'           : Compare the extension-less form, which is exactly the string handed to
'           : `Application.Run`, so an installed .accde never reads as a different file
'           : from the .accda it was built from.
'---------------------------------------------------------------------------------------
'
Private Function RedirectTargetIsSelf() As Boolean
    RedirectTargetIsSelf = (StrComp( _
        FSO.BuildPath(FSO.GetParentFolderName(CodeProject.FullName), _
            FSO.GetBaseName(CodeProject.FullName)), _
        GetRunCmdAddInFullLibName, vbTextCompare) = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunInAddIn
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Run a proceedure with optional parameters in the VCS add-in database.
'---------------------------------------------------------------------------------------
'
Public Function RunInAddIn(strProcedure As String, blnUseTimer As Boolean, Optional varArg1 As Variant, Optional varArg2 As Variant) As Variant
    Dim strLibName As String
    Dim strRunCmd As String
    Dim cRequest As clsOperationRequest
    If ManualCommandRefused Then
        RunInAddIn = RefusalJson(ERR_OPERATION_ALREADY_RUNNING, OperationRunningMessage(), False)
        Exit Function
    End If
    If Not AddinLoaded Then LoadVCSAddIn
    If RunningOnLocal Then
        strLibName = GetRunCmdAddInFullLibName
    Else
        strLibName = PROJECT_NAME
    End If
    If blnUseTimer And Not RunningOnLocal Then
        Set cRequest = CreateOperationRequest(eosUserInterface)
        ' The current database can contain an older project with the same module
        ' names. A delayed Application.Run must return to this code database.
        strRunCmd = FSO.BuildPath(FSO.GetParentFolderName(CodeProject.FullName), FSO.GetBaseName(CodeProject.FullName)) & "." & strProcedure
        RunInAddIn = SetTimer(strRunCmd, CStr(Nz(varArg1, vbNullString)), CStr(Nz(varArg2, vbNullString)), , , , cRequest.ToJson)
    Else
        strRunCmd = strLibName & "." & strProcedure
        If Not IsMissing(varArg2) Then
            RunInAddIn = Application.Run(strRunCmd, varArg1, varArg2)
        ElseIf Not IsMissing(varArg1) Then
            RunInAddIn = Application.Run(strRunCmd, varArg1)
        Else
            RunInAddIn = Application.Run(strRunCmd)
        End If
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetRunCmdAddInFullLibName
' Author    : Josef Poetzl
' Date      : 2/20/2025
' Purpose   : Return the full path to the add-in library without the file extension.
'---------------------------------------------------------------------------------------
'
Private Function GetRunCmdAddInFullLibName() As String
    GetRunCmdAddInFullLibName = FSO.BuildPath( _
        GetInstallSettings.strInstallFolder, ADDIN_BASENAME)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExampleLoadAddInAndRunExport
' Author    : Adam Waller
' Date      : 2/21/2025
' Purpose   : This function can be copied to a local database and triggered with a
'           : command line argument or other automation technique to load the VCS
'           : add-in file and initiate an export.
'           : NOTE: This expects the add-in to be installed in the default location
'           : and using the default file name.
'---------------------------------------------------------------------------------------
'
Public Function ExampleLoadAddInAndRunExport()
    Application.Run Environ$("AppData") & "\" & PROJECT_NAME & "\" & ADDIN_BASENAME & _
        ".HandleRibbonCommand", "btnExport"
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExampleBuildFromSource
' Author    : Adam Waller
' Date      : 2/21/2025
' Purpose   : This function can be copied to a local database and triggered with a
'           : command line argument or other automation technique to load the VCS
'           : add-in file and build this project from source.
'           : NOTE: This expects the add-in to be installed in the default location
'           : and using the default file name.
'---------------------------------------------------------------------------------------
'
Public Function ExampleBuildFromSource(Optional strSourcePath As String)
    Application.Run Environ$("AppData") & "\" & PROJECT_NAME & "\" & ADDIN_BASENAME & _
        ".HandleRibbonCommand", "btnBuild", strSourcePath
End Function


'---------------------------------------------------------------------------------------
' Procedure : API
' Author    : Adam Waller
' Date      : 1/22/2026
' Purpose   : Public API function for external automation (e.g., MCP tools) to call
'           : VCS methods via Application.Run and receive return values.
'           : Routes method calls to clsVersionControl class methods.
'           : Usage: Application.Run("Version Control.API", "GetVCSVersion")
'---------------------------------------------------------------------------------------
'
Public Function API(strMethod As String, _
    Optional varArg1 As Variant, _
    Optional varArg2 As Variant, _
    Optional varArg3 As Variant, _
    Optional ByVal strSessionEnvelope As String = vbNullString) As Variant

    ' The function is called by Application.Run, which can be re-entrant. A nested call
    ' that would change state the outer call owns is refused, and says so in the return
    ' value rather than coming back Empty -- see RefuseReentrantCall for why silence was
    ' worse and why this is not an Err.Raise. Reads, cancellation and retired policy
    ' calls (RequestCanInspect) are admitted beside running work.
    Dim colArgs As New Collection
    Dim cRequest As clsOperationRequest
    Dim strLibName As String
    Dim strRunCmd As String

    If strMethod = "GetVCSVersion" Then
        API = CStr(CodeDb.Properties("AppVersion"))
        Exit Function
    End If
    If strMethod = "GetCapabilities" Then
        API = APICapabilities()
        Exit Function
    End If
    If Len(strSessionEnvelope) = 0 Or strSessionEnvelope <> CurrentCompatibilityEnvelope() Or _
        Not CompatibilityDispatchAllowed Then
        API = SessionFailure("missing_stale_or_foreign_session")
        Exit Function
    End If

    SuppressErrorBreaks
    LogUnhandledErrors
    On Error GoTo ErrHandler

    If Not RequestCanInspect(strMethod) Then
        If APIDispatchBusy Then
            API = RefuseReentrantCall(strMethod, "API", RefusalReason)
            RestoreErrorBreaks
            Exit Function
        End If
        If Operation.IsActive Or TimerIsPending Then
            ' A staged root or pending timer owns the shared operation state. Dismiss
            ' validates its own action: close refuses while work is active, and cancel
            ' must name the active operation's ID (A50; docs/noninteractive-dialogs.md).
            API = RefuseReentrantCall(strMethod, "API", rrNested)
            RestoreErrorBreaks
            Exit Function
        End If
    End If

    ' Make sure we are not attempting to run this from the current database when making
    ' changes to the add-in itself. (It will re-run the command through the add-in.)
    ' Only redirect when the installed add-in is a *different* file than the one running:
    ' where they are the same, the redirect re-enters this function and refuses itself.
    ' See RedirectTargetIsSelf.
    If RunningOnLocal() And Not RedirectTargetIsSelf() And Not CompatibilityDispatchAllowed Then
        ' When running from within the add-in database, we need to use the full path
        ' to ensure we call the add-in version, not a local version.
        strLibName = GetRunCmdAddInFullLibName
        strRunCmd = strLibName & ".API"

        ' Use Application.Run to call the add-in version, which will return the value
        m_blnRedirecting = True
        If Not IsMissing(varArg3) Then
            API = Application.Run(strRunCmd, strMethod, varArg1, varArg2, varArg3)
        ElseIf Not IsMissing(varArg2) Then
            API = Application.Run(strRunCmd, strMethod, varArg1, varArg2)
        ElseIf Not IsMissing(varArg1) Then
            API = Application.Run(strRunCmd, strMethod, varArg1)
        Else
            API = Application.Run(strRunCmd, strMethod)
        End If
        GoTo CleanUp
    End If

    ' The request is captured before dispatch; the method's root takes its source and
    ' public ID from it. Every parameter of a method reachable from here must be
    ' declared ByVal: CallByName cannot bind a Variant to a ByRef typed parameter.
    CountCompatibilityDispatch
    If Not IsMissing(varArg1) Then colArgs.Add varArg1
    If Not IsMissing(varArg2) Then colArgs.Add varArg2
    If Not IsMissing(varArg3) Then colArgs.Add varArg3
    Set cRequest = CreateOperationRequest(eosExternalAPI)
    API = DispatchOperationRequest(cRequest, strMethod, colArgs)

CleanUp:
    m_blnRedirecting = False
    RestoreErrorBreaks
    Exit Function

ErrHandler:
    ' An error occurred so we need to make it available for further attempts
    ' but do not handle the error.
    m_blnRedirecting = False
    RestoreErrorBreaks

    ' Re-throw
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function


'---------------------------------------------------------------------------------------
' Procedure : APIAsync
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Public API function for async operations with MCP callbacks.
'           : Routes async operations through timer mechanism with callback support.
'           : The request (source, public ID from the callback, reservation) is
'           : captured and its ID claimed before the timer is armed. A refusal is
'           : returned and posted to this request's own callback.
'           : Usage: Application.Run("Version Control.APIAsync", strCallbackInfo, "Export")
'---------------------------------------------------------------------------------------
'
Public Function APIAsync(strCallbackInfo As String, strMethod As String, _
    Optional varArg1 As Variant, Optional varArg2 As Variant, Optional varArg3 As Variant, _
    Optional ByVal strSessionEnvelope As String = vbNullString) As String

    ' The function is called by Application.Run, which can be re-entrant. We really don't
    ' want it to be, since a nested call would run against Operation state the outer call
    ' owns. Refuse it, and say so in the return value rather than coming back Empty --
    ' see RefuseReentrantCall for why silence was worse and why this is not an Err.Raise.
    Static IsRunning As Boolean
    Dim cRequest As clsOperationRequest
    Dim colArgs As New Collection
    Dim dResult As New Dictionary
    Dim strRefusal As String
    Dim strQueue As String
    Dim strLibName As String
    Dim strRunCmd As String
    Dim lngTimeoutMs As Long
    Dim blnOwnsEntry As Boolean
    Dim lngErr As Long
    Dim strErr As String

    If Len(strSessionEnvelope) = 0 Or strSessionEnvelope <> CurrentCompatibilityEnvelope() Or _
        Not CompatibilityDispatchAllowed Then
        APIAsync = PostCompatibilityRefusal(strCallbackInfo, SessionFailure("missing_stale_or_foreign_session"))
        Exit Function
    End If

    SuppressErrorBreaks
    LogUnhandledErrors
    On Error GoTo ErrHandler

    If IsRunning Or APIDispatchBusy Then
        APIAsync = RefuseAsyncRequest(strCallbackInfo, strMethod)
        RestoreErrorBreaks
        Exit Function
    End If

    IsRunning = True
    blnOwnsEntry = True

    ' Make sure we are not attempting to run this from the current database when making
    ' changes to the add-in itself. (It will re-run the command through the add-in.)
    ' Same self-dispatch guard as API -- see RedirectTargetIsSelf.
    If RunningOnLocal() And Not RedirectTargetIsSelf() And Not CompatibilityDispatchAllowed Then
        ' When running from within the add-in database, we need to use the full path
        ' to ensure we call the add-in version, not a local version.
        strLibName = GetRunCmdAddInFullLibName
        strRunCmd = strLibName & ".APIAsync"

        ' Use Application.Run to call the add-in version
        m_blnRedirecting = True
        If Not IsMissing(varArg3) Then
            APIAsync = Application.Run(strRunCmd, strCallbackInfo, strMethod, varArg1, varArg2, varArg3)
        ElseIf Not IsMissing(varArg2) Then
            APIAsync = Application.Run(strRunCmd, strCallbackInfo, strMethod, varArg1, varArg2)
        ElseIf Not IsMissing(varArg1) Then
            APIAsync = Application.Run(strRunCmd, strCallbackInfo, strMethod, varArg1)
        Else
            APIAsync = Application.Run(strRunCmd, strCallbackInfo, strMethod)
        End If
        GoTo CleanUp
    End If

    ' Admission precedes every change to the shared operation or timer request.
    ' A staged root owns its callback too, even though its API stack has unwound.
    If Operation.IsActive Or TimerIsPending Then
        APIAsync = RefuseAsyncRequest(strCallbackInfo, strMethod)
        GoTo CleanUp
    End If

    If Not IsMissing(varArg1) Then colArgs.Add varArg1
    If Not IsMissing(varArg2) Then colArgs.Add varArg2
    If Not IsMissing(varArg3) Then colArgs.Add varArg3
    If Len(strCallbackInfo) > 0 Then
        Set cRequest = CreateOperationRequest(eosMCPTool, strCallbackInfo)
    Else
        Set cRequest = CreateOperationRequest(eosExternalAPI)
    End If

    ' Determine if this is an async operation or should fall back to sync
    Select Case LCase$(strMethod)
        Case "export", "fullexport", "exportvba", "build", "buildas", "mergebuild", "runfilteredtests", _
            "exportbytype", "importbytype", "exportobject", "importobject"
            ' These are async operations - spawn via timer with callback support.
            ' The callback is stored with the admitted timer, never before admission.
            strRefusal = ClaimOperationRequest(cRequest)
            If Len(strRefusal) > 0 Then
                APIAsync = PostCompatibilityRefusal(strCallbackInfo, strRefusal)
                GoTo CleanUp
            End If
            strQueue = SetTimer("APIAsyncOperation", strMethod, ConvertToJson(colArgs), , , strCallbackInfo, cRequest.ToJson)
            If Not TimerRequestSucceeded(strQueue) Then
                APIAsync = PostCompatibilityRefusal(strCallbackInfo, strQueue)
                GoTo CleanUp
            End If

            ' Determine timeout based on operation type
            Select Case LCase$(strMethod)
                Case "export", "mergebuild"
                    lngTimeoutMs = 300000   ' 5 minutes
                Case "exportvba"
                    lngTimeoutMs = 120000   ' 2 minutes
                Case Else
                    lngTimeoutMs = 600000   ' 10 minutes
            End Select

            ' Return async response immediately
            dResult.Add "success", True
            dResult.Add "async", True
            dResult.Add "operation_id", cRequest.PublicId
            dResult.Add "timeout_ms", lngTimeoutMs
            APIAsync = ConvertToJson(dResult)

        Case Else
            ' Fall back to sync dispatch for quick operations
            CountCompatibilityDispatch
            dResult.Add "sync", True
            dResult.Add "result", DispatchOperationRequest(cRequest, strMethod, colArgs)
            APIAsync = ConvertToJson(dResult)
    End Select

CleanUp:
    If blnOwnsEntry Then
        IsRunning = False
        m_blnRedirecting = False
    End If
    RestoreErrorBreaks
    Exit Function

ErrHandler:
    ' Admission errors belong to this incoming caller, never the running MCP singleton.
    lngErr = Err.Number
    strErr = Err.Description
    Err.Clear
    APIAsync = AsyncRequestFailure(strCallbackInfo, strErr, lngNumber:=lngErr)
    Resume CleanUp

End Function


'---------------------------------------------------------------------------------------
' Procedure : RefuseAsyncRequest
' Purpose   : Return and post this request's refusal without borrowing a running root,
'           : callback, journal, terminal flag, or log path.
'---------------------------------------------------------------------------------------
Public Function RefuseAsyncRequest(ByVal strCallbackInfo As String, ByVal strMethod As String) As String
    Dim strPattern As String
    strPattern = ERR_OPERATION_ALREADY_RUNNING
    If RefusalReason = rrSelfDispatch Then strPattern = "api_self_dispatch"
    RefuseAsyncRequest = AsyncRequestFailure(strCallbackInfo, _
        RefuseReentrantCall(strMethod, "APIAsync", RefusalReason), strPattern)
End Function


'---------------------------------------------------------------------------------------
' Procedure : AsyncRequestFailure
' Purpose   : An incoming request owns its failure emitter. Do not use RuntimeErrorJson
'           : here: it reads the shared operation.
'---------------------------------------------------------------------------------------
Public Function AsyncRequestFailure(ByVal strCallbackInfo As String, ByVal strMessage As String, _
    Optional ByVal strPattern As String, Optional ByVal lngNumber As Long) As String

    Dim dResult As New Dictionary
    Dim cCallback As New clsMCP

    dResult.Add "success", False
    dResult.Add "error", strMessage
    If Len(strPattern) > 0 Then
        dResult.Add "error_pattern", strPattern
        dResult.Add "api_refused", True
    End If
    If lngNumber <> 0 Then
        dResult.Add "errorNumber", lngNumber
        dResult.Add "runtime_error", strMessage
    End If
    AsyncRequestFailure = ConvertToJson(dResult)
    cCallback.RegisterCallback strCallbackInfo
    If cCallback.IsActive Then cCallback.PostCallback "error", -1, -1, strMessage, dResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : RefuseReentrantCall
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Build the message for a call that arrived while another is still running.
'           :
'           : The reason decides the advice, and getting it wrong is expensive. A nested
'           : call is the caller's own doing and the message says how to restructure it.
'           : A self-dispatched one is a defect here, and telling that caller to
'           : restructure sends it looking for a different database to host the call from
'           : -- a search that ends somewhere plausible and wrong, such as the sample
'           : database in Testing\. Name the defect instead.
'           :
'           : This used to be a bare `Exit Function`, returning Empty -- indistinguishable
'           : from a method that legitimately returned nothing. The most common way to
'           : reach it makes that silence actively misleading: code submitted through
'           : vcs_run_vba is itself delivered via API, so anything it calls back into API
'           : is nested by construction, and every call coming back Empty reads as a
'           : broken add-in rather than a refused call.
'           :
'           : Returning a marked string rather than raising is deliberate, and was
'           : arrived at the hard way. An Err.Raise here does not reach the caller: an
'           : error raised inside a library database does not propagate across
'           : Application.Run into the calling project's handler, so even a caller with
'           : `On Error GoTo` active gets a modal "Run-time error" dialog that blocks
'           : Access until a human dismisses it. Since this guard only ever trips on a
'           : nested call -- precisely the case that crosses that boundary -- raising is
'           : guaranteed to hit it. A blocking dialog is worse for automation than the
'           : silence it replaced. Callers match API_REFUSED_PREFIX to tell a refusal
'           : from a real return value.
'---------------------------------------------------------------------------------------
'
Private Function RefuseReentrantCall(strMethod As String, strEntryPoint As String, _
    Optional intReason As eRefusalReason = rrNested) As String

    Dim strMsg As String

    strMsg = API_REFUSED_PREFIX & "VCS " & strEntryPoint & " refused a call to '" & _
        strMethod & "': "

    Select Case intReason
        Case rrSelfDispatch
            strMsg = strMsg & "the call was dispatched to the installed add-in and " & _
                "arrived back in the project that sent it, so it refused itself. " & _
                "Nothing ran. This is a defect in the add-in's own dispatch, not " & _
                "something the caller can host its way around: another database will " & _
                "not help, and the sample database in the Testing folder is not a " & _
                "workaround. Report it with the method name above."
        Case Else
            strMsg = strMsg & "another API command is still running. Note that code " & _
                "executed through vcs_run_vba already runs inside an API call, so it " & _
                "cannot call back into the API; invoke API methods directly " & _
                "(vcs_call_vba) instead."
    End Select

    RefuseReentrantCall = strMsg

End Function


'---------------------------------------------------------------------------------------
' Procedure : RefusalReason
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Classify a refusal about to be returned. A dispatch still in flight from
'           : this project means the call we are turning away is the one we just sent.
'---------------------------------------------------------------------------------------
'
Private Function RefusalReason() As eRefusalReason
    If m_blnRedirecting Then
        RefusalReason = rrSelfDispatch
    Else
        RefusalReason = rrNested
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : MCPDebugLog
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Write debug messages to a file for MCP callback troubleshooting.
'           : Writes to logs/MCP_Debug.log in the source folder.
'           : Do not call this from the per-callback streaming path (PostCallback
'           : log/progress). File appends dominated agentic build time.
'           : A failed write is dropped silently. Most calls run before the MCP
'           : call's root begins, where a logged error would open a message box.
'---------------------------------------------------------------------------------------
'
Public Sub MCPDebugLog(strMessage As String)

    ' No logging unless MCP is active
    If Not MCP.IsActive Then Exit Sub

    ' Initialize log path if needed
    If m_strMCPDebugLogPath = vbNullString Then
        m_strMCPDebugLogPath = Options.GetExportFolder & "logs\MCP_Debug.log"
    End If

    ' Append to log file
    TryAppendToFile Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & strMessage, m_strMCPDebugLogPath

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PreferFullPowerCore
' Author    : Adam Waller
' Date      : 8/28/2026
' Purpose   : Public entry for Worker.vbs / Application.Run. Turns EcoQoS off
'           : on this Access process so MCP-launched work prefers a P-core.
'---------------------------------------------------------------------------------------
'
Public Sub PreferFullPowerCore()
    PreferFullPowerCurrentProcess
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DispatchOperationRequest
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Install and restore typed incoming request context around method dispatch.
'---------------------------------------------------------------------------------------
'
Public Function DispatchOperationRequest(ByVal cRequest As clsOperationRequest, _
    ByVal strMethod As String, ByVal colArgs As Collection) As Variant

    Dim cPrior As clsOperationRequest
    Dim blnPriorBusy As Boolean
    Dim strRefusal As String
    Dim blnInspection As Boolean
    Dim strError As String
    blnInspection = RequestCanInspect(strMethod)
    If APIDispatchBusy And Not blnInspection Then
        DispatchOperationRequest = RefusalJson("api_busy", T("Another API command is still running."), False)
        Exit Function
    End If
    If Not blnInspection Then
        strRefusal = RequestAdmissionRefusal(blnQuiet:=True)
        If Len(strRefusal) > 0 Then
            DispatchOperationRequest = strRefusal
            Exit Function
        End If
    End If
    Set cPrior = IncomingOperationRequest
    blnPriorBusy = APIDispatchBusy
    Set IncomingOperationRequest = cRequest
    APIDispatchBusy = True
    InitializeRequestErrors "DispatchOperationRequest"
    On Error GoTo Failed
    Select Case colArgs.Count
        Case 0: DispatchOperationRequest = CallByName(VCS, strMethod, VbMethod)
        Case 1: DispatchOperationRequest = CallByName(VCS, strMethod, VbMethod, RequestArgument(colArgs(1)))
        Case 2: DispatchOperationRequest = CallByName(VCS, strMethod, VbMethod, RequestArgument(colArgs(1)), RequestArgument(colArgs(2)))
        Case 3: DispatchOperationRequest = CallByName(VCS, strMethod, VbMethod, RequestArgument(colArgs(1)), RequestArgument(colArgs(2)), RequestArgument(colArgs(3)))
        Case Else: DispatchOperationRequest = RefusalJson("unsupported_arguments", T("This method accepts at most three API arguments."), False)
    End Select
CleanUp:
    Set IncomingOperationRequest = cPrior
    APIDispatchBusy = blnPriorBusy
    Exit Function
Failed:
    strError = Err.Description
    DispatchOperationRequest = RefusalJson("api_error", strError, False)
    Resume CleanUp
End Function


'---------------------------------------------------------------------------------------
' Procedure : RequestCanInspect
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Allow present read/cancel/retirement methods without admitting conflicting mutations.
'---------------------------------------------------------------------------------------
'
Private Function RequestCanInspect(ByVal strMethod As String) As Boolean
    Select Case LCase$(strMethod)
        Case "listaddindialogs", "dismissaddindialog", "getlogcontent", "getvcsversion", "isdatabaseopen", "getprojectname", "isvbacompiled", "getoption", "getexportfolder", "setoperationpolicy", "clearoperationpolicy", "setinteractionmode"
            RequestCanInspect = True
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : RequestArgument
' Author    : Codex
' Date      : 10/16/2026
' Purpose   : Restore queued JSON arrays as typed VBA method arguments.
'---------------------------------------------------------------------------------------
'
Private Function RequestArgument(ByVal varValue As Variant) As Variant
    Dim varArray() As Variant
    Dim lngIdx As Long
    If IsObject(varValue) Then
        If TypeName(varValue) = "Collection" Then
            If varValue.Count = 0 Then
                RequestArgument = Array()
                Exit Function
            End If
            ReDim varArray(0 To varValue.Count - 1)
            For lngIdx = 1 To varValue.Count
                varArray(lngIdx - 1) = RequestArgument(varValue(lngIdx))
            Next lngIdx
            RequestArgument = varArray
        Else
            Set RequestArgument = varValue
        End If
    Else
        RequestArgument = varValue
    End If
End Function
