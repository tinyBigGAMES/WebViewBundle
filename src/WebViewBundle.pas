{===============================================================================
  __      __   _  __   ___            ___              _ _
  \ \    / /__| |_\ \ / (_)_____ __ _| _ )_  _ _ _  __| | |___™
   \ \/\/ / -_) '_ \ V /| / -_) V  V / _ \ || | ' \/ _` | / -_)
    \_/\_/\___|_.__/\_/ |_\___|\_/\_/|___/\_,_|_||_\__,_|_\___|
       Bundle your HTML UI. No server required. Just run.

 Copyright © 2025-present tinyBigGAMES™ LLC
 All Rights Reserved.

 https://github.com/tinyBigGAMES/WebViewBundle

 See LICENSE file for license information
===============================================================================}

unit WebViewBundle;

{$I WebViewBundle.Defines.inc}

interface

uses
  WinApi.Windows,
  System.SysUtils,
  System.IOUtils,
  System.Classes,
  System.Math,
  System.IniFiles,
  System.Zip,
  VCL.Forms,
  Dlluminator,
  Winapi.EdgeUtils,
  WebViewBundle.Utils,
  WebViewBundle.Form;

type
  /// <summary>
  ///   Represents a lightweight framework for bundling HTML, CSS, JavaScript,
  ///   and assets into a single executable using the Microsoft Edge WebView2 engine.
  ///   <para>
  ///   WebViewBundle embeds a ZIP-packed virtual file system directly into your EXE
  ///   and serves it using WebView2's virtual host mapping (e.g., <c>app://</c> scheme),
  ///   eliminating the need for a separate server, installer, or runtime file extraction.
  ///   </para>
  ///   <para>
  ///   It is ideal for creating standalone tools, dashboards, hybrid applications,
  ///   and offline frontends where the UI is built with web technologies and powered by WebView2.
  ///   </para>
  /// </summary>
  TWebViewBundle = class
  protected
    FIndexPage: string;
    FWindowTitle: string;
    FWindowWidth: Cardinal;
    FWindowHeight: Cardinal;
    FWindowResizeable: Boolean;
    FBundleFilename: string;

    FExeIconFilename: string;
    FAddVersionInfo: Boolean;
    FMajorVer: Cardinal;
    FMinorVer: Cardinal;
    FPatchVer: Cardinal;
    FProductName: string;
    FDescription: string;
    FCompanyName: string;
    FCopyright: string;
    FUserDataFolder: string;

    function AddPayload(const AExeFilename, AZipFilename: string): Boolean;
    function UpdateManifest(const AExeFilename: string): Boolean;
    function UpdatePayloadIcon(const AExeFilename: string): Boolean;
    function UpdatePayloadVersionInfo(const AExeFilename: string): Boolean;
  public
    /// <summary>
    ///   Creates a new instance of the <c>TWebViewBundle</c> class.
    ///   This constructor initializes any internal fields or resources necessary
    ///   for operating a WebViewBundle instance. It prepares the object to be used
    ///   for further operations such as loading and displaying bundled web content.
    /// </summary>
    constructor Create(); virtual;

    /// <summary>
    ///   Destroys the current instance of the <c>TWebViewBundle</c> class.
    ///   This destructor ensures that all allocated resources associated with the
    ///   instance are properly released and cleaned up to prevent memory leaks.
    ///   It is automatically invoked when the object instance is freed.
    /// </summary>
    destructor Destroy(); override;

    /// <summary>
    ///   Retrieves the version string associated with the <c>TWebViewBundle</c> class.
    ///   This method provides a way to obtain the current version of the framework or
    ///   library, which may represent the build version, release identifier, or internal
    ///   versioning information.
    /// </summary>
    /// <returns>
    ///   A <c>string</c> containing the version information. Typically, this will be
    ///   formatted as a semantic version number (e.g., "1.0.0") or another
    ///   recognizable versioning format.
    /// </returns>
    class function GetVersion(): string;

    /// <summary>
    ///   Determines whether the current executable contains an embedded HTML payload.
    ///   The payload is expected to be packaged in ZIP format and compiled into the executable
    ///   as a resource or binary section.
    /// </summary>
    /// <returns>
    ///   <c>True</c> if a compiled HTML (ZIP) payload is present within the executable;
    ///   otherwise, <c>False</c>.
    /// </returns>
    class function HasPayload(): Boolean;

    /// <summary>
    ///   Checks whether the Microsoft WebView2 Runtime is installed on the system.
    ///   This method verifies the presence of the WebView2 Runtime, which is required
    ///   for displaying web content within the application using the WebView2 control.
    /// </summary>
    /// <returns>
    ///   <c>True</c> if the WebView2 Runtime is installed and available for use;
    ///   otherwise, <c>False</c>.
    /// </returns>
    function IsWebView2RuntimeInstalled(): Boolean;

    /// <summary>
    ///   Loads the compiled HTML payload from the embedded resource or binary section
    ///   and displays it inside the WebView2 control.
    ///   This procedure extracts the HTML content from the embedded ZIP archive and
    ///   maps it for use with the WebView2 virtual host to render web content directly
    ///   from memory, without requiring external files or a running web server.
    /// </summary>
    procedure LoadPayload();

    /// <summary>
    ///   Creates a ZIP archive containing the HTML files located in the specified folder
    ///   and all of its subfolders.
    ///   This method recursively traverses the given directory structure, adds all HTML-related
    ///   content to a ZIP file, and provides optional progress notifications through a callback.
    /// </summary>
    /// <param name="AFolderPath">
    ///   The path to the root folder containing the HTML, CSS, JavaScript, and other assets
    ///   to be packaged into the ZIP archive.
    /// </param>
    /// <param name="ADestZipFilename">
    ///   The destination filename where the resulting ZIP archive will be saved.
    /// </param>
    /// <param name="AProgressCallback">
    ///   An optional callback function that is invoked to report progress during the zipping operation.
    ///   This can be used to provide user feedback or cancel the operation.
    /// </param>
    /// <param name="AUserData">
    ///   A user-defined pointer that will be passed to the progress callback, allowing custom
    ///   context or state information to be supplied during the operation.
    /// </param>
    /// <returns>
    ///   <c>True</c> if the folder was successfully zipped into the destination file;
    ///   otherwise, <c>False</c>.
    /// </returns>
    function ZipHtml(const AFolderPath, ADestZipFilename: string; const AProgressCallback: TwvbZipProgressCallback; const AUserData: Pointer): Boolean;

    /// <summary>
    ///   Creates a bundled executable by copying the current application and embedding
    ///   a compiled HTML ZIP archive into the new executable file.
    ///   This method takes an existing compiled ZIP file and attaches it to the specified
    ///   output executable, producing a standalone application that includes the bundled
    ///   web content.
    /// </summary>
    /// <param name="AZipFilename">
    ///   The filename of the ZIP archive containing the compiled HTML, CSS, JavaScript,
    ///   and related assets to embed.
    /// </param>
    /// <param name="AOutputExeFilename">
    ///   The destination filename for the new executable that will contain the bundled payload.
    /// </param>
    /// <returns>
    ///   <c>True</c> if the executable was successfully created with the bundled HTML payload;
    ///   otherwise, <c>False</c>.
    /// </returns>
    function BundleHtml(const AZipFilename, AOutputExeFilename: string): Boolean; overload;

    /// <summary>
    ///   Specifies the default HTML page to display when the bundled content is loaded.
    ///   This typically refers to an entry page such as <c>index.html</c> within the bundled
    ///   ZIP archive or embedded payload.
    /// </summary>
    property IndexPage: string read FIndexPage write FIndexPage;

    /// <summary>
    ///   Defines the text that will be displayed in the window's title bar.
    ///   This property allows customization of the main application window's caption
    ///   when presenting the bundled web content.
    /// </summary>
    property WindowTitle: string read FWindowTitle write FWindowTitle;

    /// <summary>
    ///   Specifies the filename of an external compiled HTML ZIP archive to be loaded
    ///   if no embedded payload is present within the executable.
    ///   This provides a fallback mechanism to load external web content during development
    ///   or deployment without requiring rebundling.
    /// </summary>
    property BundleFilename: string read FBundleFilename write FBundleFilename;

    /// <summary>
    ///   Sets the initial width of the application window, in pixels.
    ///   This property defines the horizontal dimension of the window
    ///   when it is first created and displayed.
    /// </summary>
    property WindowWidth: Cardinal read FWindowWidth write FWindowWidth;

    /// <summary>
    ///   Sets the initial height of the application window, in pixels.
    ///   This property defines the vertical dimension of the window
    ///   when it is first created and displayed.
    /// </summary>
    property WindowHeight: Cardinal read FWindowHeight write FWindowHeight;

    /// <summary>
    ///   Determines whether the application window can be resized by the user.
    ///   When set to <c>True</c>, the window can be freely resized; otherwise, it will
    ///   maintain a fixed size based on <see cref="WindowWidth"/> and <see cref="WindowHeight"/>.
    /// </summary>
    property WindowResizeable: Boolean read FWindowResizeable write FWindowResizeable;

    /// <summary>
    ///   Specifies the filename of an icon to be embedded into the output executable.
    ///   If set, the icon file will be used to replace or assign the application icon
    ///   during the bundling process, enhancing the branding and appearance of the generated EXE.
    /// </summary>
    property ExeIconFilename: string read FExeIconFilename write FExeIconFilename;

    /// <summary>
    ///   Determines whether version information should be added to the output executable.
    ///   When set to <c>True</c>, standard Windows version information such as product name,
    ///   version number, company name, and other metadata will be embedded into the bundled EXE.
    /// </summary>
    property AddVersionInfo: Boolean read FAddVersionInfo write FAddVersionInfo;
    /// <summary>
    ///   Specifies the major version number to be included in the output executable's
    ///   version information metadata.
    ///   This typically represents significant releases that introduce major new features
    ///   or changes to the bundled application.
    /// </summary>
    property MajorVer: Cardinal read FMajorVer write FMajorVer;

    /// <summary>
    ///   Specifies the minor version number to be included in the output executable's
    ///   version information metadata.
    ///   Minor versions usually indicate smaller feature additions, improvements,
    ///   or minor changes that are backward-compatible.
    /// </summary>
    property MinorVer: Cardinal read FMinorVer write FMinorVer;

    /// <summary>
    ///   Specifies the patch version number to be included in the output executable's
    ///   version information metadata.
    ///   Patch versions generally represent bug fixes, security updates, or other minor
    ///   corrections that do not introduce new features or breaking changes.
    /// </summary>
    property PatchVer: Cardinal read FPatchVer write FPatchVer;
    /// <summary>
    ///   Specifies the product name to be embedded in the output executable's version
    ///   information metadata.
    ///   This value typically identifies the application name or bundled product,
    ///   and is displayed in the Windows file properties dialog under the "Product Name" field.
    /// </summary>
    property ProductName: string read FProductName write FProductName;

    /// <summary>
    ///   Specifies the description to be embedded in the output executable's version
    ///   information metadata.
    ///   This value provides a short textual explanation or tagline about the application,
    ///   typically displayed in the Windows file properties under "File Description."
    /// </summary>
    property Description: string read FDescription write FDescription;

    /// <summary>
    ///   Specifies the company name to be embedded in the output executable's version
    ///   information metadata.
    ///   This value identifies the publisher or organization responsible for producing
    ///   the bundled application.
    /// </summary>
    property CompanyName: string read FCompanyName write FCompanyName;

    /// <summary>
    ///   Specifies the copyright statement to be embedded in the output executable's
    ///   version information metadata.
    ///   This value typically includes the copyright holder's name and year,
    ///   and appears in the Windows file properties under the "Copyright" field.
    /// </summary>
    property Copyright: string read FCopyright write FCopyright;
    /// <summary>
    ///   Specifies the folder path where WebView2 will store cached data, configuration files,
    ///   and temporary files used by the <c>TWebViewBundle</c> instance.
    ///   This directory is required by the WebView2 engine to maintain browser state, local storage,
    ///   cookies, and other session-related data.
    /// </summary>
    property UserDataFolder: string read FUserDataFolder write FUserDataFolder;

  end;

implementation

{ TWebViewBundle }
function TWebViewBundle.UpdateManifest(const AExeFilename: string): Boolean;
begin
  Result := False;
  if not TFile.Exists(AExeFilename) then Exit;
  Result := wvbAddResManifestFromResource(CwvbExeManifestResName, AExeFilename)
end;

function TWebViewBundle.UpdatePayloadIcon(const AExeFilename: string): Boolean;
var
  LFilename: string;
begin
  Result := False;
  if FExeIconFilename.IsEmpty then Exit;

  LFilename := TPath.GetFullPath(FExeIconFilename);
  if not TFile.Exists(AExeFilename) then Exit;
  if not TFile.Exists(LFilename) then Exit;
  if not wvbIsValidWin64PE(AExeFilename) then Exit;
  wvbUpdateIconResource(AExeFilename,  LFilename);
  Result := True;
end;

function TWebViewBundle.UpdatePayloadVersionInfo(const AExeFilename: string): Boolean;
begin
  Result := False;
  if not TFile.Exists(AExeFilename) then Exit;
  if not wvbIsValidWin64PE(AExeFilename) then Exit;
  wvbUpdateVersionInfoResource(AExeFilename, FMajorVer, FMinorVer, FPatchVer, FProductName,
    FDescription, TPath.GetFileName(AExeFilename), FCompanyName, FCopyright);
  Result := True;
end;

constructor TWebViewBundle.Create();
begin
  inherited;

  FIndexPage := 'index.html';
  FWindowTitle := 'WebViewBundle';
  FWindowWidth := 960;
  FWindowHeight := 720;
  FWindowResizeable := True;
  FBundleFilename := 'WebViewBundle.zip';

  FExeIconFilename := '';
  FAddVersionInfo := False;
  FMajorVer := 1;
  FMinorVer := 0;
  FPatchVer := 0;
  FProductName := 'Your Project Name';
  FDescription := 'Your Product Description';
  FCompanyName := 'Your Company Name';
  FCopyright := 'Copyright (c) 2025-present';
  FUserDataFolder := 'data';
end;

destructor TWebViewBundle.Destroy();
begin
  inherited;
end;

class function TWebViewBundle.GetVersion(): string;
begin
  Result := '0.1.0';
end;

class function TWebViewBundle.HasPayload(): Boolean;
begin
  Result := wvbResourceExist(CwvbWebAppResName);
end;

function TWebViewBundle.IsWebView2RuntimeInstalled(): Boolean;
begin
  Result := True;

  if not wvbResourceExist(CwvWebview2SetupResName) then
  begin
    wvbShowMessage('Error', 'Invalid WebViewBundle App', [], smError);
    Result := False;
    Exit;
  end;

  if WebView2RuntimeInstalled() then Exit;
  if wvbConfirmDlg('This app requires the Microsoft Edge WebView2 runtime to be installed. It will be downloaded and installed for you now.', []) <> cdYes then
  begin
    Result := False;
    Exit;
  end;

  wvbExtractAndRunResEXE(CwvWebview2SetupResName, '/install', FUserDataFolder);
end;

procedure TWebViewBundle.LoadPayload();
var
  LForm: TWebViewBundleForm;
  LIni: TIniFile;
  LFilename: string;
  LStringList: TStringList;
begin
  LFilename := TPath.ChangeExtension(ParamStr(0), 'ini');
  LIni := TIniFile.Create(TPath.GetFullPath(LFilename));
  try
    LForm := TWebViewBundleForm.Create(nil);
    try

      if wvbResourceExist(CwvbConfigResName) then
      begin
        LStringList := TStringList.Create();
        try
          wvbLoadStringListFromResource(CwvbConfigResName, LStringList);
          FWindowTitle := LStringList.Values['WindowTitle'];
          FWindowWidth := LStringList.Values['WindowWidth'].ToInteger();
          FWindowHeight := LStringList.Values['WindowHeight'].ToInteger();
          FWindowResizeable := LStringList.Values['WindowResizeable'].ToBoolean();
          FIndexPage := LStringList.Values['IndexPage'];
          FUserDataFolder := LStringList.Values['UserDataFolder'];
        finally
          LStringList.Free();
        end;
      end;

      LForm.WindowTitle := FWindowTitle;
      LForm.IndexPage := FIndexPage;
      LForm.UserDataFolder := FUserDataFolder;
      LForm.BundleFilename := FBundleFilename;
      if FWindowResizeable then
        begin
          LForm.BorderStyle := bsSizeable;
          LForm.BorderIcons := [biSystemMenu, biMinimize, biMaximize];
        end
      else
        begin
          LForm.BorderStyle := bsSingle;
          LForm.BorderIcons := [biSystemMenu, biMinimize];
        end;

      LForm.Width  := FWindowWidth;
      LForm.Height := FWindowHeight;
      LForm.Caption := FWindowTitle;

      // check for WebView2Runtime
      if not IsWebView2RuntimeInstalled() then Exit;

      LForm.Load();
      if not TFile.Exists(LFilename) then
        wvbCenterForm(LForm)
      else
        wvbLoadFormState(LForm, LIni);
      LForm.ShowModal;
    finally
      wvbSaveFormState(LForm, LIni);
      LForm.Free();
    end;
  finally
    LIni.UpdateFile();
    LIni.Free();
  end;
end;

function TWebViewBundle.AddPayload(const AExeFilename, AZipFilename: string): Boolean;
var
  LZipFileStream: TFileStream;
  LStringList: TStringList;
  LMemStream: TMemoryStream;
begin
  Result := False;

  if not TFile.Exists(AExeFilename) then Exit;
  if not TFile.Exists(AZipFilename) then Exit;

  if not wvbIsValidWin64PE(AExeFilename) then Exit;
  if not TZipFile.IsValid(AZipFilename) then Exit;

  LZipFileStream := TFile.OpenRead(AZipFilename);
  try
    wvbAddFileAsResource(AExeFilename, CwvbWebAppResName, AZipFilename)

  finally
    LZipFileStream.Free();
  end;

  LStringList := TStringList.Create();
  try
    LStringList.AddPair('WindowTitle', FWindowTitle);
    LStringList.AddPair('WindowWidth', FWindowWidth.ToString);
    LStringList.AddPair('WindowHeight', FWindowHeight.ToString);
    LStringList.AddPair('WindowResizeable', FWindowResizeable.ToString);
    LStringList.AddPair('IndexPage', FIndexPage);
    LStringList.AddPair('UserDataFolder', FUserDataFolder);

    LMemStream := TMemoryStream.Create();
    try
      LStringList.SaveToStream(LMemStream);
      wvbAddDataAsResource(AExeFilename, CwvbConfigResName, LMemStream.Memory, LMemStream.Size);
      Result := True;
    finally
      LMemStream.Free();
    end;

  finally
    LStringList.Free();
  end;

end;

function TWebViewBundle.BundleHtml(const AZipFilename, AOutputExeFilename: string): Boolean;
begin
  Result := False;

  wbvCopyFile(ParamStr(0),AOutputExeFilename);

  if not wvbIsValidWin64PE(AOutputExeFilename) then Exit;

  Result := AddPayload(AOutputExeFilename, AZipFilename);

  if Result then
  begin
    UpdateManifest(AOutputExeFilename);
    UpdatePayloadIcon(AOutputExeFilename);
    if FAddVersionInfo then
    begin
      UpdatePayloadVersionInfo(AOutputExeFilename);
    end;
  end;
end;

function TWebViewBundle.ZipHtml(const AFolderPath, ADestZipFilename: string; const AProgressCallback: TwvbZipProgressCallback; const AUserData: Pointer): Boolean;
begin
  Result := wvbZipFolder(AFolderPath, ADestZipFilename, AProgressCallback, AUserData);
end;

initialization
  ReportMemoryLeaksOnShutdown := True;
  SetExceptionMask(GetExceptionMask + [exOverflow, exInvalidOp]);
  SetConsoleCP(CP_UTF8);
  SetConsoleOutputCP(CP_UTF8);
  Randomize();
  TwvbCmdLine.Reset();

finalization

end.
