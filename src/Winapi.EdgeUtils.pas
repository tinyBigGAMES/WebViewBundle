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

-------------------------------------------------------------------------------
This is a custom implementation of Winapi.EdgeUtils designed to support loading
WebView2Loader.dll directly from memory. It intentionally uses the same unit
name, Winapi.EdgeUtils, to override the original version during local
compilation. This ensures that your application uses this customized logic
instead of the default implementation provided by Delphi’s RTL.
===============================================================================}

unit Winapi.EdgeUtils;

interface

uses
  Winapi.Windows,
  Winapi.ActiveX,
  Winapi.WebView2,
  System.SysUtils,
  System.Classes,
  System.Win.Registry,
  Dlluminator;

type
  TCbStdProc1<T1>       = reference to function(const P1: T1): HResult stdcall;
  TCbStdProc2<T1, T2>   = reference to function(const P1: T1; const P2: T2): HResult stdcall;
  TCbStdProc3<T1, T2>   = reference to function(P1: T1; P2: T2): HResult stdcall;
  TCbStdMethod1<T1>     = function(P1: T1): HResult of object stdcall;
  TCbStdMethod2<T1, T2> = function(P1: T1; const P2: T2): HResult of object stdcall;

  // Helper types used to work with WinRT event interfaces
  Callback<T1, T2> = record
    type
      TStdProc1 = TCbStdProc1<T1>;
      TStdProc2 = TCbStdProc2<T1, T2>;
      TStdProc3 = TCbStdProc3<T1, T2>;
      TStdMethod1 = TCbStdMethod1<T1>;
      TStdMethod2 = TCbStdMethod2<T1, T2>;
    class function CreateAs<INTF>(P: TStdProc1): INTF; overload; static;
    class function CreateAs<INTF>(P: TStdProc2): INTF; overload; static;
    class function CreateAs<INTF>(P: TStdProc3): INTF; overload; static;
    class function CreateAs<INTF>(P: TStdMethod1): INTF; overload; static;
    class function CreateAs<INTF>(P: TStdMethod2): INTF; overload; static;
  end;

  // Microsoft's default implementation of ICoreWebView2EnvironmentOptions et al is in WebView2EnvironmentOptions.h
  TCoreWebView2EnvironmentOptions = class(TInterfacedObject, ICoreWebView2EnvironmentOptions)
  private
    FAdditionalBrowserArguments: string;
    FLanguage: string;
    FTargetCompatibleBrowserVersion: string;
    FAllowSingleSignOnUsingOSPrimaryAccount: BOOL;
    class function AllocCOMString(const ADelphiString: string): PChar; static;
    class function BOOLToInt(AValue: BOOL): Integer; inline; static;
  public
    // ICoreWebView2EnvironmentOptions
    function Get_AdditionalBrowserArguments(out AValue: PWideChar): HResult; stdcall;
    function Set_AdditionalBrowserArguments(AValue: PWideChar): HResult; stdcall;
    function Get_Language(out AValue: PWideChar): HResult; stdcall;
    function Set_Language(AValue: PWideChar): HResult; stdcall;
    function Get_TargetCompatibleBrowserVersion(out AValue: PWideChar): HResult; stdcall;
    function Set_TargetCompatibleBrowserVersion(AValue: PWideChar): HResult; stdcall;
    function Get_AllowSingleSignOnUsingOSPrimaryAccount(out AAllow: Integer): HResult; stdcall;
    function Set_AllowSingleSignOnUsingOSPrimaryAccount(AAllow: Integer): HResult; stdcall;
  end;

function IsEdgeSupported(): Boolean;
function WebView2RuntimeInstalled(): Boolean;

// WebView2 loader DLL
function CreateCoreWebView2EnvironmentWithOptions(
  const ABrowserExecutableFolder, AUserDataFolder: LPCWSTR;
  const AEnvironmentOptions: ICoreWebView2EnvironmentOptions;
  const AEnvironmentCreatedHandler: ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler): HRESULT; stdcall;

function CreateCoreWebView2Environment(
  const AEnvironmentCreatedHandler: ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler): HRESULT; stdcall;

function GetCoreWebView2BrowserVersionString(
  const ABrowserExecutableFolder: LPCWSTR;
  var AVersionInfo: LPWSTR): HRESULT; stdcall;

function CompareBrowserVersions(const AVersion1, AVersion2: LPCWSTR;
  var AResult: Integer): HRESULT; stdcall;

implementation

var
 LCreateCoreWebView2EnvironmentWithOptions: function(
  const ABrowserExecutableFolder, AUserDataFolder: LPCWSTR;
  const AEnvironmentOptions: ICoreWebView2EnvironmentOptions;
  const AEnvironmentCreatedHandler: ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler): HRESULT; stdcall;

 LCreateCoreWebView2Environment: function(
  const AEnvironmentCreatedHandler: ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler): HRESULT; stdcall;

 LGetCoreWebView2BrowserVersionString: function(
  const ABrowserExecutableFolder: LPCWSTR;
  var AVersionInfo: LPWSTR): HRESULT; stdcall;

 LCompareBrowserVersions: function(const AVersion1, AVersion2: LPCWSTR;
  var AResult: Integer): HRESULT; stdcall;

{ Callback<T1, T2> }
class function Callback<T1, T2>.CreateAs<INTF>(P: TStdProc1): INTF;
type
  PIntf = ^INTF;
begin
  Result := PIntf(@P)^;
end;

class function Callback<T1, T2>.CreateAs<INTF>(P: TStdProc2): INTF;
type
  PIntf = ^INTF;
begin
  Result := PIntf(@P)^;
end;

class function Callback<T1, T2>.CreateAs<INTF>(P: TStdProc3): INTF;
type
  PIntf = ^INTF;
begin
  Result := PIntf(@P)^;
end;

class function Callback<T1, T2>.CreateAs<INTF>(P: TStdMethod1): INTF;
begin
  Result := CreateAs<INTF>(
    function(const P1: T1): HResult stdcall
    begin
      Result := P(P1)
    end);
end;

class function Callback<T1, T2>.CreateAs<INTF>(P: TStdMethod2): INTF;
begin
  Result := CreateAs<INTF>(
    function(const P1: T1; const P2: T2): HResult stdcall
    begin
      Result := P(P1, P2)
    end);
end;

{ Routines }

function WebView2RuntimeInstalled: Boolean;
var
  Reg: TRegistry;
  Paths: array of string;
  RootKeys: array of HKEY;
  i, j, k: Integer;  // Separate variables for each loop
  RegAccesses: array of Cardinal;
begin
  Result := False;

  // Define possible registry paths
  SetLength(Paths, 4);
  Paths[0] := 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'; // Runtime
  Paths[1] := 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F1E7E0FA-42C7-4BE1-B1E1-5DA547D5FA4C}'; // Old key
  Paths[2] := 'SOFTWARE\Microsoft\EdgeWebView\Installations';
  Paths[3] := 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedgewebview2.exe';

  // Define root keys to check
  SetLength(RootKeys, 2);
  RootKeys[0] := HKEY_LOCAL_MACHINE;
  RootKeys[1] := HKEY_CURRENT_USER;

  // Define registry access modes
  SetLength(RegAccesses, 2);
  RegAccesses[0] := KEY_READ;
  RegAccesses[1] := KEY_READ or KEY_WOW64_32KEY;

  Reg := TRegistry.Create;
  try
    for i := 0 to High(RootKeys) do
    begin
      Reg.RootKey := RootKeys[i];

      for j := 0 to High(RegAccesses) do
      begin
        Reg.Access := RegAccesses[j];

        for k := 0 to High(Paths) do  // Using k instead of reusing i
        begin
          if Reg.OpenKeyReadOnly(Paths[k]) then
          begin
            // For Installations, check if there are subkeys
            if (k = 2) and Reg.HasSubKeys then
              Result := True
            // For others, check if there's a value
            else if Reg.ValueExists('pv') or Reg.ValueExists('location') then
              Result := True;

            Reg.CloseKey;

            if Result then
              Exit;
          end;
        end;
      end;
    end;
  finally
    Reg.Free;
  end;

  // If registry detection fails, check for common file paths
  if not Result then
  begin
    Result := DirectoryExists('C:\Program Files (x86)\Microsoft\EdgeWebView\Application') or
              DirectoryExists('C:\Program Files\Microsoft\EdgeWebView\Application') or
              FileExists(ExtractFilePath(ParamStr(0)) + 'WebView2Loader.dll');
  end;
end;

function CreateCoreWebView2EnvironmentWithOptions(const ABrowserExecutableFolder, AUserDataFolder: LPCWSTR; const AEnvironmentOptions: ICoreWebView2EnvironmentOptions; const AEnvironmentCreatedHandler: ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler): HRESULT; stdcall;
begin
  if Assigned(LCreateCoreWebView2EnvironmentWithOptions) then
    Result := LCreateCoreWebView2EnvironmentWithOptions(
      ABrowserExecutableFolder, AUserDataFolder, AEnvironmentOptions, AEnvironmentCreatedHandler)
  else
    Result := E_FAIL;
end;

function CreateCoreWebView2Environment(const AEnvironmentCreatedHandler: ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler): HRESULT; stdcall;
begin
  if Assigned(LCreateCoreWebView2Environment) then
    Result := LCreateCoreWebView2Environment(AEnvironmentCreatedHandler)
  else
    Result := E_FAIL;
end;

function GetCoreWebView2BrowserVersionString(const ABrowserExecutableFolder: LPCWSTR; var AVersionInfo: LPWSTR): HRESULT; stdcall;
begin
  if Assigned(LGetCoreWebView2BrowserVersionString) then
    Result := LGetCoreWebView2BrowserVersionString(ABrowserExecutableFolder, AVersionInfo)
  else
    Result := E_FAIL;
end;

function CompareBrowserVersions(const AVersion1, AVersion2: LPCWSTR; var AResult: Integer): HRESULT; stdcall;
begin
  if Assigned(LCompareBrowserVersions) then
    Result := LCompareBrowserVersions(AVersion1, AVersion2, AResult)
  else
    Result := E_FAIL;
end;

function IsEdgeSupported: Boolean;
begin
  // Microsoft Edge browser only works on Windows 7 and above
  Result := TOSVersion.Check(6, 1);
end;

{ TCoreWebView2EnvironmentOptions }
class function TCoreWebView2EnvironmentOptions.BOOLToInt(AValue: BOOL): Integer;
begin
  // BOOL aka LongBool uses 0 for False and non-zero for True.
  // Specifically Delphi will assign the value of -1 (VARIANT_TRUE in C) for True when assigning to LongBool.
  // Note that WebView2 likes a BOOL value of TRUE to be 1, so we must take care not to cast BOOL to Integer.
  // See also: https://devblogs.microsoft.com/oldnewthing/20041222-00/?p=36923
  if AValue then
    Result := 1
  else
    Result := 0;
end;

class function TCoreWebView2EnvironmentOptions.AllocCOMString(const ADelphiString: string): PChar;
var
  DelphiStringLength: Integer;
begin
  // In this class the caller (WebView2) will free the allocated memory
  DelphiStringLength := Succ(ADelphiString.Length);
  Result := CoTaskMemAlloc(DelphiStringLength * SizeOf(Char));
  StringToWideChar(ADelphiString, Result, DelphiStringLength);
end;

function TCoreWebView2EnvironmentOptions.Get_AdditionalBrowserArguments(out AValue: PWideChar): HResult;
begin
  try
    AValue := AllocCOMString(FAdditionalBrowserArguments);
    Result := S_OK
  except
    Result := E_FAIL
  end;
end;

function TCoreWebView2EnvironmentOptions.Get_Language(out AValue: PWideChar): HResult;
begin
  try
    AValue := AllocCOMString(FLanguage);
    Result := S_OK
  except
    Result := E_FAIL
  end;
end;

function TCoreWebView2EnvironmentOptions.Get_TargetCompatibleBrowserVersion(out AValue: PWideChar): HResult;
begin
  try
    AValue := AllocCOMString(FTargetCompatibleBrowserVersion);
    Result := S_OK
  except
    Result := E_FAIL
  end;
end;

function TCoreWebView2EnvironmentOptions.Get_AllowSingleSignOnUsingOSPrimaryAccount(out AAllow: Integer): HResult;
begin
  try
    AAllow := BOOLToInt(FAllowSingleSignOnUsingOSPrimaryAccount);
    Result := S_OK
  except
    Result := E_FAIL
  end;
end;

function TCoreWebView2EnvironmentOptions.Set_AdditionalBrowserArguments(AValue: PWideChar): HResult;
begin
  try
    FAdditionalBrowserArguments := string(AValue);
    Result := S_OK
  except
    Result := E_FAIL
  end;
end;

function TCoreWebView2EnvironmentOptions.Set_Language(AValue: PWideChar): HResult;
begin
  try
    FLanguage := string(AValue);
    Result := S_OK
  except
    Result := E_FAIL
  end;
end;

function TCoreWebView2EnvironmentOptions.Set_TargetCompatibleBrowserVersion(AValue: PWideChar): HResult;
begin
  try
    FTargetCompatibleBrowserVersion := string(AValue);
    Result := S_OK
  except
    Result := E_FAIL
  end;
end;

function TCoreWebView2EnvironmentOptions.Set_AllowSingleSignOnUsingOSPrimaryAccount(AAllow: Integer): HResult;
begin
  try
    FAllowSingleSignOnUsingOSPrimaryAccount := BOOL(AAllow);
    Result := S_OK
  except
    Result := E_FAIL
  end;
end;

//===========================================================================

{$R WebViewBundle.Deps.res}

var
  CDepsDLLHandle: THandle = 0;
  LError: string = '';

function LoadDepsDLL(out AError: string): Boolean;
var
  LResStream: TResourceStream;

  function df967db2378b4df5b748f7c2c52f4963(): string;
  const
    CValue = '6f0ec63bcd614c3aac7568a12a9b5040';
  begin
    Result := CValue;
  end;

  procedure SetError(const AText: string; const AArgs: array of const);
  begin
    AError := Format(AText, AArgs);
  end;

begin
  Result := False;
  AError := '';

  // load deps DLL
  if CDepsDLLHandle <> 0 then Exit;
  try
    if not Boolean((FindResource(HInstance, PWideChar(df967db2378b4df5b748f7c2c52f4963()), RT_RCDATA) <> 0)) then
    begin
      SetError('Failed to find CLibs DLL resource', []);
      Exit;
    end;

    LResStream := TResourceStream.Create(HInstance, df967db2378b4df5b748f7c2c52f4963(), RT_RCDATA);
    try
      CDepsDLLHandle := Dlluminator.LoadLibrary(LResStream.Memory, LResStream.Size);

      if CDepsDLLHandle = 0 then
      begin
        SetError('Failed to load extracted CLibs DLL', []);
        Exit;
      end;

      LCreateCoreWebView2EnvironmentWithOptions := GetProcAddress(CDepsDLLHandle, 'CreateCoreWebView2EnvironmentWithOptions');
      LCreateCoreWebView2Environment            := GetProcAddress(CDepsDLLHandle, 'CreateCoreWebView2Environment');
      LGetCoreWebView2BrowserVersionString      := GetProcAddress(CDepsDLLHandle, 'GetCoreWebView2BrowserVersionString');
      LCompareBrowserVersions                   := GetProcAddress(CDepsDLLHandle, 'CompareBrowserVersions');

      Result := True;
    finally
      LResStream.Free();
    end;

  except
    on E: Exception do
      SetError('Unexpected error: %s', [E.Message]);
  end;
end;

procedure UnloadDepsDLL();
begin
  // unload deps DLL
  if CDepsDLLHandle <> 0 then
  begin
    FreeLibrary(CDepsDLLHandle);
    CDepsDLLHandle := 0;
  end;
end;

initialization
  FSetExceptMask(femALLEXCEPT);

  if not LoadDepsDLL(LError) then
  begin
    MessageBox(0, PWideChar(LError), 'Critical Initialization Error', MB_ICONERROR);
    Halt(1); // Exit the application with a non-zero exit code to indicate failure
  end;

finalization

  UnloadDepsDLL();

end.
