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

unit WebViewBundle.Utils;

{$I WebViewBundle.Defines.inc}

interface

uses
  WinApi.Windows,
  Winapi.TlHelp32,
  System.Types,
  System.UITypes,
  System.Generics.Collections,
  System.SysUtils,
  System.IOUtils,
  System.Classes,
  System.Math,
  System.Zip,
  System.IniFiles,
  VCL.Forms,
  VCL.Dialogs;

type
  { TwvbZipProgressCallback }
  TwvbZipProgressCallback = reference to procedure(const AFilename: string; const UserData: Pointer);

  { TwvbCaptureConsoleCallback }
  TwvbCaptureConsoleCallback = procedure(const ALine: string; const AUserData: Pointer);

  { TwvbConfirmDialogResult }
  TwvbConfirmDialogResult = (cdYes, cdNo, cdCancel);

  { TwvbShowMessage }
  TwvbShowMessage = (smDefault = $00000000, smError = $00000010, smWarning = $00000030, smInformation = $00000040);

{ TwvbDirectoryStack }
  TwvbDirectoryStack = class
  protected
    FStack: TStack<String>;
  public
    constructor Create(); virtual;
    destructor Destroy; override;
    procedure Push(aPath: string);
    procedure PushFilePath(aFilename: string);
    procedure Pop;
  end;

  { TwvbCmdLine }
  TwvbCmdLine = record
  private
    class var
      FCmdLine: string;
    class function GetCmdLine: PChar; static;
    class function GetParamStr(aParamStr: PChar; var aParam: string): PChar; static;
    //class operator Initialize (out ADest: TCmdLine);
  public
    class function ParamCount: Integer; static;
    class procedure Reset; static;
    class procedure ClearParams; static;
    class procedure AddAParam(const aParam: string); static;
    class procedure AddParams(const aParams: string); static;
    class function ParamStr(aIndex: Integer): string; static;
    class function GetParamValue(const aParamName: string; aSwitchChars: TSysCharSet; aSeperator: Char; var aValue: string): Boolean; overload; static;
    class function GetParamValue(const aParamName: string; var aValue: string): Boolean; overload; static;
    class function GetParam(const aParamName: string): Boolean; static;
  end;

  { TwvbPayloadStream }
  TwvbPayloadStream = class(TStream)
  private const
    cWaterMarkGUID: TGUID = '{9FABA105-EDA8-45C3-89F4-369315A947EB}';
  private type
    EPayloadStream = class(Exception);

    TPayloadStreamFooter = packed record
      WaterMark: TGUID;
      ExeSize: Int64;
      DataSize: Int64;
    end;

    TPayloadStreamOpenMode = (
      pomRead,    // read mode
      pomWrite    // write (append) mode
    );
  private
    fMode: TPayloadStreamOpenMode;  // stream open mode
    fFileStream: TFileStream;       // file stream for payload
    fDataStart: Int64;              // start of payload data in file
    fDataSize: Int64;               // size of payload
    function GetPosition: Int64;
    procedure SetPosition(const Value: Int64);
    procedure InitFooter(out Footer: TPayloadStreamFooter);
    function ReadFooter(const FileStream: TFileStream; out Footer: TPayloadStreamFooter): Boolean;
  public
    property CurrentPos: Int64 read GetPosition write SetPosition;
    constructor Create(const FileName: string; const Mode: TPayloadStreamOpenMode);
    destructor Destroy; override;
    function Seek(const Offset: Int64; Origin: TSeekOrigin): Int64; override;
    procedure SetSize(const NewSize: Int64); override;
    function Read(var Buffer; Count: LongInt): LongInt; override;
    function Write(const Buffer; Count: LongInt): LongInt; override;
    property DataSize: Int64 read fDataSize;
  end;

{ Routines }
function  wvbRemoveQuotes(const AText: string): string;
procedure wvbFreeNilObject(const [ref] AObject: TObject);
function  wvbGetEXEPath(): string;
function  wvbGetExeBasePath(const aFilename: string): string;
function  wvbHasConsoleOutput(): Boolean;
function  wvbEmptyFolder(const AFolder: string): Boolean;
function  wvbExpandRelFilename(aBaseFilename, aRelFilename: string): string;
procedure wvbCaptureConsoleOutput(const ATitle: string; const ACommand: PChar; const AParameters: PChar; var AExitCode: DWORD; ACallback: TwvbCaptureConsoleCallback; AUserData: Pointer);
function  wvbEnableVirtualTerminalProcessing(): DWORD;
function  wvbIsValidWin64PE(const AFilePath: string): Boolean;
procedure wvbUpdateIconResource(const AExeFilePath, AIconFilePath: string);
procedure wvbUpdateVersionInfoResource(const PEFilePath: string; const AMajor, AMinor, APatch: Word; const AProductName, ADescription, AFilename, ACompanyName, ACopyright: string);
function  wvbResourceExist(const AResName: string): Boolean;
procedure wvbLoadStringListFromResource(const aResName: string; aList: TStringList);
function  wvbLoadStringFromResource(const aResName: string): string;
function  wvbAddResManifestFromResource(const aResName: string; const aModuleFile: string; aLanguage: Integer=1033): Boolean;
function  wvbGetParentProcessName(): string;
function  wvbWasStartedFromCommandLine(): Boolean;
function  wvbHasEnoughDiskSpace(const AFilePath: string; ARequiredSize: Int64): Boolean;
procedure wvbSaveFormState(Form: TForm; IniFile: TIniFile; const SectionName: string = '');
function  wvbLoadFormState(Form: TForm; IniFile: TIniFile; const SectionName: string = ''): Boolean;
function  wvbZipFolder(const AFolderPath, ADestZipFilename: string; const AProgressCallback: TwvbZipProgressCallback; const AUserData: Pointer): Boolean;
function  wbvCopyFile(const ASrcFile, ADestFile: string): Boolean;
function  wvbAddFileAsResource(const AExeFilename, AResourceName, ADataFileName: string): Boolean;
function  wvbAddDataAsResource(const AExeFilename, AResourceName: string; const AData: Pointer; const ADataSize: Integer): Boolean;
function  wvbAddStringAsResource(const AExeFilename, AResourceName, AStringData: string): Boolean;
procedure wvbCenterForm(const AForm: TForm); overload;
procedure wvbCenterForm(const AForm, AMainForm: TForm); overload;
function  wvbConfirmDlg(const AMsg: string; const AArgs: array of const): TwvbConfirmDialogResult;
function  wvbCreateDirsInPath(const AFilename: string): Boolean;
function  wvbExtractResToFile(const AResName, AFilename: string): Boolean;
function  wvbRunExe(const AFilename, ACmdLine, AStartPath: string): Boolean;
function  wvbExtractAndRunResEXE(const AResName, ACmdLine, AStartPath: string): Boolean;
procedure wvbShowMessage(const ATitle, AText: string; const AArgs: array of const; AType: TwvbShowMessage);


implementation

{ Routines }
function  wvbRemoveQuotes(const AText: string): string;
var
  S: string;
begin
  S := AnsiDequotedStr(aText, '"');
  Result := AnsiDequotedStr(S, '''');
end;

procedure wvbFreeNilObject(const [ref] AObject: TObject);
var
  Temp: TObject;
begin
  if not Assigned(AObject) then Exit;
  Temp := AObject;
  TObject(Pointer(@AObject)^) := nil;
  Temp.Free;
end;

function  wvbGetEXEPath(): string;
begin
  Result := TPath.GetDirectoryName(ParamStr(0));
end;

function  wvbGetExeBasePath(const aFilename: string): string;
begin
  Result := TPath.Combine(wvbGetEXEPath(), aFilename);
end;

function  wvbHasConsoleOutput(): Boolean;
var
  LStdOut: THandle;
  LMode: DWORD;
begin
  LStdOut := GetStdHandle(STD_OUTPUT_HANDLE);
  Result := (LStdOut <> INVALID_HANDLE_VALUE) and GetConsoleMode(LStdOut, LMode);
end;

function wvbEmptyFolder(const AFolder: string): Boolean;
var
  LFiles, LFolders: TStringDynArray;
  LFileOrDir: string;
begin
  Result := True;

  if not TDirectory.Exists(AFolder) then
  begin
    Exit(False);
  end;

  try
    LFiles := TDirectory.GetFiles(AFolder, '*', TSearchOption.soAllDirectories);
    for LFileOrDir in LFiles do
    begin
      try
        TFile.Delete(LFileOrDir);
      except
        on E: Exception do
        begin
          Exit(False);
        end;
      end;
    end;

    // Delete subfolders
    LFolders := TDirectory.GetDirectories(AFolder, '*', TSearchOption.soAllDirectories);
    for LFileOrDir in LFolders do
    begin
      try
        TDirectory.Delete(LFileOrDir, True);
      except
        on E: Exception do
        begin
          Exit(False);
        end;
      end;
    end;
  except
    on E: Exception do
    begin
      Exit(False);
    end;
  end;
end;

function PathCombine(lpszDest: PWideChar; const lpszDir, lpszFile: PWideChar): PWideChar; stdcall; external 'shlwapi.dll' name 'PathCombineW';

function wvbExpandRelFilename(aBaseFilename, aRelFilename: string): string;
var
  buff: array [0 .. MAX_PATH + 1] of WideChar;
begin
  PathCombine(@buff[0], PWideChar(ExtractFilePath(aBaseFilename)),
    PWideChar(aRelFilename));
  Result := string(buff);
end;

procedure ProcessMessages();
var
  LMsg: TMsg;
begin
  while Integer(PeekMessage(LMsg, 0, 0, 0, PM_REMOVE)) <> 0 do
  begin
    TranslateMessage(LMsg);
    DispatchMessage(LMsg);
  end;
end;

procedure wvbCaptureConsoleOutput(const ATitle: string; const ACommand: PChar; const AParameters: PChar; var AExitCode: DWORD; ACallback: TwvbCaptureConsoleCallback; AUserData: Pointer);
const
  //CReadBuffer = 2400;
  CReadBuffer = 1024*2;
var
  saSecurity: TSecurityAttributes;
  hRead: THandle;
  hWrite: THandle;
  suiStartup: TStartupInfo;
  piProcess: TProcessInformation;
  pBuffer: array [0 .. CReadBuffer] of AnsiChar;
  dBuffer: array [0 .. CReadBuffer] of AnsiChar;
  dRead: DWORD;
  dRunning: DWORD;
  dAvailable: DWORD;
  CmdLine: string;
  BufferList: TStringList;
  Line: string;
  LExitCode: DWORD;
begin
  saSecurity.nLength := SizeOf(TSecurityAttributes);
  saSecurity.bInheritHandle := true;
  saSecurity.lpSecurityDescriptor := nil;
  if CreatePipe(hRead, hWrite, @saSecurity, 0) then
    try
      FillChar(suiStartup, SizeOf(TStartupInfo), #0);
      suiStartup.cb := SizeOf(TStartupInfo);
      suiStartup.hStdInput := hRead;
      suiStartup.hStdOutput := hWrite;
      suiStartup.hStdError := hWrite;
      suiStartup.dwFlags := STARTF_USESTDHANDLES or STARTF_USESHOWWINDOW;
      suiStartup.wShowWindow := SW_HIDE;
      if ATitle.IsEmpty then
        suiStartup.lpTitle := nil
      else
        suiStartup.lpTitle := PChar(ATitle);
      CmdLine := ACommand + ' ' + AParameters;
      if CreateProcess(nil, PChar(CmdLine), @saSecurity, @saSecurity, true, NORMAL_PRIORITY_CLASS, nil, nil, suiStartup, piProcess) then
        try
          BufferList := TStringList.Create;
          try
            repeat
              dRunning := WaitForSingleObject(piProcess.hProcess, 100);
              PeekNamedPipe(hRead, nil, 0, nil, @dAvailable, nil);
              if (dAvailable > 0) then
                repeat
                  dRead := 0;
                  ReadFile(hRead, pBuffer[0], CReadBuffer, dRead, nil);
                  pBuffer[dRead] := #0;
                  OemToCharA(pBuffer, dBuffer);
                  BufferList.Clear;
                  BufferList.Text := string(pBuffer);
                  for line in BufferList do
                  begin
                    if Assigned(ACallback) then
                    begin
                      ACallback(line, AUserData);
                    end;
                  end;
                until (dRead < CReadBuffer);
              ProcessMessages;
            until (dRunning <> WAIT_TIMEOUT);

            if GetExitCodeProcess(piProcess.hProcess, LExitCode) then
            begin
              AExitCode := LExitCode;
            end;

          finally
            FreeAndNil(BufferList);
          end;
        finally
          CloseHandle(piProcess.hProcess);
          CloseHandle(piProcess.hThread);
        end;
    finally
      CloseHandle(hRead);
      CloseHandle(hWrite);
    end;
end;

function wvbEnableVirtualTerminalProcessing(): DWORD;
var
  HOut: THandle;
  LMode: DWORD;
begin
  HOut := GetStdHandle(STD_OUTPUT_HANDLE);
  if HOut = INVALID_HANDLE_VALUE then
  begin
    Result := GetLastError;
    Exit;
  end;

  if not GetConsoleMode(HOut, LMode) then
  begin
    Result := GetLastError;
    Exit;
  end;

  LMode := LMode or ENABLE_VIRTUAL_TERMINAL_PROCESSING;
  if not SetConsoleMode(HOut, LMode) then
  begin
    Result := GetLastError;
    Exit;
  end;

  Result := 0;  // Success
end;

function wvbIsValidWin64PE(const AFilePath: string): Boolean;
var
  LFile: TFileStream;
  LDosHeader: TImageDosHeader;
  LPEHeaderOffset: DWORD;
  LPEHeaderSignature: DWORD;
  LFileHeader: TImageFileHeader;
begin
  Result := False;

  if not FileExists(AFilePath) then
    Exit;

  LFile := TFileStream.Create(AFilePath, fmOpenRead or fmShareDenyWrite);
  try
    // Check if file is large enough for DOS header
    if LFile.Size < SizeOf(TImageDosHeader) then
      Exit;

    // Read DOS header
    LFile.ReadBuffer(LDosHeader, SizeOf(TImageDosHeader));

    // Check DOS signature
    if LDosHeader.e_magic <> IMAGE_DOS_SIGNATURE then // 'MZ'
      Exit;

      // Validate PE header offset
    LPEHeaderOffset := LDosHeader._lfanew;
    if LFile.Size < LPEHeaderOffset + SizeOf(DWORD) + SizeOf(TImageFileHeader) then
      Exit;

    // Seek to the PE header
    LFile.Position := LPEHeaderOffset;

    // Read and validate the PE signature
    LFile.ReadBuffer(LPEHeaderSignature, SizeOf(DWORD));
    if LPEHeaderSignature <> IMAGE_NT_SIGNATURE then // 'PE\0\0'
      Exit;

   // Read the file header
    LFile.ReadBuffer(LFileHeader, SizeOf(TImageFileHeader));

    // Check if it is a 64-bit executable
    if LFileHeader.Machine <> IMAGE_FILE_MACHINE_AMD64 then   Exit;

    // If all checks pass, it's a valid Win64 PE file
    Result := True;
  finally
    LFile.Free;
  end;
end;

procedure wvbUpdateIconResource(const AExeFilePath, AIconFilePath: string);
type
  TIconDir = packed record
    idReserved: Word;  // Reserved, must be 0
    idType: Word;      // Resource type, 1 for icons
    idCount: Word;     // Number of images in the file
  end;
  PIconDir = ^TIconDir;

  TGroupIconDirEntry = packed record
    bWidth: Byte;            // Width of the icon (0 means 256)
    bHeight: Byte;           // Height of the icon (0 means 256)
    bColorCount: Byte;       // Number of colors in the palette (0 if more than 256)
    bReserved: Byte;         // Reserved, must be 0
    wPlanes: Word;           // Color planes
    wBitCount: Word;         // Bits per pixel
    dwBytesInRes: Cardinal;  // Size of the image data
    nID: Word;               // Resource ID of the icon
  end;

  TGroupIconDir = packed record
    idReserved: Word;  // Reserved, must be 0
    idType: Word;      // Resource type, 1 for icons
    idCount: Word;     // Number of images in the file
    Entries: array[0..0] of TGroupIconDirEntry; // Variable-length array
  end;

  TIconResInfo = packed record
    bWidth: Byte;            // Width of the icon (0 means 256)
    bHeight: Byte;           // Height of the icon (0 means 256)
    bColorCount: Byte;       // Number of colors in the palette (0 if more than 256)
    bReserved: Byte;         // Reserved, must be 0
    wPlanes: Word;           // Color planes (should be 1)
    wBitCount: Word;         // Bits per pixel
    dwBytesInRes: Cardinal;  // Size of the image data
    dwImageOffset: Cardinal; // Offset of the image data in the file
  end;
  PIconResInfo = ^TIconResInfo;

var
  LUpdateHandle: THandle;
  LIconStream: TMemoryStream;
  LIconDir: PIconDir;
  LIconGroup: TMemoryStream;
  LIconRes: PByte;
  LIconID: Word;
  I: Integer;
  LGroupEntry: TGroupIconDirEntry;
begin

  if not FileExists(AExeFilePath) then
    raise Exception.Create('The specified executable file does not exist.');

  if not FileExists(AIconFilePath) then
    raise Exception.Create('The specified icon file does not exist.');

  LIconStream := TMemoryStream.Create;
  LIconGroup := TMemoryStream.Create;
  try
    // Load the icon file
    LIconStream.LoadFromFile(AIconFilePath);

    // Read the ICONDIR structure from the icon file
    LIconDir := PIconDir(LIconStream.Memory);
    if LIconDir^.idReserved <> 0 then
      raise Exception.Create('Invalid icon file format.');

    // Begin updating the executable's resources
    LUpdateHandle := BeginUpdateResource(PChar(AExeFilePath), False);
    if LUpdateHandle = 0 then
      raise Exception.Create('Failed to begin resource update.');

    try
      // Process each icon image in the .ico file
      LIconRes := PByte(LIconStream.Memory) + SizeOf(TIconDir);
      for I := 0 to LIconDir^.idCount - 1 do
      begin
        // Assign a unique resource ID for the RT_ICON
        LIconID := I + 1;

        // Add the icon image data as an RT_ICON resource
        if not UpdateResource(LUpdateHandle, RT_ICON, PChar(LIconID), LANG_NEUTRAL,
          Pointer(PByte(LIconStream.Memory) + PIconResInfo(LIconRes)^.dwImageOffset),
          PIconResInfo(LIconRes)^.dwBytesInRes) then
          raise Exception.CreateFmt('Failed to add RT_ICON resource for image %d.', [I]);

        // Move to the next icon entry
        Inc(LIconRes, SizeOf(TIconResInfo));
      end;

      // Create the GROUP_ICON resource
      LIconGroup.Clear;
      LIconGroup.Write(LIconDir^, SizeOf(TIconDir)); // Write ICONDIR header

      LIconRes := PByte(LIconStream.Memory) + SizeOf(TIconDir);
      // Write each GROUP_ICON entry
      for I := 0 to LIconDir^.idCount - 1 do
      begin
        // Populate the GROUP_ICON entry
        LGroupEntry.bWidth := PIconResInfo(LIconRes)^.bWidth;
        LGroupEntry.bHeight := PIconResInfo(LIconRes)^.bHeight;
        LGroupEntry.bColorCount := PIconResInfo(LIconRes)^.bColorCount;
        LGroupEntry.bReserved := 0;
        LGroupEntry.wPlanes := PIconResInfo(LIconRes)^.wPlanes;
        LGroupEntry.wBitCount := PIconResInfo(LIconRes)^.wBitCount;
        LGroupEntry.dwBytesInRes := PIconResInfo(LIconRes)^.dwBytesInRes;
        LGroupEntry.nID := I + 1; // Match resource ID for RT_ICON

        // Write the populated GROUP_ICON entry to the stream
        LIconGroup.Write(LGroupEntry, SizeOf(TGroupIconDirEntry));

        // Move to the next ICONDIRENTRY
        Inc(LIconRes, SizeOf(TIconResInfo));
      end;

      // Add the GROUP_ICON resource to the executable
      if not UpdateResource(LUpdateHandle, RT_GROUP_ICON, 'MAINICON', LANG_NEUTRAL,
        LIconGroup.Memory, LIconGroup.Size) then
        raise Exception.Create('Failed to add RT_GROUP_ICON resource.');

      // Commit the resource updates
      if not EndUpdateResource(LUpdateHandle, False) then
        raise Exception.Create('Failed to commit resource updates.');
    except
      EndUpdateResource(LUpdateHandle, True); // Discard changes on failure
      raise;
    end;
  finally
    LIconStream.Free;
    LIconGroup.Free;
  end;
end;

procedure wvbUpdateVersionInfoResource(const PEFilePath: string; const AMajor, AMinor, APatch: Word; const AProductName, ADescription, AFilename, ACompanyName, ACopyright: string);
type
  { TVSFixedFileInfo }
  TVSFixedFileInfo = packed record
    dwSignature: DWORD;        // e.g. $FEEF04BD
    dwStrucVersion: DWORD;     // e.g. $00010000 for version 1.0
    dwFileVersionMS: DWORD;    // e.g. $00030075 for version 3.75
    dwFileVersionLS: DWORD;    // e.g. $00000031 for version 0.31
    dwProductVersionMS: DWORD; // Same format as dwFileVersionMS
    dwProductVersionLS: DWORD; // Same format as dwFileVersionLS
    dwFileFlagsMask: DWORD;    // = $3F for version "0011 1111"
    dwFileFlags: DWORD;        // e.g. VFF_DEBUG | VFF_PRERELEASE
    dwFileOS: DWORD;           // e.g. VOS_NT_WINDOWS32
    dwFileType: DWORD;         // e.g. VFT_APP
    dwFileSubtype: DWORD;      // e.g. VFT2_UNKNOWN
    dwFileDateMS: DWORD;       // file date
    dwFileDateLS: DWORD;       // file date
  end;

  { TStringPair }
  TStringPair = record
    Key: string;
    Value: string;
  end;

var
  LHandleUpdate: THandle;
  LVersionInfoStream: TMemoryStream;
  LFixedInfo: TVSFixedFileInfo;
  LDataPtr: Pointer;
  LDataSize: Integer;
  LStringFileInfoStart, LStringTableStart, LVarFileInfoStart: Int64;
  LStringPairs: array of TStringPair;
  LVErsion: string;
  LMajor, LMinor,LPatch: Word;
  LVSVersionInfoStart: Int64;
  LPair: TStringPair;
  LStringInfoEnd, LStringStart: Int64;
  LStringEnd, LFinalPos: Int64;
  LTranslationStart: Int64;

  procedure AlignStream(const AStream: TMemoryStream; const AAlignment: Integer);
  var
    LPadding: Integer;
    LPadByte: Byte;
  begin
    LPadding := (AAlignment - (AStream.Position mod AAlignment)) mod AAlignment;
    LPadByte := 0;
    while LPadding > 0 do
    begin
      AStream.WriteBuffer(LPadByte, 1);
      Dec(LPadding);
    end;
  end;

  procedure WriteWideString(const AStream: TMemoryStream; const AText: string);
  var
    LWideText: WideString;
  begin
    LWideText := WideString(AText);
    AStream.WriteBuffer(PWideChar(LWideText)^, (Length(LWideText) + 1) * SizeOf(WideChar));
  end;

  procedure SetFileVersionFromString(const AVersion: string; out AFileVersionMS, AFileVersionLS: DWORD);
  var
    LVersionParts: TArray<string>;
    LMajor, LMinor, LBuild, LRevision: Word;
  begin
    // Split the version string into its components
    LVersionParts := AVersion.Split(['.']);
    if Length(LVersionParts) <> 4 then
      raise Exception.Create('Invalid version string format. Expected "Major.Minor.Build.Revision".');

    // Parse each part into a Word
    LMajor := StrToIntDef(LVersionParts[0], 0);
    LMinor := StrToIntDef(LVersionParts[1], 0);
    LBuild := StrToIntDef(LVersionParts[2], 0);
    LRevision := StrToIntDef(LVersionParts[3], 0);

    // Set the high and low DWORD values
    AFileVersionMS := (DWORD(LMajor) shl 16) or DWORD(LMinor);
    AFileVersionLS := (DWORD(LBuild) shl 16) or DWORD(LRevision);
  end;

begin
  LMajor := EnsureRange(AMajor, 0, MaxWord);
  LMinor := EnsureRange(AMinor, 0, MaxWord);
  LPatch := EnsureRange(APatch, 0, MaxWord);
  LVersion := Format('%d.%d.%d.0', [LMajor, LMinor, LPatch]);

  SetLength(LStringPairs, 8);
  LStringPairs[0].Key := 'CompanyName';
  LStringPairs[0].Value := ACompanyName;
  LStringPairs[1].Key := 'FileDescription';
  LStringPairs[1].Value := ADescription;
  LStringPairs[2].Key := 'FileVersion';
  LStringPairs[2].Value := LVersion;
  LStringPairs[3].Key := 'InternalName';
  LStringPairs[3].Value := ADescription;
  LStringPairs[4].Key := 'LegalCopyright';
  LStringPairs[4].Value := ACopyright;
  LStringPairs[5].Key := 'OriginalFilename';
  LStringPairs[5].Value := AFilename;
  LStringPairs[6].Key := 'ProductName';
  LStringPairs[6].Value := AProductName;
  LStringPairs[7].Key := 'ProductVersion';
  LStringPairs[7].Value := LVersion;

  // Initialize fixed info structure
  FillChar(LFixedInfo, SizeOf(LFixedInfo), 0);
  LFixedInfo.dwSignature := $FEEF04BD;
  LFixedInfo.dwStrucVersion := $00010000;
  LFixedInfo.dwFileVersionMS := $00010000;
  LFixedInfo.dwFileVersionLS := $00000000;
  LFixedInfo.dwProductVersionMS := $00010000;
  LFixedInfo.dwProductVersionLS := $00000000;
  LFixedInfo.dwFileFlagsMask := $3F;
  LFixedInfo.dwFileFlags := 0;
  LFixedInfo.dwFileOS := VOS_NT_WINDOWS32;
  LFixedInfo.dwFileType := VFT_APP;
  LFixedInfo.dwFileSubtype := 0;
  LFixedInfo.dwFileDateMS := 0;
  LFixedInfo.dwFileDateLS := 0;

  // SEt MS and LS for FileVersion and ProductVersion
  SetFileVersionFromString(LVersion, LFixedInfo.dwFileVersionMS, LFixedInfo.dwFileVersionLS);
  SetFileVersionFromString(LVersion, LFixedInfo.dwProductVersionMS, LFixedInfo.dwProductVersionLS);

  LVersionInfoStream := TMemoryStream.Create;
  try
    // VS_VERSION_INFO
    LVSVersionInfoStart := LVersionInfoStream.Position;

    LVersionInfoStream.WriteData<Word>(0);  // Length placeholder
    LVersionInfoStream.WriteData<Word>(SizeOf(TVSFixedFileInfo));  // Value length
    LVersionInfoStream.WriteData<Word>(0);  // Type = 0
    WriteWideString(LVersionInfoStream, 'VS_VERSION_INFO');
    AlignStream(LVersionInfoStream, 4);

    // VS_FIXEDFILEINFO
    LVersionInfoStream.WriteBuffer(LFixedInfo, SizeOf(TVSFixedFileInfo));
    AlignStream(LVersionInfoStream, 4);

    // StringFileInfo
    LStringFileInfoStart := LVersionInfoStream.Position;
    LVersionInfoStream.WriteData<Word>(0);  // Length placeholder
    LVersionInfoStream.WriteData<Word>(0);  // Value length = 0
    LVersionInfoStream.WriteData<Word>(1);  // Type = 1
    WriteWideString(LVersionInfoStream, 'StringFileInfo');
    AlignStream(LVersionInfoStream, 4);

    // StringTable
    LStringTableStart := LVersionInfoStream.Position;
    LVersionInfoStream.WriteData<Word>(0);  // Length placeholder
    LVersionInfoStream.WriteData<Word>(0);  // Value length = 0
    LVersionInfoStream.WriteData<Word>(1);  // Type = 1
    WriteWideString(LVersionInfoStream, '040904B0'); // Match Delphi's default code page
    AlignStream(LVersionInfoStream, 4);

    // Write string pairs
    for LPair in LStringPairs do
    begin
      LStringStart := LVersionInfoStream.Position;

      LVersionInfoStream.WriteData<Word>(0);  // Length placeholder
      LVersionInfoStream.WriteData<Word>((Length(LPair.Value) + 1) * 2);  // Value length
      LVersionInfoStream.WriteData<Word>(1);  // Type = 1
      WriteWideString(LVersionInfoStream, LPair.Key);
      AlignStream(LVersionInfoStream, 4);
      WriteWideString(LVersionInfoStream, LPair.Value);
      AlignStream(LVersionInfoStream, 4);

      LStringEnd := LVersionInfoStream.Position;
      LVersionInfoStream.Position := LStringStart;
      LVersionInfoStream.WriteData<Word>(LStringEnd - LStringStart);
      LVersionInfoStream.Position := LStringEnd;
    end;

    LStringInfoEnd := LVersionInfoStream.Position;

    // Write StringTable length
    LVersionInfoStream.Position := LStringTableStart;
    LVersionInfoStream.WriteData<Word>(LStringInfoEnd - LStringTableStart);

    // Write StringFileInfo length
    LVersionInfoStream.Position := LStringFileInfoStart;
    LVersionInfoStream.WriteData<Word>(LStringInfoEnd - LStringFileInfoStart);

    // Start VarFileInfo where StringFileInfo ended
    LVarFileInfoStart := LStringInfoEnd;
    LVersionInfoStream.Position := LVarFileInfoStart;

    // VarFileInfo header
    LVersionInfoStream.WriteData<Word>(0);  // Length placeholder
    LVersionInfoStream.WriteData<Word>(0);  // Value length = 0
    LVersionInfoStream.WriteData<Word>(1);  // Type = 1 (text)
    WriteWideString(LVersionInfoStream, 'VarFileInfo');
    AlignStream(LVersionInfoStream, 4);

    // Translation value block
    LTranslationStart := LVersionInfoStream.Position;
    LVersionInfoStream.WriteData<Word>(0);  // Length placeholder
    LVersionInfoStream.WriteData<Word>(4);  // Value length = 4 (size of translation value)
    LVersionInfoStream.WriteData<Word>(0);  // Type = 0 (binary)
    WriteWideString(LVersionInfoStream, 'Translation');
    AlignStream(LVersionInfoStream, 4);

    // Write translation value
    LVersionInfoStream.WriteData<Word>($0409);  // Language ID (US English)
    LVersionInfoStream.WriteData<Word>($04B0);  // Unicode code page

    LFinalPos := LVersionInfoStream.Position;

    // Update VarFileInfo block length
    LVersionInfoStream.Position := LVarFileInfoStart;
    LVersionInfoStream.WriteData<Word>(LFinalPos - LVarFileInfoStart);

    // Update translation block length
    LVersionInfoStream.Position := LTranslationStart;
    LVersionInfoStream.WriteData<Word>(LFinalPos - LTranslationStart);

    // Update total version info length
    LVersionInfoStream.Position := LVSVersionInfoStart;
    LVersionInfoStream.WriteData<Word>(LFinalPos);

    LDataPtr := LVersionInfoStream.Memory;
    LDataSize := LVersionInfoStream.Size;

    // Update the resource
    LHandleUpdate := BeginUpdateResource(PChar(PEFilePath), False);
    if LHandleUpdate = 0 then
      RaiseLastOSError;

    try
      if not UpdateResourceW(LHandleUpdate, RT_VERSION, MAKEINTRESOURCE(1),
         MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL), LDataPtr, LDataSize) then
        RaiseLastOSError;

      if not EndUpdateResource(LHandleUpdate, False) then
        RaiseLastOSError;
    except
      EndUpdateResource(LHandleUpdate, True);
      raise;
    end;
  finally
    LVersionInfoStream.Free;
  end;
end;

function  wvbResourceExist(const AResName: string): Boolean;
begin
  Result := Boolean((FindResource(HInstance, PChar(AResName), RT_RCDATA) <> 0));
end;

procedure wvbLoadStringListFromResource(const aResName: string; aList: TStringList);
var
  ResStream: TResourceStream;
begin
  ResStream := TResourceStream.Create(HInstance, aResName, RT_RCDATA);
  try
    aList.LoadFromStream(ResStream);
  finally
    ResStream.Free;
  end;
end;

function wvbLoadStringFromResource(const aResName: string): string;
var
  ResStream: TResourceStream;
  StrList: TStringLIst;
begin
  ResStream := TResourceStream.Create(HInstance, aResName, RT_RCDATA);
  try
    StrList := TStringList.Create;
    try
      StrList.LoadFromStream(ResStream);
      Result := StrList.Text;
    finally
      FreeAndNil(StrList);
    end;
  finally
    FreeAndNil(ResStream);
  end;
end;

function wvbAddResManifestFromResource(const aResName: string; const aModuleFile: string; aLanguage: Integer): Boolean;
var
  LHandle: THandle;
  LManifestStream: TResourceStream;
begin
  Result := False;

  if not wvbResourceExist(aResName) then Exit;
  if not TFile.Exists(aModuleFile) then Exit;

  LManifestStream := TResourceStream.Create(HInstance, aResName, RT_RCDATA);
  try
    LHandle := WinAPI.Windows.BeginUpdateResourceW(System.PWideChar(aModuleFile), LongBool(False));

    if LHandle <> 0 then
    begin
      Result := WinAPI.Windows.UpdateResourceW(LHandle, RT_MANIFEST, CREATEPROCESS_MANIFEST_RESOURCE_ID, aLanguage, LManifestStream.Memory, LManifestStream.Size);
      WinAPI.Windows.EndUpdateResourceW(LHandle, False);
    end;
  finally
    FreeAndNil(LManifestStream);
  end;
end;

function wvbGetParentProcessName(): string;
var
  LSnapshot: THandle;
  LProcEntry: TProcessEntry32;
  LCurrentPID: DWORD;
  LParentPID: DWORD;
begin
  Result := '';
  LParentPID := 0; // <-- Initialize here
  LCurrentPID := GetCurrentProcessId();
  LSnapshot := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if LSnapshot = INVALID_HANDLE_VALUE then Exit;

  try
    LProcEntry.dwSize := SizeOf(TProcessEntry32);
    if Process32First(LSnapshot, LProcEntry) then
    begin
      repeat
        if LProcEntry.th32ProcessID = LCurrentPID then
        begin
          LParentPID := LProcEntry.th32ParentProcessID;
          Break;
        end;
      until not Process32Next(LSnapshot, LProcEntry);
    end;

    // Reset and search for parent process
    LProcEntry.dwSize := SizeOf(TProcessEntry32);
    if Process32First(LSnapshot, LProcEntry) then
    begin
      repeat
        if LProcEntry.th32ProcessID = LParentPID then
        begin
          Result := LowerCase(ExtractFileName(LProcEntry.szExeFile));
          Break;
        end;
      until not Process32Next(LSnapshot, LProcEntry);
    end;
  finally
    CloseHandle(LSnapshot);
  end;
end;

function wvbWasStartedFromCommandLine(): Boolean;
var
  LParentProc: string;
begin
  LParentProc := wvbGetParentProcessName();
  Result := LParentProc = 'cmd.exe';
end;

function wvbHasEnoughDiskSpace(const AFilePath: string; ARequiredSize: Int64): Boolean;
var
  LFreeAvailable, LTotalSpace, LTotalFree: Int64;
  LDrive: string;
begin
  Result := False;

  // Resolve the absolute path in case of a relative path
  LDrive := ExtractFileDrive(TPath.GetFullPath(AFilePath));

  // If there is no drive letter, use the current drive
  if LDrive = '' then
    LDrive := ExtractFileDrive(TDirectory.GetCurrentDirectory);

  // Ensure drive has a trailing backslash
  if LDrive <> '' then
    LDrive := LDrive + '\';

  if GetDiskFreeSpaceEx(PChar(LDrive), LFreeAvailable, LTotalSpace, @LTotalFree) then
    Result := LFreeAvailable >= ARequiredSize;
end;

procedure wvbSaveFormState(Form: TForm; IniFile: TIniFile; const SectionName: string = '');
var
  Section: string;
  NormalLeft, NormalTop, NormalWidth, NormalHeight: Integer;
  Placement: TWindowPlacement;
begin
  if not Assigned(Form) then
    raise Exception.Create('Form parameter cannot be nil');

  if not Assigned(IniFile) then
    raise Exception.Create('IniFile parameter cannot be nil');

  // Use form name as section name if not provided
  if SectionName = '' then
    Section := Form.Name
  else
    Section := SectionName;

  // For maximized or minimized forms, get normal position using Win32 API
  if Form.WindowState <> wsNormal then
  begin
    Placement.length := SizeOf(TWindowPlacement);
    if GetWindowPlacement(Form.Handle, @Placement) then
    begin
      with Placement.rcNormalPosition do
      begin
        NormalLeft := Left;
        NormalTop := Top;
        NormalWidth := Right - Left;
        NormalHeight := Bottom - Top;
      end;
    end
    else
    begin
      // Fallback to current values if API fails
      NormalLeft := Form.Left;
      NormalTop := Form.Top;
      NormalWidth := Form.Width;
      NormalHeight := Form.Height;
    end;
  end
  else
  begin
    // Normal window state - use current values
    NormalLeft := Form.Left;
    NormalTop := Form.Top;
    NormalWidth := Form.Width;
    NormalHeight := Form.Height;
  end;

  // Save position and size
  IniFile.WriteInteger(Section, 'Left', NormalLeft);
  IniFile.WriteInteger(Section, 'Top', NormalTop);
  IniFile.WriteInteger(Section, 'Width', NormalWidth);
  IniFile.WriteInteger(Section, 'Height', NormalHeight);

  // Save window state
  IniFile.WriteInteger(Section, 'WindowState', Ord(Form.WindowState));

  // For multi-monitor setups, save which monitor the form is on
  IniFile.WriteInteger(Section, 'MonitorIndex', Screen.MonitorFromWindow(Form.Handle).MonitorNum);
end;

function wvbLoadFormState(Form: TForm; IniFile: TIniFile; const SectionName: string = ''): Boolean;
var
  Section: string;
  WindowState: Integer;
  Left, Top, Width, Height: Integer;
  MonitorIndex: Integer;
  Monitor: TMonitor;
  WorkArea: TRect;
begin
  Result := False;

  if not Assigned(Form) then
    raise Exception.Create('Form parameter cannot be nil');

  if not Assigned(IniFile) then
    raise Exception.Create('IniFile parameter cannot be nil');

  // Use form name as section name if not provided
  if SectionName = '' then
    Section := Form.Name
  else
    Section := SectionName;

  // Check if section exists by reading WindowState
  WindowState := IniFile.ReadInteger(Section, 'WindowState', -1);

  // If no saved state found, return False
  if WindowState = -1 then
    Exit;

  // Determine which monitor to use (multi-monitor support)
  MonitorIndex := IniFile.ReadInteger(Section, 'MonitorIndex', -1);
  if (MonitorIndex >= 0) and (MonitorIndex < Screen.MonitorCount) then
    Monitor := Screen.Monitors[MonitorIndex]
  else
    Monitor := Screen.MonitorFromWindow(Form.Handle);

  // Get working area of the monitor
  WorkArea := Monitor.WorkareaRect;

  // Load position and size
  Left := IniFile.ReadInteger(Section, 'Left', Form.Left);
  Top := IniFile.ReadInteger(Section, 'Top', Form.Top);
  Width := IniFile.ReadInteger(Section, 'Width', Form.Width);
  Height := IniFile.ReadInteger(Section, 'Height', Form.Height);

  // Make sure dimensions are reasonable
  if Width < Form.Constraints.MinWidth then
    Width := Form.Constraints.MinWidth;
  if Width > WorkArea.Width then
    Width := WorkArea.Width;

  if Height < Form.Constraints.MinHeight then
    Height := Form.Constraints.MinHeight;
  if Height > WorkArea.Height then
    Height := WorkArea.Height;

  // Ensure form is visible on screen
  // Make sure at least 100px of the form is visible on the screen
  Left := Max(Left, WorkArea.Left - Width + 100);
  Left := Min(Left, WorkArea.Right - 100);
  Top := Max(Top, WorkArea.Top - Height + 100);
  Top := Min(Top, WorkArea.Bottom - 100);

  // Apply normal size and position
  // (This needs to be done before setting WindowState)
  Form.SetBounds(Left, Top, Width, Height);

  // Set window state (normal, minimized, maximized)
  // Only apply the window state if it's valid
  if (WindowState >= 0) and (WindowState <= 2) then
    Form.WindowState := TWindowState(WindowState);

  Result := True;
end;

function wvbZipFolder(const AFolderPath, ADestZipFilename: string; const AProgressCallback: TwvbZipProgressCallback; const AUserData: Pointer): Boolean;
var
  LFileList: TArray<string>;
  LTotalFiles: Integer;
  LRelativePath: string;
  LFullPath: string;
  I: Integer;
  LZipFile: TZipFile;
begin
  Result := False;

  if not TDirectory.Exists(AFolderPath) then Exit;
  if ADestZipFilename.IsEmpty then Exit;

  TDirectory.CreateDirectory(TPath.GetDirectoryName(ADestZipFilename));

  LZipFile := TZipFile.Create();
  try
    LZipFile.Open(ADestZipFilename, zmWrite);

    // Get all files in the folder and subfolders using TDirectory
    LFileList := TDirectory.GetFiles(AFolderPath, '*', TSearchOption.soAllDirectories);

    LTotalFiles := Length(LFileList);

    // Process each file
    for I := 0 to LTotalFiles - 1 do
    begin
      LFullPath := LFileList[I];
      // Calculate the relative path to use as the filename in the zip
      LRelativePath := ExtractRelativePath(IncludeTrailingPathDelimiter(AFolderPath), LFullPath);

      // Add the file to the zip
      LZipFile.Add(LFullPath, LRelativePath);

      // Call the progress callback
      if Assigned(AProgressCallback) then
      begin
        AProgressCallback(LRelativePath, AUserData);
      end;
    end;

    Result := True;
  finally
    LZipFile.Free();
  end;

  if Result then
    Result := TZipFile.IsValid(ADestZipFilename);
end;

function wbvCopyFile(const ASrcFile, ADestFile: string): Boolean;
var
  FSrc, FDest: TFileStream;
  Buffer: array[0..65535] of Byte;
  BytesRead: Integer;
begin
  Result := False;
  try
    // Make sure the destination directory exists
    TDirectory.CreateDirectory(TPath.GetDirectoryName(ADestFile));

    // Open source with sharing
    FSrc := TFileStream.Create(ASrcFile, fmOpenRead or fmShareDenyNone);
    try
      // Create destination with specific access rights
      FDest := TFileStream.Create(ADestFile, fmCreate);
      try
        // Copy in chunks
        repeat
          BytesRead := FSrc.Read(Buffer, SizeOf(Buffer));
          if BytesRead > 0 then
            FDest.WriteBuffer(Buffer, BytesRead);
        until BytesRead = 0;

        Result := True;
      finally
        FDest.Free;
      end;
    finally
      FSrc.Free;
    end;
  except
    on E: Exception do
    begin
      // Access denied likely means we need elevated privileges
      // Could log error here if needed
    end;
  end;
end;

function wvbAddFileAsResource(const AExeFilename, AResourceName, ADataFileName: string): Boolean;
var
  UpdateHandle: THandle;
  DataStream: TMemoryStream;
begin
  Result := False;

  if not (FileExists(AExeFilename) and FileExists(ADataFileName)) then
    Exit;

  DataStream := TMemoryStream.Create;
  try
    DataStream.LoadFromFile(ADataFileName);

    UpdateHandle := BeginUpdateResource(PChar(AExeFilename), False);
    if UpdateHandle = 0 then
      Exit;

    Result := UpdateResource(UpdateHandle, RT_RCDATA, PChar(AResourceName),
                           MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL),
                           DataStream.Memory, DataStream.Size)
              and EndUpdateResource(UpdateHandle, False);

    if not Result then
      EndUpdateResource(UpdateHandle, True); // Discard changes if failed
  finally
    DataStream.Free;
  end;
end;

function wvbAddDataAsResource(const AExeFilename, AResourceName: string; const AData: Pointer; const ADataSize: Integer): Boolean;
var
  UpdateHandle: THandle;
begin
  Result := False;
  if not FileExists(AExeFilename) then
    Exit;

  UpdateHandle := BeginUpdateResource(PChar(AExeFilename), False);
  if UpdateHandle = 0 then
  begin
    Exit;
  end;

  if UpdateResource(UpdateHandle, RT_RCDATA, PChar(AResourceName),
                         MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL),
                         AData, ADataSize)
            and EndUpdateResource(UpdateHandle, False) then
    Result := True;

  if not Result then
    EndUpdateResource(UpdateHandle, True); // Discard changes if failed
end;

// Convenience function specifically for strings
function wvbAddStringAsResource(const AExeFilename, AResourceName, AStringData: string): Boolean;
var
  UpdateHandle: THandle;
begin
  Result := False;
  if not FileExists(AExeFilename) then
    Exit;

  UpdateHandle := BeginUpdateResource(PChar(AExeFilename), False);
  if UpdateHandle = 0 then
    Exit;

  Result := UpdateResource(UpdateHandle, RT_RCDATA, PChar(AResourceName),
                         MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL),
                         PChar(AStringData), Length(AStringData) * SizeOf(Char))
            and EndUpdateResource(UpdateHandle, False);

  if not Result then
    EndUpdateResource(UpdateHandle, True); // Discard changes if failed
end;

procedure wvbCenterForm(const AForm: TForm);
begin
  AForm.Top := (Screen.WorkAreaHeight - AForm.Height) div 2;
  if (AForm.Top < 0) then
    AForm.Top := 0;
  AForm.Left := (Screen.WorkAreaWidth - AForm.Width) div 2;
  if (AForm.Left < 0) then
    AForm.Left := 0;
end;

procedure wvbCenterForm(const AForm, AMainForm: TForm);
begin
  AForm.Top := AMainForm.Top + ((AMainForm.Height - AForm.Height) div 2);
  if (AForm.Top < 0) then
    AForm.Top := 0;
  AForm.Left := AMainForm.Left + ((AMainForm.Width - AForm.Width) div 2);
  if (AForm.Left < 0) then
    AForm.Left := 0;
end;

function wvbConfirmDlg(const AMsg: string; const AArgs: array of const): TwvbConfirmDialogResult;
begin
  Result := cdCancel;
  case MessageDlg(Format(AMsg, AArgs), mtConfirmation,
    [mbYes, mbNo, mbCancel], 0) of
    mrYes:
      Result := cdYes;
    mrNo:
      Result := cdNo;
    mrCancel:
      Result := cdCancel;
  end;
end;

function wvbCreateDirsInPath(const AFilename: string): Boolean;
var
  s: string;
begin
  Result := False;

  if AFilename = '' then
    Exit;

  s := TPath.GetDirectoryName(AFilename);
  if s = '' then
    Exit;

  TDirectory.CreateDirectory(s);

  Result := TDirectory.Exists(s)
end;

function  wvbExtractResToFile(const AResName, AFilename: string): Boolean;
var
  LResStream: TResourceStream;
begin
  Result := False;
  if AResName.IsEmpty then Exit;
  if AFilename.IsEmpty then Exit;
  if not wvbResourceExist(AResName) then Exit;
  LResStream := TResourceStream.Create(HInstance, AResName, RT_RCDATA);
  try
    wvbCreateDirsInPath(AFilename);
    LResStream.SaveToFile(AFilename);
  finally
    FreeAndNil(LResStream);
  end;
  Result := TFile.Exists(AFilename);
end;

function wvbRunExe(const AFilename, ACmdLine, AStartPath: string): Boolean;
var
  LStartInfo: TStartupInfo;
  LProcInfo: TProcessInformation;
  LCreateOk: boolean;
  LFilename: string;
begin
  Result := False;
  if AFilename.IsEmpty then Exit;
  //LFilename := TPath.ChangeExtension(aFilename, '.exe');
  LFilename := AFilename;
  if not TFile.Exists(LFilename) then Exit;

  { fill with known state }
  FillChar(LStartInfo, SizeOf(TStartupInfo), #0);
  FillChar(LProcInfo, SizeOf(TProcessInformation), #0);
  LStartInfo.cb := SizeOf(TStartupInfo);
  LStartInfo.lpTitle := 'Test';

  LCreateOk := CreateProcess(PChar(AFilename), PChar(ACmdLine), nil, nil, false,
    CREATE_NEW_PROCESS_GROUP + NORMAL_PRIORITY_CLASS, nil, PChar(AStartPath),
    LStartInfo, LProcInfo);


  { check to see if successful }
  if LCreateOk then
  begin
    // may or may not be needed. Usually wait for child processes
    WaitForSingleObject(LProcInfo.hProcess, INFINITE);
    CloseHandle(LProcInfo.hProcess);
    CloseHandle(LProcInfo.hThread);
  end;
end;

function wvbExtractAndRunResEXE(const AResName, ACmdLine, AStartPath: string): Boolean;
var
  LResStream: TResourceStream;
  LFilename: string;
begin
  Result := False;
  if not wvbResourceExist(AResName) then Exit;
  LResStream := TResourceStream.Create(HInstance, AResName, RT_RCDATA);
  LFilename := TPath.Combine(AStartPath, TPath.GetTempFileName);

  TDirectory.CreateDirectory(AStartPath);
  LResStream.SaveToFile(LFilename);
  FreeAndNil(LResStream);
  Result := wvbRunExe(LFilename, ACmdLine, AStartPath);
  TFile.Delete(LFilename);
end;

procedure wvbShowMessage(const ATitle, AText: string; const AArgs: array of const; AType: TwvbShowMessage);
var
  LText: string;
begin
  LText := Format(AText, AArgs);
  MessageBox(0, PChar(LText), PChar(ATitle), Ord(AType));
end;


{ TlfpDirectoryStack }
constructor TwvbDirectoryStack.Create;
begin
  inherited;
  FStack := TStack<String>.Create;
end;

destructor TwvbDirectoryStack.Destroy;
begin
  FreeAndNil(FStack);
  inherited;
end;

procedure TwvbDirectoryStack.Push(aPath: string);
var
  s: string;
begin
  s := GetCurrentDir;
  FStack.Push(s);
  if not s.IsEmpty then
  begin
    SetCurrentDir(aPath);
  end;
end;

procedure TwvbDirectoryStack.PushFilePath(aFilename: string);
var
  LDir: string;
begin
  LDir := TPath.GetDirectoryName(aFilename);
  if LDir.IsEmpty then Exit;
  Push(LDir);
end;

procedure TwvbDirectoryStack.Pop;
var
  s: string;
begin
  if FStack.Count = 0 then Exit;
  s := FStack.Pop;
  SetCurrentDir(s);
end;

{ TlfpCmdLine }
class function TwvbCmdLine.GetCmdLine: PChar;
begin
  Result := PChar(FCmdLine);
end;

class function TwvbCmdLine.GetParamStr(aParamStr: PChar; var aParam: string): PChar;
var
  i, Len: Integer;
  Start, S: PChar;
begin
  // U-OK
  while True do
  begin
    while (aParamStr[0] <> #0) and (aParamStr[0] <= ' ') do
      Inc(aParamStr);
    if (aParamStr[0] = '"') and (aParamStr[1] = '"') then Inc(aParamStr, 2) else Break;
  end;
  Len := 0;
  Start := aParamStr;
  while aParamStr[0] > ' ' do
  begin
    if aParamStr[0] = '"' then
    begin
      Inc(aParamStr);
      while (aParamStr[0] <> #0) and (aParamStr[0] <> '"') do
      begin
        Inc(Len);
        Inc(aParamStr);
      end;
      if aParamStr[0] <> #0 then
        Inc(aParamStr);
    end
    else
    begin
      Inc(Len);
      Inc(aParamStr);
    end;
  end;

  SetLength(aParam, Len);

  aParamStr := Start;
  S := Pointer(aParam);
  i := 0;
  while aParamStr[0] > ' ' do
  begin
    if aParamStr[0] = '"' then
    begin
      Inc(aParamStr);
      while (aParamStr[0] <> #0) and (aParamStr[0] <> '"') do
      begin
        S[i] := aParamStr^;
        Inc(aParamStr);
        Inc(i);
      end;
      if aParamStr[0] <> #0 then Inc(aParamStr);
    end
    else
    begin
      S[i] := aParamStr^;
      Inc(aParamStr);
      Inc(i);
    end;
  end;

  Result := aParamStr;
end;

class function TwvbCmdLine.ParamCount: Integer;
var
  P: PChar;
  S: string;
begin
  // U-OK
  Result := 0;
  P := TwvbCmdLine.GetParamStr(GetCmdLine, S);
  while True do
  begin
    P := TwvbCmdLine.GetParamStr(P, S);
    if S = '' then Break;
    Inc(Result);
  end;
end;

class procedure TwvbCmdLine.ClearParams;
begin
  FCmdLine := '';
end;

class procedure TwvbCmdLine.Reset;
begin
  // init commandline
  FCmdLine := System.CmdLine + ' ';
end;

class procedure TwvbCmdLine.AddAParam(const aParam: string);
var
  LParam: string;
begin
  LParam := aParam.Trim;
  if LParam.IsEmpty then Exit;
  FCmdLine := FCmdLine + LParam + ' ';
end;

class procedure TwvbCmdLine.AddParams(const aParams: string);
begin
  var LParams := aParams.Split([' '], TStringSplitOptions.ExcludeEmpty);
  for var I := 0 to Length(LParams)-1 do
  begin
    AddAParam(LParams[I]);
  end;
end;

class function TwvbCmdLine.ParamStr(aIndex: Integer): string;
var
  P: PChar;
  Buffer: array[0..260] of Char;
begin
  Result := '';
  if aIndex = 0 then
    SetString(Result, Buffer, GetModuleFileName(0, Buffer, Length(Buffer)))
  else
  begin
    P := GetCmdLine;
    while True do
    begin
      P := TwvbCmdLine.GetParamStr(P, Result);
      if (aIndex = 0) or (Result = '') then Break;
      Dec(aIndex);
    end;
  end;
end;

class function TwvbCmdLine.GetParamValue(const aParamName: string; aSwitchChars: TSysCharSet; aSeperator: Char; var aValue: string): Boolean;
var
  i, Sep: Longint;
  s: string;
begin

  Result := False;
  aValue := '';

  // check for first non switch param when aParamName = '' and no
  // other params are found
  if (aParamName = '') then
  begin
    for i := 1 to TwvbCmdLine.ParamCount do
    begin
      s := TwvbCmdLine.ParamStr(i);
      if Length(s) > 0 then
        // if S[1] in aSwitchChars then
        if not CharInSet(s[1], aSwitchChars) then
        begin
          aValue := s;
          Result := True;
          Exit;
        end;
    end;
    Exit;
  end;

  // check for switch params
  for i := 1 to TwvbCmdLine.ParamCount do
  begin
    s := TwvbCmdLine.ParamStr(i);
    if Length(s) > 0 then
      // if S[1] in aSwitchChars then
      if CharInSet(s[1], aSwitchChars) then

      begin
        Sep := Pos(aSeperator, s);

        case Sep of
          0:
            begin
              if CompareText(Copy(s, 2, Length(s) - 1), aParamName) = 0 then
              begin
                Result := True;
                Break;
              end;
            end;
          1 .. MaxInt:
            begin
              if CompareText(Copy(s, 2, Sep - 2), aParamName) = 0 then
              // if CompareText(Copy(S, 1, Sep -1), aParamName) = 0 then
              begin
                aValue := Copy(s, Sep + 1, Length(s));
                Result := True;
                Break;
              end;
            end;
        end; // case
      end
  end;

end;

// GetParameterValue('p', ['/', '-'], '=', sValue);
class function TwvbCmdLine.GetParamValue(const aParamName: string; var aValue: string): Boolean;
begin
  Result := TwvbCmdLine.GetParamValue(aParamName, ['/', '-'], '=', aValue);
end;

class function TwvbCmdLine.GetParam(const aParamName: string): Boolean;
var
  LValue: string;
begin
  Result := TwvbCmdLine.GetParamValue(aParamName, ['/', '-'], '=', LValue);
  if not Result then
  begin
    Result := SameText(aParamName, TwvbCmdLine.ParamStr(1));
  end;
end;

{ TwvbPayloadStream }
procedure TwvbPayloadStream.InitFooter(out Footer: TPayloadStreamFooter);
begin
  FillChar(Footer, SizeOf(Footer), 0);
  Footer.WaterMark := cWaterMarkGUID;
end;

function TwvbPayloadStream.ReadFooter(const FileStream: TFileStream;
  out Footer: TPayloadStreamFooter): Boolean;
var
  FileLen: Int64;
begin
  // Check that file is large enough for a footer!
  FileLen := FileStream.Size;
  if FileLen > SizeOf(Footer) then
  begin
    // Big enough: move to start of footer and read it
    FileStream.Seek(-SizeOf(Footer), soEnd);
    FileStream.Read(Footer, SizeOf(Footer));
  end
  else
    // File not large enough for footer: zero it
    // .. this ensures watermark is invalid
    FillChar(Footer, SizeOf(Footer), 0);
  // Return if watermark is valid
  Result := IsEqualGUID(Footer.WaterMark, cWaterMarkGUID);
end;

constructor TwvbPayloadStream.Create(const FileName: string; const Mode: TPayloadStreamOpenMode);
var
  Footer: TPayloadStreamFooter; // footer record for payload data
begin
  inherited Create;
  // Open file stream
  fMode := Mode;
  case fMode of
    pomRead: fFileStream := TFileStream.Create(FileName, fmOpenRead or fmShareDenyWrite);
    pomWrite: fFileStream := TFileStream.Create(FileName, fmOpenReadWrite or fmShareExclusive);
  end;
  // Check for existing payload
  if ReadFooter(fFileStream, Footer) then
  begin
    // We have payload: record start and size of data
    fDataStart := Footer.ExeSize;
    fDataSize := Footer.DataSize;
  end
  else
  begin
    // There is no existing payload: start is end of file
    fDataStart := fFileStream.Size;
    fDataSize := 0;
  end;
  // Set initial file position per mode
  case fMode of
    pomRead: fFileStream.Seek(fDataStart, soBeginning);
    pomWrite: fFileStream.Seek(fDataStart + fDataSize, soBeginning);
  end;
end;

destructor TwvbPayloadStream.Destroy;
var
  Footer: TPayloadStreamFooter; // payload footer record
begin
  if fMode = pomWrite then
  begin
    // We're in write mode: we need to update footer
    if fDataSize > 0 then
    begin
      // We have payload, so need a footer record
      InitFooter(Footer);
      Footer.ExeSize := fDataStart;
      Footer.DataSize := fDataSize;
      fFileStream.Seek(0, soEnd);
      fFileStream.WriteBuffer(Footer, SizeOf(Footer));
    end
    else
    begin
      // No payload => no footer
      fFileStream.Size := fDataStart;
    end;
  end;
  // Free file stream
  FreeAndNil(fFileStream);
  inherited;
end;

function TwvbPayloadStream.Seek(const Offset: Int64;
  Origin: TSeekOrigin): Int64;
begin
  // Perform actual seek in underlying file stream
  Result := fFileStream.Seek(Offset, Origin);
end;

procedure TwvbPayloadStream.SetSize(const NewSize: Int64);
begin
  // Set size of file stream
  fFileStream.Size := NewSize;
end;

function TwvbPayloadStream.Read(var Buffer; Count: LongInt): LongInt;
begin
  // Read data from file stream and return bytes read
  Result := fFileStream.Read(Buffer, Count);
end;

function TwvbPayloadStream.Write(const Buffer; Count: LongInt): LongInt;
begin
  // Check in write mode
  if fMode <> pomWrite then
  begin
    raise EPayloadStream.Create(
      'TkaPayloadStream can''t write in read mode.');
  end;
  // Write the data to file stream and return bytes written
  Result := fFileStream.Write(Buffer, Count);
  // Check if stream has grown
  fDataSize := Max(fDataSize, fFileStream.Position - fDataStart);
end;

function TwvbPayloadStream.GetPosition: Int64;
begin
  Result := fFileStream.Position - fDataStart;
end;

procedure TwvbPayloadStream.SetPosition(const Value: Int64);
begin
  fFileStream.Position := fDataStart + Value;
end;

end.
