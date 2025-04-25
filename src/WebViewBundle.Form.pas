unit WebViewBundle.Form;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Winapi.WebView2, Winapi.ActiveX,
  Vcl.Edge, System.IOUtils, System.Zip, Vcl.ExtCtrls, Vcl.Imaging.jpeg, WebViewBundle.Utils;

const
  CwvbLocalHost  = 'https://wvb.app';
  CwvbWebAppResName = '767ba3c42c4d4626a73321eac2ad6baf';
  CwvbConfigResName = '4666a4e82a32489fbba33e00a0ae985c';
  CwvbExeManifestResName =  'a679ac7eea6e49f19eddce39e09425f1';
  CwvWebview2SetupResName = 'bf991b5cb5fb4fd18f43ff9c8fe3de2a';

type
  TWebViewBundleForm = class(TForm)
    Browser: TEdgeBrowser;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BrowserCreateWebViewCompleted(Sender: TCustomEdgeBrowser;
      AResult: HRESULT);
    procedure BrowserDocumentTitleChanged(Sender: TCustomEdgeBrowser;
      const ADocumentTitle: string);
    procedure BrowserNewWindowRequested(Sender: TCustomEdgeBrowser;
      Args: TNewWindowRequestedEventArgs);
    procedure BrowserWebResourceRequested(Sender: TCustomEdgeBrowser;
      Args: TWebResourceRequestedEventArgs);
    procedure BrowserWindowCloseRequested(Sender: TObject);
    procedure BrowserNavigationCompleted(Sender: TCustomEdgeBrowser;
      IsSuccess: Boolean; WebErrorStatus: COREWEBVIEW2_WEB_ERROR_STATUS);
  private
    { Private declarations }
    FFileStream: TStream;
    FZipFile: TZipFile;
    FIndexPage: string;
    FWindowTitle: string;
    FBundleFilename: string;
    FUserDataFolder: string;
    function StreamFileExists(const AFilename: string): Boolean;
    function GetStream(const AFilename: string): TStream;
  public
    { Public declarations }
    property IndexPage: string read FIndexPage write FIndexPage;
    property WindowTitle: string read FWindowTitle write FWindowTitle;
    property BundleFilename: string read FBundleFilename write FBundleFilename;
    property UserDataFolder: string read FUserDataFolder write FUserDataFolder;


    function  Load(): Boolean;
    procedure Unload();
  end;

var
  WebViewBundleForm: TWebViewBundleForm;

implementation

{$R *.dfm}

function TWebViewBundleForm.StreamFileExists(const AFilename: string): Boolean;
begin
  Result := Boolean(FZipFile.IndexOf(AFilename) <> -1);
end;

function TWebViewBundleForm.GetStream(const AFilename: string): TStream;
var
  LFilename: string;
  LHeader: TZipHeader;
  LResult: TStream;
begin
  LFilename := AFilename;

  if LFilename.StartsWith('/') then
    LFilename := LFilename.Remove(0, 1);
  LFilename := LFilename.Trim;
  if LFilename.IsEmpty then
  begin
    LFilename := FIndexPage;
  end;

  LFilename := LFilename.Replace('/', '\', [rfReplaceAll]);

  Result := TMemoryStream.Create();

  if not StreamFileExists(LFilename) then
  begin
    //WriteLn(Format('File was not found: "%s"', [LFilename]));
    Exit;
  end;
  LResult := nil;
  FZipFile.Read(LFilename, LResult, LHeader);

  Result.CopyFrom(LResult);
  LResult.Free();
end;

procedure TWebViewBundleForm.BrowserCreateWebViewCompleted(Sender: TCustomEdgeBrowser; AResult: HRESULT);
begin
  //
  Sender.AddWebResourceRequestedFilter(CwvbLocalHost + '*', COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL);
end;

procedure TWebViewBundleForm.BrowserDocumentTitleChanged(Sender: TCustomEdgeBrowser; const ADocumentTitle: string);
begin
  //
  if FWindowTitle.IsEmpty then
    Caption := ADocumentTitle
  else
    Caption := FWindowTitle + ' - ' + ADocumentTitle;
end;

procedure TWebViewBundleForm.BrowserNavigationCompleted(
  Sender: TCustomEdgeBrowser; IsSuccess: Boolean;
  WebErrorStatus: COREWEBVIEW2_WEB_ERROR_STATUS);
begin
  //
end;

procedure TWebViewBundleForm.BrowserNewWindowRequested(Sender: TCustomEdgeBrowser; Args: TNewWindowRequestedEventArgs);
var
  URI: PWideChar;
begin
  //
  Args.ArgsInterface.Get_uri(URI);
  if String(URI).StartsWith(CwvbLocalHost, True) then
  begin
    Args.ArgsInterface.Set_Handled(1);
    Browser.Navigate(URI);
  end;
end;

procedure TWebViewBundleForm.BrowserWebResourceRequested(Sender: TCustomEdgeBrowser; Args: TWebResourceRequestedEventArgs);
var
  LRequest:  ICoreWebView2WebResourceRequest;
  LURI:      PWideChar;
  LFileName: String;
  LResponse: ICoreWebView2WebResourceResponse;
  LContentType: string;
  LExtension: string;
begin
  Args.ArgsInterface.Get_Request(LRequest);
  LRequest.Get_uri(LURI);
  LFileName := LURI;
  if LFileName.StartsWith(CwvbLocalHost, True) then
    LFileName := Copy(LFileName, Length(CwvbLocalHost) + 1)
  else
    Exit;
  if LFileName.Contains('?') then
    LFileName := Copy(LFileName, 1, Pos('?', LFileName) - 1);

  // Default to empty content type (let WebView decide)
  LContentType := '';

  // Extract file extension
  LExtension := ExtractFileExt(LFileName).ToLower();

  // Set content type for specific file types
  // HTML/Web related
  if LExtension = '.svg' then
    LContentType := 'Content-Type: image/svg+xml'
  else if LExtension = '.html' then
    LContentType := 'Content-Type: text/html'
  else if LExtension = '.htm' then
    LContentType := 'Content-Type: text/html'
  else if LExtension = '.css' then
    LContentType := 'Content-Type: text/css'
  else if LExtension = '.js' then
    LContentType := 'Content-Type: application/javascript'
  else if LExtension = '.mjs' then
    LContentType := 'Content-Type: application/javascript'
  else if LExtension = '.json' then
    LContentType := 'Content-Type: application/json'
  else if LExtension = '.xml' then
    LContentType := 'Content-Type: application/xml'
  else if LExtension = '.txt' then
    LContentType := 'Content-Type: text/plain'
  else if LExtension = '.md' then
    LContentType := 'Content-Type: text/markdown'
  // Images
  else if LExtension = '.png' then
    LContentType := 'Content-Type: image/png'
  else if LExtension = '.jpg' then
    LContentType := 'Content-Type: image/jpeg'
  else if LExtension = '.jpeg' then
    LContentType := 'Content-Type: image/jpeg'
  else if LExtension = '.gif' then
    LContentType := 'Content-Type: image/gif'
  else if LExtension = '.webp' then
    LContentType := 'Content-Type: image/webp'
  else if LExtension = '.ico' then
    LContentType := 'Content-Type: image/x-icon'
  // Fonts
  else if LExtension = '.ttf' then
    LContentType := 'Content-Type: font/ttf'
  else if LExtension = '.otf' then
    LContentType := 'Content-Type: font/otf'
  else if LExtension = '.woff' then
    LContentType := 'Content-Type: font/woff'
  else if LExtension = '.woff2' then
    LContentType := 'Content-Type: font/woff2'
  // Documents
  else if LExtension = '.pdf' then
    LContentType := 'Content-Type: application/pdf'
  // Media
  else if LExtension = '.mp3' then
    LContentType := 'Content-Type: audio/mpeg'
  else if LExtension = '.mp4' then
    LContentType := 'Content-Type: video/mp4'
  else if LExtension = '.webm' then
    LContentType := 'Content-Type: video/webm'
  // Data formats
  else if LExtension = '.csv' then
    LContentType := 'Content-Type: text/csv'
  else if LExtension = '.yaml' then
    LContentType := 'Content-Type: application/x-yaml'
  else if LExtension = '.yml' then
    LContentType := 'Content-Type: application/x-yaml';

  try
    Sender.EnvironmentInterface.CreateWebResourceResponse(
      TStreamAdapter.Create(GetStream(LFileName), soOwned),
      200,
      'OK',
      PWideChar(LContentType),
      LResponse);
  except
    Sender.EnvironmentInterface.CreateWebResourceResponse(
      nil, 404, 'Not Found', '', LResponse);
  end;

  Args.ArgsInterface.Set_Response(LResponse);
end;

procedure TWebViewBundleForm.BrowserWindowCloseRequested(Sender: TObject);
begin
  //
end;

procedure TWebViewBundleForm.FormCreate(Sender: TObject);
begin
  //
  FIndexPage := 'index.html';
  FWindowTitle := 'WebViewBundle';
  FBundleFilename := 'WebViewBundle.zip';
  FUserDataFolder := 'data';
end;

procedure TWebViewBundleForm.FormDestroy(Sender: TObject);
begin
  //
  Unload();
end;


function TWebViewBundleForm.Load(): Boolean;
begin
  Result := False;

  Unload();

  if wvbResourceExist(CwvbWebAppResName) then
    begin
      FFileStream := TResourceStream.Create(HInstance, CwvbWebAppResName, RT_RCDATA)
    end
  else
    begin
      if TFile.Exists(FBundleFilename) then
      begin
        FFileStream := TFile.OpenRead(FBundleFilename);
      end;
    end;

  if Assigned(FFileStream) then
  begin
    FZipFile := TZipFile.Create();
    FZipFile.Open(FFileStream, zmRead);
    Browser.UserDataFolder := FUserDataFolder;
    Browser.Navigate(CwvbLocalHost);
    Result := True;
  end;
end;

procedure TWebViewBundleForm.Unload();
begin
  if Assigned(FZipFile) then
  begin
    FZipFile.Free();
    FZipFile := nil;
  end;

  if Assigned(FFileStream) then
  begin
    FFileStream.Free();
    FFileStream := nil;
  end;
end;

end.
