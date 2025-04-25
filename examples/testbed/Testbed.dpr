program Testbed;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  WebViewBundle in '..\..\src\WebViewBundle.pas',
  WebViewBundle.Utils in '..\..\src\WebViewBundle.Utils.pas',
  WebViewBundle.Form in '..\..\src\WebViewBundle.Form.pas' {WebViewBundleForm},
  UTestbed in 'UTestbed.pas',
  Dlluminator in '..\..\src\Dlluminator.pas',
  Winapi.EdgeUtils in '..\..\src\Winapi.EdgeUtils.pas';

begin
  try
    RunTests();
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.
