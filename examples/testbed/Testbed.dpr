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
