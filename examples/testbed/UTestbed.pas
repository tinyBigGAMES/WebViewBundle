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

unit UTestbed;

interface

uses
  WinApi.Windows,
  System.SysUtils,
  WebViewBundle;

procedure RunTests();

implementation

{ -----------------------------------------------------------------------------
 Test Constants and Utility Procedure
 This section defines constants used throughout the test suite for specifying
 source locations and output destinations related to HTML content, ZIP bundles,
 and standalone executables. It also includes a simple utility procedure to
 pause program execution, which is often useful in console applications to
 allow users to review output before the program terminates.
------------------------------------------------------------------------------ }

// --- Constants defining paths for tests ---
const
  // CHtml: Specifies the relative path to the source directory containing the
  // raw HTML files and assets.
  CHtml    = 'res\untree.co-sterial';

  // CHtmlZip: Specifies the relative path and filename for the output ZIP
  // archive bundle created from the HTML source.
  CHtmlZip = 'output\zip\Html.zip';

  // CHtmlExe: Specifies the relative path and filename for the final bundled
  // standalone executable file.
  CHtmlExe = 'output\exe\Html.exe';

{ -----------------------------------------------------------------------------
 Pause: Wait for User Input
 This procedure pauses the execution of the program and waits for the user
 to press the ENTER key. It is typically used in console applications or
 test programs to keep the output window open and visible until the user
 is ready to continue or close the application.
------------------------------------------------------------------------------ }
procedure Pause();
begin
  // Print a blank line for spacing in the console output.
  WriteLn;
  // Display a message prompting the user to press the Enter key.
  Write('Press ENTER to continue...');
  // Wait for the user to press the Enter key. The program execution halts here.
  ReadLn;
  // Print another blank line for spacing after the user has pressed Enter.
  WriteLn;
end;

{ -----------------------------------------------------------------------------
 LoadPayload: Load Embedded Web Content from Executable
 This procedure is designed to be called when the application is running as a
 standalone executable that has been bundled with web content (HTML, CSS,
 JavaScript, etc.) using the TWebViewBundle library. Its primary function
 is to create an instance of the TWebViewBundle component and then instruct
 it to extract and load the previously embedded web resources. The result
 is typically the display of the web content within a web view window.
 NOTE: This procedure assumes that the web content payload has been successfully
       embedded into the executable file during a prior bundling step (e.g.,
       using TWebViewBundle.BundleHtml). It demonstrates the simple runtime
       step required to make that embedded content accessible and visible.
------------------------------------------------------------------------------ }
procedure LoadPayload();
var
  LWebViewBundle: TWebViewBundle; // Instance of the TWebViewBundle component used for loading the embedded content
begin
  // Create a new instance of the TWebViewBundle object.
  // This object manages the web view component and access to embedded data.
  LWebViewBundle := TWebViewBundle.Create();
  try
    // Use a try..finally block to ensure that the created object is always properly destroyed,
    // even if an error occurs during payload loading.

    // Call the LoadPayload method. This is the core action which tells the TWebViewBundle
    // instance to find, extract, and render the web content that was embedded
    // within the current executable file. This will typically open the web view window.
    LWebViewBundle.LoadPayload();

    // Note: After LoadPayload is called, the procedure usually finishes.
    // The web view window will remain open and interactive, controlled by the
    // application's message loop (often managed internally by TWebViewBundle
    // or the framework), until it is closed by the user or programmatically.

  finally
    // This block is guaranteed to execute regardless of how the try block exits.

    // Free the TWebViewBundle object. This releases associated system resources,
    // including closing the web view window if it's still open and wasn't already closed.
    LWebViewBundle.Free();
  end;
end;

{ -----------------------------------------------------------------------------
 Test01: Zip HTML Content from Folder
 This procedure demonstrates how to package HTML content and its associated
 assets (like CSS, JavaScript, images) from a specified source directory
 into a single ZIP archive file. It utilizes the TWebViewBundle library's
 internal zipping functionality (ZipHtml method) for this purpose. The
 procedure includes a callback function to display the name of each file
 being added to the archive in real-time console output, providing feedback
 during the zipping process.
------------------------------------------------------------------------------ }
procedure Test01();
var
  LWebViewBundle: TWebViewBundle; // Instance of the TWebViewBundle object used to access zipping functions
begin
  // Create a new instance of the TWebViewBundle component.
  LWebViewBundle := TWebViewBundle.Create();
  try
    // Use a try..finally block to guarantee that the object is destroyed.

    // Output a message to the console indicating that the zipping operation is starting.
    Writeln('Zipping Html...');

    // Call the ZipHtml method to compress the contents of the source HTML folder (CHtml)
    // into a ZIP file specified by CHtmlZip.
    // The third parameter is an optional callback procedure that is executed for each file added.
    // The fourth parameter (nil in this case) is UserData passed to the callback.
    if LWebViewBundle.ZipHtml(CHtml, CHtmlZip,
      // This is an anonymous procedure that serves as the callback.
      // It receives the filename being added to the ZIP.
      procedure(const AFilename: string; const UserData: Pointer)
      begin
        // Format and print the filename to the console as it's being added to the archive.
        WriteLn(Format('Adding "%s"...', [AFilename]));
      end,
    nil) then // Check the boolean result of the ZipHtml function
    begin
      // If ZipHtml returns True, the operation was successful.
      WriteLn('Success!');
    end
    else
    begin
      // If ZipHtml returns False, the operation failed.
      WriteLn('Failed');
    end;

  finally
    // This block is executed regardless of whether the try block completed normally or raised an exception.

    // Destroy the TWebViewBundle object to release its resources and memory.
    LWebViewBundle.Free();
  end;
end;

{ -----------------------------------------------------------------------------
 Test02: Load ZIP-Packed HTML Bundle from File
 This procedure demonstrates how to load and display HTML content that has been
 bundled into a separate ZIP file using the TWebViewBundle library. It sets
 the title of the web view window, specifies the path to the external ZIP
 bundle file, and then instructs the TWebViewBundle instance to load the
 payload from that file.
 NOTE: This example assumes that the file specified by LWebViewBundle.BundleFilename
       ('output/zip/Html.zip') exists in the correct location relative to the
       executable and contains a valid ZIP-packed HTML structure that TWebViewBundle
       can interpret and display.
------------------------------------------------------------------------------ }
procedure Test02();
var
  LWebViewBundle: TWebViewBundle; // Instance of the TWebViewBundle component
begin
  // Create a new instance of the TWebViewBundle object
  LWebViewBundle := TWebViewBundle.Create();
  try
    // Use a try..finally block to ensure proper resource cleanup

    // Set the text that will appear in the title bar of the created web view window
    LWebViewBundle.WindowTitle := 'Test02';

    // Specify the full path or relative path to the external ZIP file containing the HTML content
    LWebViewBundle.BundleFilename := 'output/zip/Html.zip';

    // Instruct the TWebViewBundle instance to load and display the HTML content from the specified BundleFilename
    // This method initiates the process of extracting and rendering the web content.
    LWebViewBundle.LoadPayload();

    // Note: LoadPayload will typically open a window and run the content.
    // The procedure will likely complete here, and the application's message loop
    // (handled internally by TWebViewBundle or the application framework)
    // will keep the window open until closed by the user.

  finally
    // This block is always executed when the try block is exited (normally or via exception)

    // Free the TWebViewBundle object to release system resources and close the window (if still open)
    LWebViewBundle.Free();
  end;
end;

{ -----------------------------------------------------------------------------
 Test03: Bundle ZIP-Packed HTML into Standalone EXE
 This procedure demonstrates how to bundle a ZIP-packed HTML archive into a
 standalone executable using the TWebViewBundle library. It sets various
 properties for the resulting EXE file, such as the window title, application
 icon, resize capabilities, version information, and other metadata fields.
 Finally, it performs the bundling operation to create the executable file.
------------------------------------------------------------------------------ }
procedure Test03();
var
  LWebViewBundle: TWebViewBundle; // Instance of the TWebViewBundle component used for the bundling process
begin
  // Create a new instance of the TWebViewBundle object
  LWebViewBundle := TWebViewBundle.Create();
  try
    // Use a try..finally block to ensure the object is always freed

    // Print a message to the console indicating the start of the bundling operation
    WriteLn(Format('Bundling "%s" inside "%s"...', [CHtmlZip, CHtmlExe]));

    // --- Configure Properties for the Output Executable ---

    // Set the title bar text for the bundled application's window
    LWebViewBundle.WindowTitle := 'Test03';
    // Specify the filename of the icon file to be embedded in the EXE
    LWebViewBundle.ExeIconFilename := 'res/icons/html5.ico';
    // Determine if the bundled application's main window can be resized by the user
    LWebViewBundle.WindowResizeable := True;

    // Enable and set the version information resources for the EXE file's properties
    LWebViewBundle.AddVersionInfo := True;
    // Set the major version number
    LWebViewBundle.MajorVer := 0;
    // Set the minor version number
    LWebViewBundle.MinorVer := 1;
    // Set the patch version number
    LWebViewBundle.PatchVer := 0;
    // Note: Build version is typically added automatically or set separately if needed

    // Set standard file metadata properties for the bundled EXE
    LWebViewBundle.ProductName := 'Test03';      // The name of the product
    LWebViewBundle.Description := 'Test03';     // A brief description of the file
    LWebViewBundle.CompanyName := ' tinyBigGAMES™ LLC'; // The company that produced the file
    LWebViewBundle.Copyright := 'Copyright © 2024-present tinyBigGAMES™ LLC'; // Copyright information

    // --- Perform the Bundling ---

    // Execute the bundling operation, embedding the content from CHtmlZip into the executable specified by CHtmlExe
    LWebViewBundle.BundleHtml(CHtmlZip, CHtmlExe);

  finally
    // This block is guaranteed to execute regardless of errors

    // Free the TWebViewBundle object to release allocated memory and resources
    LWebViewBundle.Free();
  end;
end;

{ -----------------------------------------------------------------------------
 RunTests: Main Test Execution or Payload Loader
 This procedure serves as the entry point for either running a test suite for
 the TWebViewBundle library or loading an embedded payload if the application
 is running as a bundled executable.

 It first checks if the executable contains an embedded payload using
 TWebViewBundle.HasPayload(). If a payload is found, it proceeds to load
 and execute it via LoadPayload().

 If no payload is found (indicating it's likely a test/development build),
 it displays the library version, selects a specific test number (currently
 hardcoded), executes the corresponding test procedure (Test01, Test02, etc.),
 and finally pauses execution to allow the user to review the output.
------------------------------------------------------------------------------ }
procedure RunTests();
var
  LNum: Integer; // Variable to hold the selected test number
begin
  // Check if the current executable contains an embedded payload
  if TWebViewBundle.HasPayload() then
  begin
    // If a payload exists, load and execute it and then exit the procedure
    LoadPayload();
    Exit;
  end;

  // If no payload is found (running in test mode), display library version
  WriteLn('WebViewBundle v', TWebViewBundle.GetVersion());
  WriteLn; // Add an empty line for formatting

  // Set the desired test number to run
  LNum := 01;

  // Execute the selected test based on the value of LNum
  case LNum of
    01: Test01(); // Call the Test01 procedure
    02: Test02(); // Call the Test02 procedure
    03: Test03(); // Call the Test03 procedure
    // Add more cases for additional tests
  end;

  // Pause execution to keep the console window open after tests finish
  Pause();
end;

end.
