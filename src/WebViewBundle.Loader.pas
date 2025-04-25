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

unit WebViewBundle.Loader;

{$I WebViewBundle.Defines.inc}

interface

/// <summary>
///   Loads a dynamic-link library (DLL) from memory into the calling process's address space.
/// </summary>
/// <param name="AData">
///   A pointer to the memory location containing the DLL image.
/// </param>
/// <returns>
///   A handle to the loaded module if successful; otherwise, <c>0</c>.
/// </returns>
/// <remarks>
///   - This function allows for loading a DLL directly from memory, bypassing the need for a file on disk.
///   - The memory block pointed to by <paramref name="AData"/> must contain a valid DLL image.
///   - Upon successful loading, the returned handle can be used with <see cref="lfpGetProcAddress"/> to access exported functions.
///   - To unload the library and free associated resources, use <see cref="lfpFreeLibrary"/>.
/// </remarks>
function wvbLoadLibrary(const AData: Pointer): THandle;

/// <summary>
///   Retrieves the address of an exported function or variable from the specified dynamic-link library.
/// </summary>
/// <param name="AHandle">
///   A handle to the DLL module that contains the function or variable.
/// </param>
/// <param name="AProcName">
///   The name of the function or variable whose address is to be retrieved.
/// </param>
/// <returns>
///   A pointer to the function or variable if successful; otherwise, <c>nil</c>.
/// </returns>
/// <remarks>
///   - This function allows for dynamic retrieval of addresses for exported functions or variables within a loaded DLL.
///   - The <paramref name="AHandle"/> parameter must be a valid handle obtained from <see cref="lfpLoadLibrary"/>.
///   - The <paramref name="AProcName"/> parameter should be a null-terminated string specifying the name of the function or variable.
///   - If the function or variable is not found, the return value is <c>nil</c>.
/// </remarks>
function wvbGetProcAddress(const AHandle: THandle; const AProcName: PAnsiChar): Pointer;

/// <summary>
///   Frees the loaded dynamic-link library (DLL) and releases associated resources.
/// </summary>
/// <param name="AHandle">
///   A handle to the loaded DLL module to be freed.
/// </param>
/// <remarks>
///   - This function unloads a DLL that was previously loaded into the calling process's address space using <see cref="lfpLoadLibrary"/>.
///   - After calling this function, any pointers obtained via <see cref="lfpGetProcAddress"/> for the specified module become invalid and should not be used.
///   - It is important to match each call to <see cref="lfpLoadLibrary"/> with a corresponding call to <see cref="lfpFreeLibrary"/> to prevent memory leaks.
/// </remarks>
procedure wvbFreeLibrary(const AHandle: THandle);

implementation

uses
  WinApi.Windows;

//------------------------------------------------------------------------------
// PE Structure Pointer Type Definitions
// These types define pointers to standard Windows PE file structures.
// They are used to navigate and interpret the PE file data loaded into memory.
// Reference: winnt.h in the Windows SDK.
//------------------------------------------------------------------------------
type
  PIMAGE_NT_HEADERS      = ^IMAGE_NT_HEADERS;
  PIMAGE_FILE_HEADER     = ^IMAGE_FILE_HEADER;
  PIMAGE_OPTIONAL_HEADER = ^IMAGE_OPTIONAL_HEADER;
  PIMAGE_SECTION_HEADER  = ^IMAGE_SECTION_HEADER;
  PIMAGE_DATA_DIRECTORY  = ^IMAGE_DATA_DIRECTORY;

//------------------------------------------------------------------------------
// Large Array Type Definitions (Helper Types)
// These define pointers to large arrays of WORD and Cardinal (DWORD).
// They are primarily used for typecasting when accessing PE structures
// like relocation tables or export tables, where the exact size might not
// be known at compile time, but we need pointer access to elements.
// The large upper bounds are likely placeholders and not indicative of
// expected actual usage size.
//------------------------------------------------------------------------------
type
  TDWORDArray = array[0..9999999] of Cardinal; // Array of 32-bit unsigned integers
  PDWORDArray = ^TDWORDArray;                 // Pointer to TDWORDArray
  TWORDArray  = array[0..99999999] of WORD;     // Array of 16-bit unsigned integers
  PWORDArray  = ^TWORDArray;                   // Pointer to TWORDArray

//------------------------------------------------------------------------------
// IMAGE_DOS_HEADER Structure Definition
// Defines the MS-DOS header structure found at the beginning of every PE file.
// The most important field for PE loading is 'e_lfanew', which points to
// the 'IMAGE_NT_HEADERS'.
//------------------------------------------------------------------------------
type
  IMAGE_DOS_HEADER = packed record     // DOS .EXE header
    e_magic: WORD;                     // Magic number (Should be 'MZ')
    e_cblp: WORD;                      // Bytes on last page of file
    e_cp: WORD;                        // Pages in file
    e_crlc: WORD;                      // Relocations
    e_cparhdr: WORD;                   // Size of header in paragraphs
    e_minalloc: WORD;                  // Minimum extra paragraphs needed
    e_maxalloc: WORD;                  // Maximum extra paragraphs needed
    e_ss: WORD;                        // Initial (relative) SS value
    e_sp: WORD;                        // Initial SP value
    e_csum: WORD;                      // Checksum
    e_ip: WORD;                        // Initial IP value
    e_cs: WORD;                        // Initial (relative) CS value
    e_lfarlc: WORD;                    // File address of relocation table
    e_ovno: WORD;                      // Overlay number
    e_res: array[0..3] of WORD;        // Reserved words
    e_oemid: WORD;                     // OEM identifier (for e_oeminfo)
    e_oeminfo: WORD;                   // OEM information; e_oemid specific
    e_res2: array[0..9] of WORD;       // Reserved words
    e_lfanew: Cardinal;                // File address of new exe header (PE header offset)
  end;

  PIMAGE_DOS_HEADER = ^IMAGE_DOS_HEADER; // Pointer to IMAGE_DOS_HEADER

//------------------------------------------------------------------------------
// IMAGE_BASE_RELOCATION Structure Definition
// Defines the header for a block of base relocations in the PE file.
// Relocations are necessary when the PE file cannot be loaded at its
// preferred base address ('ImageBase' in Optional Header).
// Each block contains a list of offsets (following this structure)
// that need patching relative to the 'VirtualAddress'.
//------------------------------------------------------------------------------
type
  IMAGE_BASE_RELOCATION = packed record
    VirtualAddress: Cardinal; // RVA (Relative Virtual Address) of the page to be fixed up.
    SizeOfBlock: Cardinal;    // Total size of this relocation block, including this header and all relocation entries.
  end;
  PIMAGE_BASE_RELOCATION = ^IMAGE_BASE_RELOCATION; // Pointer to IMAGE_BASE_RELOCATION

//------------------------------------------------------------------------------
// PIMAGE_EXPORT_DIRECTORY Type Definition
// Pointer to the IMAGE_EXPORT_DIRECTORY structure (defined in Winapi.Windows).
// This structure contains information about functions exported by the PE file.
// It's used by `Internal_GetProcAddress` to find exported functions by name.
//------------------------------------------------------------------------------
type
  PIMAGE_EXPORT_DIRECTORY = ^IMAGE_EXPORT_DIRECTORY; // Pointer to IMAGE_EXPORT_DIRECTORY (defined in Windows unit)

//------------------------------------------------------------------------------
// DLLMAIN Function Type Definition
// Defines the signature of the standard entry point function for a DLL.
// This function is called by the loader (or in this case, `Internal_Load` and
// `Internal_Unload`) when the DLL is loaded/unloaded or threads attach/detach.
//------------------------------------------------------------------------------
type
  DLLMAIN = function(hinstDLL: Pointer; fdwReason: Cardinal; lpvReserved: Pointer): Integer; stdcall;
  PDLLMAIN = ^DLLMAIN; // Pointer to a DLLMAIN function

//---------------------------------------------------------------------------
// Internal_CopyMemory
// Purpose: Copies a block of memory from a source location to a destination.
//          This acts as a wrapper around the standard 'memcpy' function,
//          retrieving its address dynamically from ntdll.dll.
// Note: Uses a local variable `MemCpy` which is re-fetched on every call if nil check passes.
//       (Original code checked `nil = @MemCpy` which always evaluates based on the address,
//       not the content, so GetProcAddress is called every time unless optimized out).
// Parameters:
//   Destination: Pointer to the starting address of the destination block.
//   Source     : Pointer to the starting address of the source block.
//   Count      : Number of bytes to copy. Original type was UInt64, kept as NativeUInt for compatibility.
//---------------------------------------------------------------------------
type
  TMemCpy = procedure(ADestination: Pointer; ASource: Pointer; ACount: NativeUInt);cdecl; // Function pointer type for memcpy
  PMemCpy = ^TMemCpy; // Pointer to the function pointer type

procedure Internal_CopyMemory(ADestination: Pointer; ASource: Pointer; ACount: NativeUInt); //was UInt64
var
  LMemCpy: TMemCpy; // Local variable to hold the function pointer
begin
  // Get address of 'memcpy' from ntdll.dll
    LMemCpy := TMemCpy(GetProcAddress(GetModuleHandleA('ntdll.dll'), 'memcpy')); // Fetch address every time for safety based on original code structure

  // Call the obtained memcpy function pointer
  LMemCpy(ADestination, ASource, ACount);
end;

//---------------------------------------------------------------------------
// Internal_ZeroMemory
// Purpose: Fills a block of memory with zeros.
//          Acts as a wrapper around 'RtlZeroMemory' from kernel32.dll,
//          retrieving its address dynamically on each call.
// Note: Similar to Internal_CopyMemory, uses a local variable and fetches
//       the address on each call due to the `if (nil = @ZeroMem)` check logic.
// Parameters:
//   What : Pointer to the starting address of the memory block to zero out.
//   Count: Number of bytes to fill with zero. Original type was UInt64, kept as NativeUInt.
//---------------------------------------------------------------------------
type
  TZeroMem = procedure(AWhat: Pointer; ACount: NativeUInt); stdcall; // Function pointer type for RtlZeroMemory

procedure Internal_ZeroMemory(AWhat: Pointer; ACount: NativeUInt); //was UInt64
var
  LZeroMem: TZeroMem; // Local variable to hold the function pointer
begin
  // Get address of 'RtlZeroMemory' from kernel32.dll
  LZeroMem := TZeroMem(GetProcAddress(GetModuleHandleA('kernel32.dll'), 'RtlZeroMemory'));

  // Call the obtained RtlZeroMemory function pointer
  LZeroMem(AWhat, ACount);
end;

//---------------------------------------------------------------------------
// Pointer Arithmetic Helper Functions
// Provide ways to perform arithmetic operations on pointers using NativeUInt
// for 32/64-bit compatibility.
//---------------------------------------------------------------------------

/// <summary>
/// Adds a Cardinal offset to a Pointer. Casts Pointer and Cardinal to NativeUInt.
/// </summary>
/// <param name="source">The base pointer.</param>
/// <param name="value">The Cardinal offset to add.</param>
/// <returns>A new Pointer offset by the given value.</returns>
function AddToPointer(ASource: Pointer; AValue: Cardinal) : Pointer;overload;
begin
  // Cast pointer to NativeUInt, cast Cardinal value to NativeUInt, add, cast result back to Pointer.
  Result := Pointer(NativeUInt(ASource) + NativeUInt(AValue)); // Int64 cast removed, assuming NativeUInt is sufficient and correct for Delphi pointer math
end;

/// <summary>
/// Adds a NativeUInt offset to a Pointer.
/// </summary>
/// <param name="source">The base pointer.</param>
/// <param name="value">The NativeUInt offset to add.</param>
/// <returns>A new Pointer offset by the given value.</returns>
function AddToPointer(ASource: Pointer; AValue: NativeUInt) : Pointer; overload;
begin
  // Cast pointer to NativeUInt, add NativeUInt value, cast result back to Pointer.
  Result := Pointer(NativeUInt(ASource) + AValue);
end;

/// <summary>
/// Calculates the difference (in bytes) between two Pointers.
/// </summary>
/// <param name="source">The first pointer (minuend).</param>
/// <param name="value">The second pointer (subtrahend).</param>
/// <returns>The difference between the pointers as a NativeUInt.</returns>
function DecPointer(ASource: Pointer; AValue: Pointer) : NativeUInt;
begin
  // Cast both pointers to NativeUInt and subtract.
  Result := NativeUInt(ASource) - NativeUInt(AValue);
end;

/// <summary>
/// Subtracts a NativeUInt offset from a Pointer and returns the result as NativeUInt.
/// Calculates `Pointer_Address - Offset`.
/// </summary>
/// <param name="source">The base pointer.</param>
/// <param name="value">The NativeUInt offset to subtract.</param>
/// <returns>The resulting memory address as a NativeUInt.</returns>
function DecPointerInt(ASource: Pointer; AValue: NativeUInt) : NativeUInt;
begin
  // Cast pointer to NativeUInt and subtract the NativeUInt offset.
  Result := NativeUInt(ASource) - NativeUInt(AValue);
end;

/// <summary>
/// Returns the minimum of two Integer values.
/// </summary>
function min(a: Integer; b: Integer): Integer;
begin
  if (a<b) then
    Result := a
  else
    Result := b;
end;

//------------------------------------------------------------------------------
// Internal_Load
// Purpose: Loads a PE image (DLL/EXE) from the memory buffer pointed to by pData.
//          Performs memory allocation, section mapping, base relocations,
//          import resolution, and calls the entry point.
// Parameters:
//   pData: A pointer to the raw PE file data in memory.
// Returns:
//   A Pointer to the base address where the PE image has been loaded in the
//   current process's virtual memory. Returns nil on failure (error handling is minimal).
//------------------------------------------------------------------------------
function Internal_Load(AData: Pointer) : Pointer;
var
  LPtr: Pointer;                              // General purpose pointer, used sequentially for DOS Header -> NT Headers -> Sections
  LImageNTHeaders: PIMAGE_NT_HEADERS;         // Pointer to the NT Headers structure
  LSectionIndex: Integer;                     // Loop counter for processing sections
  LImageBaseDelta: Size_t;                    // Difference between actual load address and preferred ImageBase. Size_t is platform dependent (usually NativeUInt).
  LRelocationInfoSize: UInt;                  // Size of the base relocation data directory (UInt is typically Cardinal/DWORD)
  LImageBaseRelocations,                      // Pointer to the start of the relocation data
  LReloc: PIMAGE_BASE_RELOCATION;             // Pointer to the current relocation block header being processed
  LImports,                                   // Pointer to the start of the import directory table
  LImport: PIMAGE_IMPORT_DESCRIPTOR;          // Pointer to the current import descriptor being processed
  LDllMain: DLLMAIN;                          // Pointer to the module's entry point function (if any)
  LImageBase: Pointer;                        // Base address of the allocated memory for the loaded image

  LImageSectionHeader: PIMAGE_SECTION_HEADER; // Pointer to the current section header being processed
  LVirtualSectionSize: Integer;               // The virtual size of the current section
  LRawSectionSize: Integer;                   // The size of the raw data for the current section
  LSectionBase: Pointer;                      // Pointer to the base address of the current section in allocated memory

  LRelocCount: Integer;                       // Number of relocations in the current block
  LRelocInfo: PWORD;                          // Pointer to the current relocation entry (WORD containing type/offset)
  LRelocIndex: Integer;                       // Loop counter for relocations within a block

  LMagic: PNativeUInt;                        // Pointer to the memory location that needs relocation patching (Original name 'magic')

  LLibName: LPSTR;                            // Name of the DLL to import functions from (LPSTR is PAnsiChar)
  LLib: HMODULE;                              // Handle to the loaded dependency DLL
  LPRVAImport: PNativeUInt;                   // Pointer to the IAT/ILT entry to read/patch. (Original comment: UInt, not PNativeUInt!) - Code uses PNativeUInt.
  LFunctionName: LPSTR;                       // Name or Ordinal (as LPSTR) of the function to import
begin
  // Initialize pPtr with the start of the raw PE data
  LPtr := AData;

  // 1. Locate NT Headers
  // Use e_lfanew field from DOS header to find the offset to NT headers.
  // Note the Int64 cast, which might be unnecessary if pData is properly aligned and offsets fit NativeInt.
  LPtr := Pointer(Int64(LPtr) + Int64(PIMAGE_DOS_HEADER(LPtr).e_lfanew));
  LImageNTHeaders := PIMAGE_NT_HEADERS(LPtr); // Cast the calculated address to NT Headers pointer

  // 2. Allocate Memory for the Image
  // Allocate virtual memory block. Attempt at preferred ImageBase (nil),
  // fallback to any address if needed (handled by OS).
  // Size from OptionalHeader.SizeOfImage. Permissions EXECUTE_READWRITE.
  LImageBase := VirtualAlloc(nil, LImageNTHeaders^.OptionalHeader.SizeOfImage, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
  // NOTE: Original code doesn't check if VirtualAlloc returns nil (failure). Add check if robustness needed.

  // 3. Copy PE Headers
  // Copy DOS header, NT headers, section headers from raw data (pData) to allocated memory (pImageBase).
  Internal_CopyMemory(LImageBase, AData, LImageNTHeaders^.OptionalHeader.SizeOfHeaders);

  // 4. Map Sections into Memory
  // Advance pPtr to point to the first section header (after NT Headers structure).
  // SizeOfOptionalHeader determines the offset.
  LPtr := AddToPointer(LPtr,sizeof(LImageNTHeaders.Signature) + sizeof(LImageNTHeaders.FileHeader) + LImageNTHeaders.FileHeader.SizeOfOptionalHeader);

  // Loop through each section header
  for LSectionIndex := 0 to LImageNTHeaders.FileHeader.NumberOfSections-1 do
  begin
    // Calculate address of the current section header
    LImageSectionHeader := PIMAGE_SECTION_HEADER(AddToPointer(LPtr,LSectionIndex*sizeof(IMAGE_SECTION_HEADER)));

    // Get virtual size and raw data size for the section
    // Original comment: PhysicalAddress new code - referring to Misc field union? Using VirtualSize.
    LVirtualSectionSize := LImageSectionHeader.Misc.VirtualSize;
    LRawSectionSize := LImageSectionHeader.SizeOfRawData;

    // Calculate the target address for this section within the allocated memory block
    LSectionBase := AddToPointer(LImageBase,LImageSectionHeader.VirtualAddress);

    // Zero out the memory allocated for the section's virtual size
    Internal_ZeroMemory(LSectionBase, LVirtualSectionSize);

    // Copy the raw section data from input buffer (pData) to the allocated memory (pSectionBase).
    // Use 'min' to avoid writing past virtual size or reading past raw data size.
    Internal_CopyMemory(LSectionBase,
      AddToPointer(AData,LImageSectionHeader.PointerToRawData),
      min(LVirtualSectionSize, LRawSectionSize));
  end; // End section mapping loop

  // 5. Process Base Relocations
  // Calculate the difference (delta) between actual load address (pImageBase) and preferred (ImageBase).
  // Note: Using DecPointerInt returns NativeUInt, cast to Size_t (platform-dependent int).
  LImageBaseDelta := DecPointerInt(LImageBase,LImageNTHeaders.OptionalHeader.ImageBase);

  // Get the size and starting address of the relocation data.
  LRelocationInfoSize := LImageNTHeaders.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_BASERELOC].Size;
  LImageBaseRelocations := PIMAGE_BASE_RELOCATION(AddToPointer(LImageBase,
    LImageNTHeaders.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_BASERELOC].VirtualAddress));

  LReloc := LImageBaseRelocations; // Initialize pointer to the first relocation block

  // Iterate through relocation blocks as long as we are within the relocation data size.
  while DecPointer(LReloc,LImageBaseRelocations) < LRelocationInfoSize do
  begin
    // Calculate number of relocation entries (WORDs) in this block.
    LRelocCount := (LReloc.SizeOfBlock - sizeof(IMAGE_BASE_RELOCATION)) Div sizeof(WORD);

    // Get pointer to the first relocation entry (WORD TypeOffset) following the block header.
    LRelocInfo := PWORD(AddToPointer(LReloc,sizeof(IMAGE_BASE_RELOCATION)));

    // Iterate through each relocation entry in the current block.
    for LRelocIndex := 0 to LRelocCount-1 do
    begin
      // Check the relocation type (high 4 bits of the WORD).
      // If type is not 0 (IMAGE_REL_BASED_ABSOLUTE), apply relocation.
      // Original code checks `(pwRelocInfo^ and $f000) <> 0`. This handles common types
      // like HIGHLOW (3) and DIR64 (A) but isn't strictly type-specific.
      if (LRelocInfo^ and $f000) <> 0 then
      begin
        // Calculate the address in memory that needs patching.
        // Base + Block RVA + Offset (low 12 bits of WORD).
        LMagic := PNativeUInt(AddToPointer(LImageBase,LReloc.VirtualAddress+(LRelocInfo^ and $0fff)));

        // Apply the delta: Add the difference between actual and preferred base addresses.
        // Direct pointer arithmetic on the target location.
        LMagic^ := NativeUInt(LMagic^ + LImageBaseDelta);
        // Original C++ comment equivalent: *(char* *)((char*)pImageBase + pReloc->VirtualAddress + (pwRelocInfo^ and $0fff)) += intImageBaseDelta;
      end;

      Inc(LRelocInfo); // Move to the next relocation entry (WORD)
    end; // End loop for entries in current block

    // Move pReloc pointer to the next relocation block header.
    // pwRelocInfo now points just past the last entry of the current block.
    LReloc := PIMAGE_BASE_RELOCATION(LRelocInfo);
  end; // End loop for relocation blocks

  // 6. Process Import Table
  // Locate the Import Directory Table.
  LImports := PIMAGE_IMPORT_DESCRIPTOR(AddToPointer(LImageBase,
    LImageNTHeaders.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT].VirtualAddress));

  LImport := LImports; // Initialize pointer to the first import descriptor

  // Iterate through import descriptors until the Name RVA is 0 (null terminator).
  while 0 <> LImport.Name do
  begin
    // Get the name of the DLL to import from the descriptor's Name RVA.
    LLibName := LPSTR(AddToPointer(LImageBase,LImport.Name));

    // Load the required DLL.
    LLib := LoadLibraryA(LLibName);
    // NOTE: Original code doesn't check if hLib is 0 (load failure). Add check if needed.

    // Determine the table to iterate for function lookups (ILT or IAT).
    // Original logic: Uses TimeDateStamp == 0 to choose FirstThunk (IAT), otherwise Characteristics (OriginalFirstThunk/ILT).
    // Standard practice often uses OriginalFirstThunk (ILT) for lookup and patches FirstThunk (IAT).
    // Documenting the code as written.
    if 0 = LImport.TimeDateStamp then
      // If TimeDateStamp is 0, iterate the IAT (FirstThunk).
      LPRVAImport := AddToPointer(LImageBase,LImport.FirstThunk)
    else
      // Otherwise, iterate the ILT (OriginalFirstThunk, stored in Characteristics field pre-binding).
      LPRVAImport := AddToPointer(LImageBase,LImport.Characteristics); //new code comment implies this field choice is newer

    // Iterate through the ILT/IAT entries (list is null-terminated).
    // lpPRVA_Import points to the RVA/Ordinal entry.
    while LPRVAImport^ <> 0 do
    begin
      // Check if import is by Ordinal (high bit set).
      // Note: Uses IMAGE_ORDINAL_FLAG32 specifically. Should ideally use IMAGE_ORDINAL_FLAG.
      if (PDWORD(LPRVAImport)^ and IMAGE_ORDINAL_FLAG32) <> 0 then
      begin
        // Import by Ordinal: Ordinal is in the low 16 bits.
        // Cast ordinal directly to LPSTR for GetProcAddress.
        LFunctionName := LPSTR(PDWORD(LPRVAImport)^ and $ffff);
      end
      else
      begin
        // Import by Name: Entry is an RVA to IMAGE_IMPORT_BY_NAME structure.
        // Get address of the structure, then the Name field within it.
        // Note: `PUInt(lpPRVA_Import)^` dereferences the pointer to get the RVA.
        LFunctionName := LPSTR(@PIMAGE_IMPORT_BY_NAME(AddToPointer(LImageBase, PUInt(LPRVAImport)^)).Name[0]);
      end;

      // Get the actual address of the imported function using GetProcAddress.
      // Note: GetProcAddress returns FARPROC which needs casting. Result stored back into IAT entry.
      // It seems lpPRVA_Import itself is being modified, which implies it points to the IAT entry directly.
      // This suggests the `if 0 = pImport.TimeDateStamp` logic might be reversed or the variable use is subtle.
      // Assuming lpPRVA_Import points to the IAT entry that needs patching.
      LPRVAImport^ := NativeUInt(GetProcAddress(LLib, LFunctionName));

      Inc(LPRVAImport); // Move to the next IAT entry.
    end; // End loop for functions in current DLL

    Inc(LImport); // Move to the next import descriptor.
  end; // End loop for imported DLLs

  // 7. Flush Instruction Cache
  // Ensures CPU sees the newly written/patched code in memory.
  FlushInstructionCache(GetCurrentProcess(), LImageBase, LImageNTHeaders.OptionalHeader.SizeOfImage);

  // 8. Call Entry Point (DllMain)
  // Check if an entry point exists (AddressOfEntryPoint RVA is non-zero).
  if 0 <> LImageNTHeaders.OptionalHeader.AddressOfEntryPoint then
  begin
    // Calculate the absolute address of the entry point function.
    LDllMain := DLLMAIN(AddToPointer(LImageBase,LImageNTHeaders.OptionalHeader.AddressOfEntryPoint));

    // Check if the function pointer seems valid (check address, not content).
    // The `nil <> @pDllMain` check compares the *address* of the local variable `pDllMain`.
    // This is always true if pDllMain is a stack variable. Assume intent was `Assigned(pDllMain)`.
    if nil <> @LDllMain then
    // if Assigned(pDllMain) then // Use this check instead if the original is a typo
    begin
      // Call DllMain with DLL_PROCESS_ATTACH.
      LDllMain(Pointer(LImageBase), DLL_PROCESS_ATTACH, nil);
      // Call DllMain immediately with DLL_THREAD_ATTACH.
      // NOTE: This is generally incorrect usage. Thread attach/detach calls
      // should typically be made by the OS loader for new/exiting threads *after*
      // process attach is complete. Documenting the code as written.
      LDllMain(Pointer(LImageBase), DLL_THREAD_ATTACH, nil);
    end;
  end;

  // 9. Success: Return the base address of the loaded image.
  Result := Pointer(LImageBase);
end;

//------------------------------------------------------------------------------
// Internal_GetProcAddress
// Purpose: Finds the address of an exported function within a module that was
//          previously loaded into memory using `Internal_Load`. Mimics standard
//          `GetProcAddress` but operates on the manually loaded module data.
// Parameters:
//   hModule   : The base pointer of the manually loaded module (returned by `Internal_Load`).
//   lpProcName: The null-terminated ANSI string containing the name of the function to find.
//               Note: Does not support lookup by ordinal via this parameter.
// Returns:
//   A Pointer to the exported function's entry point. Returns nil if not found or on error.
//------------------------------------------------------------------------------
function Internal_GetProcAddress(hModule: Pointer; lpProcName: PAnsiChar) : Pointer;
var
  LImageNTHeaders: PIMAGE_NT_HEADERS;       // Pointer to the NT Headers
  LExports: PIMAGE_EXPORT_DIRECTORY;        // Pointer to the Export Directory structure
  LExportedSymbolIndex: Cardinal;           // Loop counter for exported names
  LPtr: Pointer;                            // Pointer used to navigate from module base to NT Headers
  LVirtualAddressOfName: Cardinal;          // RVA of the current function name string being checked
  LName: PAnsiChar;                         // Pointer to the current function name string in memory
  LIndex: WORD;                             // Index into the AddressOfFunctions table (obtained from NameOrdinals table)
  LVirtualAddressOfAddressOfProc: Cardinal; // RVA of the found function's entry point
begin
  Result := nil; // Initialize result to nil (not found)

  // Basic validation of input parameters
  if nil <> hModule then // Check if module handle is valid
  begin
    // 1. Locate NT Headers
    LPtr := hModule; // Start at module base
    // Calculate address of NT Headers using e_lfanew from DOS header. Int64 cast used.
    LPtr := Pointer(Int64(LPtr) + Int64(PIMAGE_DOS_HEADER(LPtr).e_lfanew));
    LImageNTHeaders := PIMAGE_NT_HEADERS(LPtr);
    // NOTE: Add signature checks (DOS, NT) for robustness if needed.

    // 2. Locate Export Directory Table
    // Get pointer to the export directory using the RVA from the data directory.
    LExports := PIMAGE_EXPORT_DIRECTORY(AddToPointer(hModule,
      LImageNTHeaders.OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT].VirtualAddress));
    // NOTE: Add check if VirtualAddress is 0 (no export table) if needed.

    // 3. Search for the function by name
    // Iterate through the array of exported function names (AddressOfNames).
    for LExportedSymbolIndex := 0 to LExports.NumberOfNames-1 do
    begin
      // Get the RVA of the name string for the current index. Uses PDWORDArray cast for indexing.
      LVirtualAddressOfName := PDWORDArray(AddToPointer(hModule,LExports.AddressOfNames))[LExportedSymbolIndex];

      // Calculate the actual address of the name string.
      LName := LPSTR(AddToPointer(hModule,LVirtualAddressOfName));

      // Compare the current exported name with the requested name (case-sensitive).
      // Use lstrcmpiA for case-insensitive if that's desired (standard GetProcAddress is case-sensitive technically, but LoadLibrary often handles case).
      if lstrcmpA(LName, lpProcName) = 0 then // Function name found!
      begin
        // 4. Get the function's address RVA
        // Find the ordinal associated with this name from the AddressOfNameOrdinals table.
        // This ordinal acts as the INDEX into the AddressOfFunctions table.
        // Uses PWORDArray cast for indexing.
        LIndex := PWORDArray(AddToPointer(hModule,LExports.AddressOfNameOrdinals))[LExportedSymbolIndex]; //new code comment implies this was changed

        // Use the index (wIndex) to get the RVA of the function's entry point from AddressOfFunctions table.
        // Uses PDWORDArray cast for indexing.
        LVirtualAddressOfAddressOfProc := PDWORDArray(AddToPointer(hModule,LExports.AddressOfFunctions))[LIndex];

        // 5. Calculate the absolute address
        // Add the function's RVA to the module's base address.
        Result := AddToPointer(hModule,LVirtualAddressOfAddressOfProc);
        Exit; // Function found, exit the loop and function.
        // NOTE: This does not handle forwarded exports (where dwVirtualAddressOfAddressOfProc points inside the export dir).
      end;
    end;
  end;

  // Function name not found or hModule was nil. Result remains nil.
end;

//------------------------------------------------------------------------------
// Internal_Unload
// Purpose: Unloads a module previously loaded by `Internal_Load`. Calls the
//          module's entry point with detach reasons and frees allocated memory.
// Parameters:
//   hModule: The base pointer of the manually loaded module (returned by `Internal_Load`).
//------------------------------------------------------------------------------
procedure Internal_Unload(hModule: Pointer);
var
  LImageBase: Pointer;                // Variable to hold the module base address (same as hModule parameter)
  LPtr: Pointer;                      // Pointer used to navigate from module base to NT Headers
  LImageNTHeaders: PIMAGE_NT_HEADERS; // Pointer to the NT Headers
  LDllMain: DLLMAIN;                  // Pointer to the module's entry point function
begin
  // Check if the module handle is valid
  if nil <> hModule then
  begin
    LImageBase := hModule; // Store the base address

    // 1. Locate NT Headers to find the entry point
    LPtr := Pointer(hModule); // Start at module base

    // Calculate address of NT Headers using e_lfanew. Int64 cast used.
    LPtr := Pointer(Int64(LPtr) + Int64(PIMAGE_DOS_HEADER(LPtr).e_lfanew));
    LImageNTHeaders := PIMAGE_NT_HEADERS(LPtr);

    // NOTE: Add signature checks for robustness if needed.
    // 2. Calculate Entry Point Address
    LDllMain := DLLMAIN(AddToPointer(LImageBase,LImageNTHeaders.OptionalHeader.AddressOfEntryPoint));

    // Check if the entry point address seems valid (using address-of check).
    // `nil <> @pDllMain` check compares the *address* of the local variable.
    // Assume intent was `Assigned(pDllMain)` and AddressOfEntryPoint > 0.
    if nil <> @LDllMain then
    // if (pImageNTHeaders.OptionalHeader.AddressOfEntryPoint <> 0) and Assigned(pDllMain) then // Safer check
    begin
      // 3. Call Entry Point with Detach Reasons
      // Call DllMain for thread detach first.
      // NOTE: Manually calling DLL_THREAD_DETACH is complex and may not behave as expected compared to OS handling.
      LDllMain(LImageBase, DLL_THREAD_DETACH, nil);

      // Call DllMain for process detach.
      LDllMain(LImageBase, DLL_PROCESS_DETACH, nil);
    end;

    // 4. Free the Allocated Memory
    // Release the virtual memory block allocated during Internal_Load.
    // Size must be 0, dwFreeType must be MEM_RELEASE.
    VirtualFree(hModule, 0, MEM_RELEASE);
  end;
end;

function  wvbLoadLibrary(const AData: Pointer): THandle;
begin
  Result := THandle(Internal_Load(AData));
end;

function  wvbGetProcAddress(const AHandle: THandle; const AProcName: PAnsiChar): Pointer;
begin
  Result := Internal_GetProcAddress(Pointer(AHandle), AProcName);
end;

procedure wvbFreeLibrary(const AHandle: THandle);
begin
  Internal_Unload(Pointer(AHandle));
end;

end.
