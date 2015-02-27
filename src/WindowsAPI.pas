unit WindowsAPI;

interface
uses
	Windows,
	ShellAPI;

type
	TWinVersion = (
		wvUnknown,
		wvWinXPHome,
		wvWinXPPro,
		wvWinXPMCE,
		wvWinXPTab,
		wvWinVistaHB,
		wvWinVistaHP,
		wvWinVistaBS,
		wvWinVistaUT,
		wvWinSevenHP,
		wvWinSevenPR,
		wvWinSevenUT
	);

type
	POSVersionInfoEx = ^TOSVersionInfoEx;
	TOSVersionInfoEx = record
		dwOSVersionInfoSize: DWORD;
		dwMajorVersion: DWORD;
		dwMinorVersion: DWORD;
		dwBuildNumber: DWORD;
		dwPlatformId: DWORD;
		szCSDVersion: array [0..127] of Char;     // Maintenance string for PSS usage
		wServicePackMajor: Word;
		wServicePackMinor: Word;
		wSuiteMask: Word;
		wProductType: Byte;
		wReserved: Byte;
	end;

	function GetVersionEx(lpVersionInformation: POSVersionInfoEx): BOOL; stdcall;

	function GetProductInfo(
		dwOSMajorVersion,
		dwOSMinorVersion,
		dwSpMajorVersion,
		dwSpMinorVersion : DWord;
		var pdwReturnedProductType: Dword): BOOL;

	function IsSafeMode : boolean;
	function GetWindowsVersion : TWinVersion;

	procedure RunAndWait(ExecuteFile : string; ParamString : string);

	function scShellDeleteFile(StrFile : String; BlnSilent : Boolean = False;
		BlnConfirmation : Boolean = True; BlnUndo : Boolean = True) : Boolean;
	function scShellCopyFile(StrFrom, StrTo : string;
		BlnSilent : Boolean = False) : Boolean;
	function scShellMoveFile(StrFrom, StrTo : string;
		BlnSilent : Boolean = False) : Boolean;

	function SetupSFCException(FileNameAndPath : WideString) : boolean;
	function IsSFCProtected(FileNameAndPath : WideString) : boolean;


implementation
Uses Dialogs, sysutils, Forms;

function GetVersionEx; external kernel32 name 'GetVersionExA';
function GetProductInfo; external kernel32 name 'GetProductInfo';

function LoadVistaDLL : Pointer;
begin
  {$IFDEF WARNDIRS}{$WARN UNSAFE_TYPE OFF}{$ENDIF}
  Result := GetProcAddress(GetModuleHandle('kernel32.dll'), 'GetProductInfo');
	{$IFDEF WARNDIRS}{$WARN UNSAFE_TYPE ON}{$ENDIF}
end;


function GetWindowsVersion : TWinVersion;

	function LoadKernelFunc(const FuncName: string): Pointer;
		{Loads a function from the OS kernel.
			@param FuncName [in] Name of required function.
			@return Pointer to function or nil if function not found in kernel.
		}
	const
		cKernel = 'kernel32.dll'; // kernel DLL
	begin
		{$IFDEF WARNDIRS}{$WARN UNSAFE_TYPE OFF}{$ENDIF}
		Result := GetProcAddress(GetModuleHandle(cKernel), PChar(FuncName));
		{$IFDEF WARNDIRS}{$WARN UNSAFE_TYPE ON}{$ENDIF}
	end;

type
	// Function type of the GetProductInfo API function
	TGetProductInfo = function(OSMajor, OSMinor, SPMajor, SPMinor: DWORD;
		out ProductType: DWORD): BOOL; stdcall;

var
	osVerInfo: TOSVersionInfoEx;
	majorVersion, minorVersion: Integer;
	GetProductInfo: TGetProductInfo;  // pointer to GetProductInfo API function
	Win32ProductInfo : cardinal;

begin
	Result := wvUnknown;
	Win32ProductInfo := 0;
	osVerInfo.dwOSVersionInfoSize := SizeOf(TOSVersionInfoEx) ;
	if GetVersionEx(@osVerInfo) then
	begin
		minorVersion := osVerInfo.dwMinorVersion;
		majorVersion := osVerInfo.dwMajorVersion;
		case osVerInfo.dwPlatformId of
			VER_PLATFORM_WIN32_NT:
			begin
      	//XP
				if (majorVersion = 5) and (minorVersion = 1) then
				begin
					if (osVerInfo.wSuiteMask AND 512) <> 0 then
					begin
						Result := wvWinXPHome;
					end
					else begin
						if Boolean(GetSystemMetrics(87)) then Result := wvWinXPMCE
						else if Boolean(GetSystemMetrics(86)) then Result := wvWinXPTab
						else Result := wvWinXPPro;
					end;
				end
				//Vista / 7
				else if (majorVersion = 6) then
				begin
					ShowMessage('Minor: ' + IntToStr(minorVersion));
					GetProductInfo := LoadKernelFunc('GetProductInfo');
					if Assigned(GetProductInfo) then
					begin
						if not GetProductInfo(
							osVerInfo.dwMajorVersion, osVerInfo.dwMinorVersion,
							osVerInfo.wServicePackMajor, osVerInfo.wServicePackMinor,
							Win32ProductInfo
						) then Win32ProductInfo := 0;
					end;
					//Vista is ver 6.0.etc
					if minorVersion = 0 then
					begin
						case Win32ProductInfo of
							$02: Result := wvWinVistaHB;
							$03: Result := wvWinVistaHP;
							$06: Result := wvWinVistaBS;
							$01: Result := wvWinVistaUT;
						end;
					end else
          //7 is ver 6.1.etc
					if minorVersion = 1 then
					begin
						case Win32ProductInfo of
							$03: Result := wvWinSevenHP;
							$06: Result := wvWinSevenPR;
							$01: Result := wvWinSevenUT;
						end;
					end;
				end;
			end;
		end;
	end;
end;

function IsSafeMode : boolean;
begin
	Result := Bool(GetSystemMetrics(SM_CLEANBOOT));
end;

procedure RunAndWait(ExecuteFile : string; ParamString : string);
var
	SEInfo: TShellExecuteInfo;
	ExitCode: DWORD;
//	StartInString: string;
begin
	FillChar(SEInfo, SizeOf(SEInfo), 0) ;
	SEInfo.cbSize := SizeOf(TShellExecuteInfo) ;
	with SEInfo do begin
		fMask := SEE_MASK_NOCLOSEPROCESS;
		Wnd := GetDesktopWindow;
		lpFile := PChar(ExecuteFile) ;

		{
		ParamString can contain the
		application parameters.
		}
		lpParameters := PChar(ParamString) ;
		{
		StartInString specifies the
		name of the working directory.
		If ommited, the current directory is used.
		}
		// lpDirectory := PChar(StartInString) ;

		nShow := SW_SHOWNORMAL;
	end;
	if ShellExecuteEx(@SEInfo) then begin
		repeat
			Application.ProcessMessages;
			GetExitCodeProcess(SEInfo.hProcess, ExitCode) ;
		until (ExitCode <> STILL_ACTIVE) or Application.Terminated;
		//ShowMessage('Completed') ;
	end else ShowMessage('Error starting '+ ExecuteFile + ' ' + ParamString) ;
end;

// ----------------------------------------------------------------
// Delete file or folder
// ----------------------------------------------------------------
function scShellDeleteFile(StrFile : String; BlnSilent : Boolean = False;
	BlnConfirmation : Boolean = True; BlnUndo : Boolean = True) : Boolean;
var
	F : TShFileOpStruct;
begin
	F.Wnd:=GetDesktopWindow;
	F.wFunc:=FO_DELETE;
	F.pFrom:=PChar(StrFile+#0);
	F.pTo:=nil;
	F.fFlags := FOF_NOCONFIRMATION or FOF_SILENT;
	if BlnUndo then
		F.fFlags := FOF_ALLOWUNDO;
	if not BlnConfirmation then
		F.fFlags := FOF_NOCONFIRMATION;
	if BlnSilent then
		F.fFlags := F.fFlags or FOF_SILENT;
	if ShFileOperation(F) <> 0 then
		result:=False
	else
		result:=True;
end;


// ----------------------------------------------------------------
// Copy files
// ----------------------------------------------------------------
function scShellCopyFile(StrFrom, StrTo : string;
	BlnSilent : Boolean = False) : Boolean;
var
	F : TShFileOpStruct;
begin
	F.Wnd:=GetDesktopWindow;
	F.wFunc:=FO_COPY;
	F.pFrom:=PChar(StrFrom+#0);
	F.pTo:=PChar(StrTo+#0);
	F.fFlags := FOF_ALLOWUNDO or FOF_RENAMEONCOLLISION or FOF_NOCONFIRMATION;
	if BlnSilent then
		F.fFlags := F.fFlags or FOF_SILENT;
	if ShFileOperation(F) <> 0 then
		result:=False
	else
		result:=True;
end;

// ----------------------------------------------------------------
// Move files
// ----------------------------------------------------------------
function scShellMoveFile(StrFrom, StrTo : string;
	BlnSilent : Boolean = False) : Boolean;
var
	F : TShFileOpStruct;
begin
	F.Wnd:=GetDesktopWindow;
	F.wFunc:=FO_MOVE;
	F.pFrom:=PChar(StrFrom+#0);
	F.pTo:=PChar(StrTo+#0);
	F.fFlags := FOF_ALLOWUNDO or FOF_RENAMEONCOLLISION or FOF_NOCONFIRMATION;
	if BlnSilent then
		F.fFlags := F.fFlags or FOF_SILENT;
	if ShFileOperation(F) <> 0 then
		result:=False
	else
		result:=True;
end;



//SFC!!!!!!
function SetupSFCException(FileNameAndPath : WideString) : boolean;
type
	TSfcFileException = function(dwUnknown0:Integer;pwszFile:PWideChar;dwUnknown1:Integer):Integer; stdcall;
var
	hSfc: HMODULE;
	SfcFileException : TSfcFileException;
begin
	Result := False;
	if FileNameAndPath = '' then exit;
	hSfc := LoadLibrary('sfc_os.dll');
	SfcFileException := GetProcAddress(hSfc,MAKEINTRESOURCE(5));
	if SfcFileException(0,PWideChar(FileNameAndPath ),-1) = 0 then Result := true;
	FreeLibrary(hSfc);
end;


Function SfcIsFileProtected(RpcHandle: THandle; ProtFileName: PWideChar): LongBool; StdCall;
		External 'Sfc.dll' name 'SfcIsFileProtected';

function IsSFCProtected(FileNameAndPath : WideString) : boolean;
begin
	Result := SfcIsFileProtected(0,PWideChar(FileNameAndPath));
end;

end.
